# -*- coding: utf-8 -*-
"""
wc_chat_parser.py
=================
Wandelt einen ROHEN WhatsApp-Wochenchallenge-Chat in das saubere Zeilenformat um,
das auto_wochenchallenge.parse_email_body() bereits versteht:

    DD.MM Kanonischer-Name
    z.B.  06.06 Felix E.
          07.06 Erik E.

Damit bleibt die komplette restliche Pipeline (Parsen, Gruppen, Excel/PDF, Upload)
UNVERAENDERT - minimales Bruchrisiko. Die KI macht nur die schmutzige Deutungsarbeit:
  * WhatsApp-Anzeigenamen / Spitznamen / Emojis  -> echter Turner (Roster)
  * relative Tage ("gestern", "heute", "Montag")  -> echtes Datum (aus dem
    Zeitstempel der jeweiligen Nachricht im Export)
  * Geschwaetz/irrelevante Nachrichten            -> ignorieren
  * doppelte Meldungen                            -> nur einmal

KI-Backend: GitHub Models (kostenlos fuer jeden GitHub-Account, kein extra Account,
keine Kreditkarte). Laeuft direkt im bestehenden GitHub-Actions-Runner.
Token: GH_MODELS_TOKEN  oder  GITHUB_TOKEN  (Workflow-Permission: models: read).
"""

import json, os, re, urllib.request, urllib.error

GH_MODELS_ENDPOINT = os.environ.get(
    "GH_MODELS_ENDPOINT", "https://models.github.ai/inference/chat/completions")
GH_MODELS_MODEL = os.environ.get("GH_MODELS_MODEL", "openai/gpt-4o-mini")


# ──────────────────────────────────────────────────────────────────────────
#  Roster-Text fuer den KI-Prompt aus den vorhandenen Strukturen aufbauen
# ──────────────────────────────────────────────────────────────────────────
def _aliases_for(canon, name_map):
    """Alle bekannten Schreibweisen (Keys aus NAME_MAP) fuer einen kanon. Namen."""
    al = sorted({k for k, v in name_map.items() if v[1] == canon})
    return al


def build_roster_text(name_map, gruppen_template):
    lines = []
    for grp, names in gruppen_template.items():
        lines.append(f"Gruppe {grp}:")
        for canon in names:
            al = [a for a in _aliases_for(canon, name_map) if a != canon.lower()]
            extra = f"   (auch: {', '.join(al)})" if al else ""
            lines.append(f"  - {canon}{extra}")
    return "\n".join(lines)


SYSTEM_PROMPT = """Du bist ein praeziser Auswerter fuer eine Kinder-Turn-Wochenchallenge.
Du bekommst einen rohen WhatsApp-Chat-Export. Die Kinder melden darin, an welchen
Tagen sie trainiert/geuebt haben. Der Chat ist chaotisch: Emojis, Spitznamen,
Tippfehler, relative Tagesangaben ("gestern", "heute", "am Montag"), Geschwaetz.

Deine Aufgabe: Erzeuge eine SAUBERE Liste aller gueltigen Trainings-Meldungen.

REGELN:
1. Ordne jede Meldung einem Turner aus dem ROSTER zu (echte Person). Anzeigenamen,
   Spitznamen, Vornamen, Nachname-Initialen und Emojis muessen auf den ROSTER-Namen
   gemappt werden. Gib IMMER exakt den kanonischen Roster-Namen aus (z.B. "Felix E.",
   "Ben F."), nie den Spitznamen.
2. Wer eine Nachricht schreibt, meldet normalerweise SICH SELBST - ausser der Text
   nennt klar eine andere Person.
3. DATUM: Jede WhatsApp-Zeile beginnt mit einem Zeitstempel wie
   "[06.06.26, 14:23] Name: ...". Nutze diesen Zeitstempel als Bezug:
     - "gestern"  -> Tag vor dem Nachrichtendatum
     - "heute" / keine Angabe -> das Nachrichtendatum
     - "am Montag"/"Montag" -> der betreffende Wochentag in der Woche der Nachricht
   Gib das aufgeloeste Datum als "DD.MM" aus (Tag.Monat, zweistellig).
4. Ignoriere alles, was keine Trainings-Meldung ist (Smalltalk, Fragen, Emojis allein).
5. Dieselbe Person am selben Tag = nur EIN Eintrag (keine Duplikate).
6. Kannst du eine Meldung keinem Roster-Namen sicher zuordnen, LASS SIE WEG und
   trage sie unter "unsure" mit kurzer Begruendung ein.

Antworte AUSSCHLIESSLICH mit JSON in genau dieser Form, ohne Markdown, ohne Text:
{"entries":[{"date":"DD.MM","name":"Kanonischer Name"}, ...],
 "unsure":["kurze Begruendung", ...]}"""


def _call_github_models(messages, token, model=None, timeout=60):
    body = json.dumps({
        "model": model or GH_MODELS_MODEL,
        "messages": messages,
        "temperature": 0,
        "response_format": {"type": "json_object"},
    }).encode("utf-8")
    req = urllib.request.Request(GH_MODELS_ENDPOINT, data=body, method="POST")
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Content-Type", "application/json")
    req.add_header("Accept", "application/json")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        data = json.loads(r.read().decode("utf-8"))
    return data["choices"][0]["message"]["content"]


def _canon_set(name_map):
    """lowercase-key -> kanonischer Name, fuer Validierung der KI-Ausgabe."""
    return {k: v[1] for k, v in name_map.items()}


def _validate_name(raw_name, lookup):
    """Bringt einen KI-Namen auf einen gueltigen NAME_MAP-Key zurueck (oder None)."""
    n = re.sub(r"[()]", " ", str(raw_name).lower())
    n = re.sub(r"\s+", " ", n).strip()
    if n in lookup:
        return lookup[n]
    # ohne Punkt versuchen ("felix e" statt "felix e.")
    n2 = n.rstrip(".").strip()
    if n2 in lookup:
        return lookup[n2]
    return None


def chat_to_clean_body(raw_chat, name_map, gruppen_template,
                       token=None, model=None, verbose=True):
    """
    Gibt (clean_body_text, info_dict) zurueck.
    clean_body_text passt direkt in parse_email_body().
    info_dict enthaelt 'unsure' (Liste) und 'dropped' (ungueltige Namen).
    """
    token = token or os.environ.get("GH_MODELS_TOKEN") or os.environ.get("GITHUB_TOKEN")
    if not token:
        raise RuntimeError("Kein GH_MODELS_TOKEN / GITHUB_TOKEN gesetzt.")

    roster = build_roster_text(name_map, gruppen_template)
    user_msg = (f"ROSTER (nur diese Personen sind gueltig):\n{roster}\n\n"
                f"WhatsApp-Chat:\n{raw_chat.strip()}")

    content = _call_github_models(
        [{"role": "system", "content": SYSTEM_PROMPT},
         {"role": "user", "content": user_msg}],
        token, model)

    return clean_from_model_json(content, name_map, verbose=verbose)


def clean_from_model_json(content, name_map, verbose=True):
    """Validiert + normalisiert die KI-JSON-Antwort zu sauberem Body-Text.
    Separat testbar (ohne API)."""
    # JSON robust extrahieren (falls doch mal Text drumherum kommt)
    m = re.search(r"\{.*\}", content, re.DOTALL)
    raw = m.group(0) if m else content
    data = json.loads(raw)

    lookup = _canon_set(name_map)
    seen = set()
    lines = []
    dropped = []
    for e in data.get("entries", []):
        date = str(e.get("date", "")).strip()
        dm = re.match(r"^(\d{1,2})\.(\d{1,2})", date)
        canon = _validate_name(e.get("name", ""), lookup)
        if not dm or not canon:
            dropped.append(f"{e.get('date','?')} {e.get('name','?')}")
            continue
        date_norm = f"{int(dm.group(1)):02d}.{int(dm.group(2)):02d}"
        key = (date_norm, canon)
        if key in seen:
            continue
        seen.add(key)
        lines.append(f"{date_norm} {canon}")

    lines.sort()
    info = {"unsure": data.get("unsure", []), "dropped": dropped, "count": len(lines)}
    if verbose:
        print(f"[KI-PARSER] {len(lines)} gueltige Eintraege, "
              f"{len(dropped)} verworfen, {len(info['unsure'])} unsicher.")
        for u in info["unsure"]:
            print(f"  [unsicher] {u}")
        for d in dropped:
            print(f"  [verworfen] {d}")
    return "\n".join(lines), info
