# -*- coding: utf-8 -*-
"""
wc_chat_parser.py
=================
Wandelt einen ROHEN WhatsApp-Wochenchallenge-Chat in das saubere Zeilenformat um,
das auto_wochenchallenge.parse_email_body() bereits versteht:

    DD.MM Kanonischer-Name
    z.B.  06.06 Vorname X.
          07.06 Vorname Y.

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
   gemappt werden. Gib IMMER exakt den kanonischen Roster-Namen aus (z.B. "Vorname X.",
   "Vorname Y."), nie den Spitznamen.
2. Wer eine Nachricht schreibt, meldet normalerweise SICH SELBST - ausser der Text
   nennt klar eine andere Person.
3. DATUM pro Meldung aus dem Zeitstempel "[TT.MM.JJ, HH:MM] Name: ...":
     - "heute" / keine Tagesangabe   -> das Nachrichtendatum
     - "gestern" / "fuer gestern" / "und fuer gestern" / "noch fuer gestern" /
       "nachtraeglich" / "war gestern auch da"  -> Tag VOR dem Nachrichtendatum.
       Das ist ein EIGENER, ZUSAETZLICHER Trainingstag - nicht das Nachrichtendatum ersetzen!
     - "vorgestern"                   -> zwei Tage vor dem Nachrichtendatum
     - "am Montag" / "Montag" / Wochentagsname -> der betreffende Wochentag der Nachrichtenwoche
   Ausgabe als "DD.MM" (zweistellig).
4. Ignoriere alles, was keine Trainings-Meldung ist (Smalltalk, Fragen, reine Emojis).
5. Eine Person kann an MEHREREN Tagen trainiert haben. Sammle ALLE gemeldeten Tage einer
   Person aus ALLEN ihren Nachrichten - auch spaetere Nachtraege und "fuer gestern"-Meldungen.
   Pro Person und Tag nur EIN Eintrag (keine Duplikate am selben Tag).
6. SAMMEL-/FREITEXT-Aussagen (auch ohne Zeitstempel, egal wo im Text), z.B.
   "X hat jeden Tag trainiert", "X war immer/jedes Mal da", "X an allen Tagen":
   bedeuten, dass X an ALLEN Trainingstagen des Zeitraums trainiert hat. Der Zeitraum
   ergibt sich aus den Datumsangaben im Chat (fruehestes bis spaetestes Datum, i.d.R.
   Samstag bis Dienstag, 4-5 Tage). Gib fuer X je EINEN Eintrag pro Tag dieses Zeitraums aus.
7. Kannst du eine Meldung keinem Roster-Namen sicher zuordnen, LASS SIE WEG und trage sie
   unter "unsure" mit kurzer Begruendung ein. Aber verwirf NIEMALS eine sichere Meldung nur
   weil sie nachtraeglich, als "fuer gestern" oder als Sammelaussage kam.

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


def _augment_deterministic(raw_chat, clean_str, name_map):
    """Sichert zwei KI-Schwachstellen deterministisch ab:
       - 'X hat jeden Tag trainiert' / 'immer da'  -> X an ALLEN Tagen des Zeitraums
       - 'fuer gestern' / 'gestern'-Nachtraege     -> zusaetzlicher Tag (Nachrichtendatum - 1)
    Ergaenzt fehlende Eintraege ohne Doppelung. Namen ueber NAME_MAP-Aliasse."""
    import datetime
    lookup = {str(k).lower(): v[1] for k, v in name_map.items() if k}
    entries = set()
    for ln in clean_str.splitlines():
        ln = ln.strip()
        if not ln:
            continue
        parts = ln.split(" ", 1)
        if len(parts) == 2:
            entries.add((parts[0], parts[1]))
    added = []
    def to_dt(dm):
        dd, mm = dm.split("."); return datetime.date(2026, int(mm), int(dd))
    def find_name(text):
        t = " " + str(text).lower() + " "
        best = None
        for alias, canon in lookup.items():
            if re.search(r"(?<![0-9a-z\u00e4\u00f6\u00fc\u00df])" + re.escape(alias) + r"(?![0-9a-z\u00e4\u00f6\u00fc\u00df])", t):
                if best is None or len(alias) > len(best[0]):
                    best = (alias, canon)
        return best[1] if best else None
    dates = []
    for d, _ in entries:
        try: dates.append(to_dt(d))
        except Exception: pass
    if dates:
        dmin, dmax = min(dates), max(dates)
        rng = []; d = dmin
        while d <= dmax:
            rng.append(d); d = d + datetime.timedelta(days=1)
        for m in re.finditer(r"[^\n.!?]*\b(?:jeden\s+tag|jeden\s+trainingstag|immer\s+da|jedes\s*mal|an\s+allen\s+tagen|alle\s+tage)\b[^\n.!?]*", raw_chat, re.I):
            canon = find_name(m.group(0))
            if not canon: continue
            for dd in rng:
                key = ("%02d.%02d" % (dd.day, dd.month), canon)
                if key not in entries:
                    entries.add(key); added.append("jeden-tag %s %s" % key)
    msg_re = re.compile(r"\[\s*\d{1,2}:\d{2},\s*(\d{1,2})\.(\d{1,2})\.(\d{2,4})\s*\]\s*([^:]*):\s*(.*)")
    for line in raw_chat.splitlines():
        mm = msg_re.match(line.strip())
        if not mm: continue
        day, mon, yr = int(mm.group(1)), int(mm.group(2)), int(mm.group(3))
        body = mm.group(5)
        if yr < 100: yr += 2000
        if not re.search(r"\bgestern\b|nachtr\u00e4glich|nachtraeglich", body, re.I): continue
        try: prev = datetime.date(yr, mon, day) - datetime.timedelta(days=1)
        except Exception: continue
        canon = find_name(body) or find_name(mm.group(4))
        if not canon: continue
        key = ("%02d.%02d" % (prev.day, prev.month), canon)
        if key not in entries:
            entries.add(key); added.append("fuer-gestern %s %s" % key)
    lines = sorted("%s %s" % (d, n) for d, n in entries)
    return "\n".join(lines), added


def chat_to_clean_body(raw_chat, name_map, gruppen_template,
                       token=None, model=None, verbose=True):
    """
    Gibt (clean_body_text, info_dict) zurueck.
    clean_body_text passt direkt in parse_email_body().
    info_dict enthaelt 'unsure' (Liste) und 'dropped' (ungueltige Namen).
    """
    # Mehrere Token-Quellen nacheinander probieren (selbstheilend): explizit
    # uebergebenes Token, dann GH_MODELS_TOKEN-Secret, dann das in Actions immer
    # gueltige GITHUB_TOKEN (Job-Permission models: read). So scheitert der
    # KI-Aufruf nicht mehr, wenn ein einzelnes Token ungueltig ist.
    candidates = []
    for _t in (token, os.environ.get("GH_MODELS_TOKEN"), os.environ.get("GITHUB_TOKEN")):
        if _t and _t not in candidates:
            candidates.append(_t)
    if not candidates:
        raise RuntimeError("Kein GH_MODELS_TOKEN / GITHUB_TOKEN gesetzt.")

    roster = build_roster_text(name_map, gruppen_template)
    user_msg = (f"ROSTER (nur diese Personen sind gueltig):\n{roster}\n\n"
                f"WhatsApp-Chat:\n{raw_chat.strip()}")
    messages = [{"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_msg}]

    content = None
    errors = []
    for _tk in candidates:
        try:
            content = _call_github_models(messages, _tk, model)
            break
        except urllib.error.HTTPError as e:
            try: _d = e.read().decode("utf-8", "replace")[:160]
            except Exception: _d = ""
            errors.append(f"HTTP {e.code} (Token ...{_tk[-4:]}): {_d}")
        except Exception as e:
            errors.append(f"{type(e).__name__} (Token ...{_tk[-4:]}): {e}")
    if content is None:
        raise RuntimeError("GitHub-Models-Aufruf fehlgeschlagen: " + " | ".join(errors))

    clean, info = clean_from_model_json(content, name_map, verbose=False)
    clean, added = _augment_deterministic(raw_chat, clean, name_map)
    info["deterministic_added"] = added
    info["count"] = len([l for l in clean.splitlines() if l.strip()])
    if verbose:
        print("[KI-PARSER] %d Eintraege, %d deterministisch ergaenzt, %d verworfen, %d unsicher." % (
            info["count"], len(added), len(info.get("dropped", [])), len(info.get("unsure", []))))
        for a in added: print("  [det+] " + a)
    return clean, info


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
