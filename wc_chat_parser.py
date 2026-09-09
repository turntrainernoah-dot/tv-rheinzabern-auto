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


# ====================================================================
#  DETERMINISTISCHER CHAT-PARSER (ab 03.09.2026)
#  --------------------------------------------------------------
#  GitHub Models (die bisherige KI-Grundlage fuer _call_github_models
#  oben) wurde von GitHub am 30.07.2026 vollstaendig abgeschaltet
#  (HTTP 410 "github_models_retirement_brownout"). Seitdem schlug die
#  KI-Auswertung im Hintergrund still fehl -- der Workflow lief
#  "gruen" durch, aber es passierte nichts ("Kein verwertbarer Chat
#  gefunden"), ohne jede Fehlermeldung an Noah.
#
#  Ersatz: ein regelbasierter Parser, der dieselben Regeln aus dem
#  SYSTEM_PROMPT oben nachbildet (Selbst-/Fremdmeldung, "gestern",
#  "vorgestern", Wochentagsnamen, "jeden Tag"-Sammelaussagen). Braucht
#  keinen Account/Token/Internetzugriff mehr -- laeuft rein lokal.
#  _call_github_models() bleibt oben unveraendert erhalten, falls
#  GitHub oder ein anderer Anbieter spaeter wieder eine kostenlose
#  Chat-Vervollstaendigung anbietet.
# ====================================================================

import datetime as _dt

_WOCHENTAGE_CHAT = {
    "montag": 0, "dienstag": 1, "mittwoch": 2, "donnerstag": 3,
    "freitag": 4, "samstag": 5, "sonntag": 6,
}

# Bugfix (Noah, 09.09.2026): WhatsApp exportiert den Zeitstempel je nach
# Geraet/Locale in UNTERSCHIEDLICHER Reihenfolge -- "[TT.MM.JJ, HH:MM]"
# (Datum zuerst) ODER "[HH:MM, TT.MM.JJJJ]" (Zeit zuerst). Der deterministische
# Parser (ab 03.09.2026) akzeptierte bisher NUR die Datum-zuerst-Form --
# Noahs echter Chat-Export nutzt aber durchgehend Zeit-zuerst ("[14:04,
# 5.9.2026] ..."), wodurch KEINE einzige Zeile erkannt wurde (0 Eintraege).
# Beide Formen werden jetzt akzeptiert, damit es unabhaengig vom
# Export-Geraet funktioniert.
_MSG_START_RE_DATUM_ZUERST = re.compile(
    r"^\[\s*(\d{1,2})\.(\d{1,2})\.(\d{2,4})\s*,\s*(\d{1,2}):(\d{2})(?::\d{2})?\s*\]\s*([^:]+?):\s*(.*)$"
)
_MSG_START_RE_ZEIT_ZUERST = re.compile(
    r"^\[\s*(\d{1,2}):(\d{2})(?::\d{2})?\s*,\s*(\d{1,2})\.(\d{1,2})\.(\d{2,4})\s*\]\s*([^:]+?):\s*(.*)$"
)
_GESTERN_RE = re.compile(
    r"f[uü]r\s+gestern|nachtr[aä]glich|war\s+gestern\s+auch\s+da|\bgestern\b", re.I)
_VORGESTERN_RE = re.compile(r"\bvorgestern\b", re.I)
_JEDEN_TAG_RE = re.compile(
    r"jeden\s+tag|jeden\s+trainingstag|immer\s+da|jedes\s*mal|an\s+allen\s+tagen|alle\s+tage", re.I)
_WEEKDAY_RE = re.compile(
    r"\b(?:am\s+)?(montag|dienstag|mittwoch|donnerstag|freitag|samstag|sonntag)\b", re.I)


def _match_msg_start(line):
    """Matcht eine Zeitstempel-Zeile in BEIDEN moeglichen Reihenfolgen
    (Datum-zuerst oder Zeit-zuerst, siehe Bugfix-Kommentar oben) und gibt
    einheitlich (day, mon, yr, hh, mi, sender, body) zurueck, oder None."""
    m = _MSG_START_RE_DATUM_ZUERST.match(line)
    if m:
        day, mon, yr, hh, mi, sender, body = m.groups()
        return day, mon, yr, hh, mi, sender, body
    m = _MSG_START_RE_ZEIT_ZUERST.match(line)
    if m:
        hh, mi, day, mon, yr, sender, body = m.groups()
        return day, mon, yr, hh, mi, sender, body
    return None


def _split_chat_messages(raw_chat):
    """Zerlegt den rohen WhatsApp-Export in Nachrichten {date, sender, body}.
    Mehrzeilige Nachrichten (kein neuer Zeitstempel) werden an die vorherige
    Nachricht angehaengt."""
    messages = []
    cur = None
    for raw_line in raw_chat.splitlines():
        line = raw_line.strip()
        matched = _match_msg_start(line)
        if matched:
            if cur:
                messages.append(cur)
            day, mon, yr, hh, mi, sender, body = matched
            yr = int(yr)
            if yr < 100:
                yr += 2000
            try:
                mdate = _dt.date(yr, int(mon), int(day))
            except ValueError:
                cur = None
                continue
            cur = {"date": mdate, "sender": sender.strip(), "body": body.strip()}
        elif cur is not None and line:
            cur["body"] = (cur["body"] + " " + line).strip()
    if cur:
        messages.append(cur)
    return messages


def _find_name_in_text(text, lookup, exclude_canon=None):
    """Laengster passender Roster-Alias in text (wortgrenzen-sicher), optional
    unter Ausschluss eines kanonischen Namens. None wenn nichts passt."""
    t = " " + str(text).lower() + " "
    best = None
    for alias, canon in lookup.items():
        if not alias or (exclude_canon and canon == exclude_canon):
            continue
        pat = r"(?<![0-9a-zäöüß])" + re.escape(alias) + r"(?![0-9a-zäöüß])"
        if re.search(pat, t):
            if best is None or len(alias) > len(best[0]):
                best = (alias, canon)
    return best[1] if best else None


def _resolve_weekday(msg_date, weekday_name_de):
    """Findet den Tag mit dem genannten Wochentagsnamen in der Woche der
    Nachricht (Mo-So), analog resolve_date_modifier() in auto_wochenchallenge.py."""
    target = _WOCHENTAGE_CHAT.get(weekday_name_de.lower())
    if target is None:
        return msg_date
    days_diff = (target - msg_date.weekday()) % 7
    if days_diff > 3:
        days_diff -= 7
    return msg_date + _dt.timedelta(days=days_diff)


def parse_chat_deterministic(raw_chat, name_map):
    """Regelbasierter Ersatz fuer die KI-Auswertung (siehe SYSTEM_PROMPT oben).
    Gibt (clean_body_text, info_dict) zurueck -- gleiche Form wie zuvor mit KI.
    info_dict: 'unsure' (Liste), 'dropped' (Liste), 'count' (int)."""
    lookup = _canon_set(name_map)
    messages = _split_chat_messages(raw_chat)

    entries = set()          # {(date_str, canon)}
    unsure = []
    dropped = []
    jeden_tag_targets = []   # kanon. Namen aus "jeden Tag"-Aussagen
    all_msg_dates = [m["date"] for m in messages]

    for msg in messages:
        body = msg["body"]
        if not body or "?" in body:
            continue  # leer oder vermutlich eine Frage -> kein Report

        sender_canon = _find_name_in_text(msg["sender"], lookup)

        if _JEDEN_TAG_RE.search(body):
            target = _find_name_in_text(body, lookup) or sender_canon
            if target:
                jeden_tag_targets.append(target)
            else:
                unsure.append("%s %s: 'jeden Tag' aber kein Name erkannt" % (
                    msg["date"].strftime("%d.%m"), msg["sender"]))
            continue

        other_canon = _find_name_in_text(body, lookup, exclude_canon=sender_canon)
        subject = other_canon or sender_canon
        if not subject:
            unsure.append("%s %s: kein Roster-Name erkannt (%r)" % (
                msg["date"].strftime("%d.%m"), msg["sender"], body[:60]))
            continue

        if _VORGESTERN_RE.search(body):
            day = msg["date"] - _dt.timedelta(days=2)
        elif _GESTERN_RE.search(body):
            day = msg["date"] - _dt.timedelta(days=1)
        else:
            wd = _WEEKDAY_RE.search(body)
            day = _resolve_weekday(msg["date"], wd.group(1)) if wd else msg["date"]

        entries.add(("%02d.%02d" % (day.day, day.month), subject))

    if jeden_tag_targets:
        if all_msg_dates:
            dmin, dmax = min(all_msg_dates), max(all_msg_dates)
            rng = []
            d = dmin
            while d <= dmax:
                rng.append(d)
                d += _dt.timedelta(days=1)
            for target in jeden_tag_targets:
                for dd in rng:
                    entries.add(("%02d.%02d" % (dd.day, dd.month), target))
        else:
            unsure.append("'jeden Tag'-Aussage(n) gefunden, aber kein Datum im Chat: " +
                           ", ".join(jeden_tag_targets))

    lines = sorted("%s %s" % (d, n) for d, n in entries)
    info = {"unsure": unsure, "dropped": dropped, "count": len(lines)}
    return "\n".join(lines), info


def chat_to_clean_body(raw_chat, name_map, gruppen_template,
                       token=None, model=None, verbose=True):
    """
    Gibt (clean_body_text, info_dict) zurueck.
    clean_body_text passt direkt in parse_email_body().
    info_dict enthaelt 'unsure' (Liste) und 'dropped' (ungueltige Namen).

    Seit 03.09.2026: nutzt parse_chat_deterministic() statt GitHub Models
    (von GitHub am 30.07.2026 komplett abgeschaltet, siehe Kommentar oben
    bei _call_github_models). token/model bleiben aus Kompatibilitaets-
    gruenden im Funktionskopf, werden aber nicht mehr benutzt.
    """
    clean, info = parse_chat_deterministic(raw_chat, name_map)
    if verbose:
        print("[CHAT-PARSER] %d Eintraege, %d unsicher." % (
            info["count"], len(info.get("unsure", []))))
        for u in info.get("unsure", []):
            print("  [unsicher] " + u)
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
