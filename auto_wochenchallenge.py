#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
auto_wochenchallenge.py – Automatische Wochenchallenge-Erstellung
==================================================================
Ablauf:
  1. IMAP: Letzten ungelesenen Mail mit Betreff "WC" lesen
  2. Parse: TAG.MONAT Name Zeilen extrahieren
  3. Prüfen ob diese Woche schon verarbeitet (wc_state_auto.json)
  4. create_wochenchallenge_v2.py mit neuer Config per subprocess ausführen
  5. PDF via SFTP hochladen nach /wochen-challenge/
  6. Punkte-State aktualisieren
  7. E-Mail Bestätigung senden

Name-Format im Mail:  TAG.MONAT Name [optional: Emoji / für gestern / Wochentag]
Bsp:  13.6 Felix G1
      14.6 Mika
      14.6 Finn G3 für gestern
      15.6 Sinan 💪
"""

import json, os, re, sys, imaplib, subprocess, urllib.request, urllib.parse, hashlib
from datetime import date, timedelta, datetime
import email as email_lib
from email.header import decode_header

import paramiko

# ════════════════════════════════════════════════════════════════
#  KONFIGURATION
# ════════════════════════════════════════════════════════════════

SSH_HOST      = os.environ.get("SSH_HOST",     "access-5017462830.webspace-host.com")
SSH_USER      = os.environ.get("SSH_USER",     "a2358459")
SSH_PASSWORD  = os.environ.get("SSH_PASSWORD", "")
SSH_PORT      = int(os.environ.get("SSH_PORT", "22"))

GMAIL_USER         = os.environ.get("GMAIL_USER",         "turntrainernoah@gmail.com")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
WHATSAPP_PHONE    = os.environ.get("WHATSAPP_PHONE", "")
CALLMEBOT_APIKEY  = os.environ.get("CALLMEBOT_APIKEY", "")

# ════════════════════════════════════════════════════════════════
#  NAME-MAPPING
#  Schlüssel: lowercase wie er im Mail vorkommen kann
#  Wert: (Gruppe, interner Name im WC-Script)
# ════════════════════════════════════════════════════════════════

NAME_MAP = {
    # ── G1 ──────────────────────────────────────────────────────
    # Felix E. – alle Schreibweisen
    "felix g1":      ("G1", "Felix E."),
    "felix e.":      ("G1", "Felix E."),
    "felix e":       ("G1", "Felix E."),
    # Finn M. – alle Schreibweisen
    "finn g1":       ("G1", "Finn M."),
    "finn klein":    ("G1", "Finn M."),
    "finn m.":       ("G1", "Finn M."),
    "finn m":        ("G1", "Finn M."),
    # Sinan Y.
    "sinan":         ("G1", "Sinan Y."),
    "sinan y.":      ("G1", "Sinan Y."),
    "sinan y":       ("G1", "Sinan Y."),
    # Ilyas E.
    "ilyas":         ("G1", "Ilyas E."),
    "ilyas e.":      ("G1", "Ilyas E."),
    "ilyas e":       ("G1", "Ilyas E."),
    # Jonathan S.
    "jonathan":      ("G1", "Jonathan S."),
    "jonathan s.":   ("G1", "Jonathan S."),
    "jonathan s":    ("G1", "Jonathan S."),
    # Hannes G.
    "hannes":        ("G1", "Hannes G."),
    "hannes g.":     ("G1", "Hannes G."),
    "hannes g":      ("G1", "Hannes G."),
    # Ben B.
    "ben g1":        ("G1", "Ben B."),
    "ben klein":     ("G1", "Ben B."),
    "ben baron":     ("G1", "Ben B."),
    "ben b.":        ("G1", "Ben B."),
    "ben b":         ("G1", "Ben B."),
    # ── G2 ──────────────────────────────────────────────────────
    "henry":         ("G2", "Henry K."),
    "henry k.":      ("G2", "Henry K."),
    "henry k":       ("G2", "Henry K."),
    "matti":         ("G2", "Matti G."),
    "matti g.":      ("G2", "Matti G."),
    "matti g":       ("G2", "Matti G."),
    "levent":        ("G2", "Levent K."),
    "levent k.":     ("G2", "Levent K."),
    "levent k":      ("G2", "Levent K."),
    "caius":         ("G2", "Caius C."),
    "caius c.":      ("G2", "Caius C."),
    "caius c":       ("G2", "Caius C."),
    # ── G3 ──────────────────────────────────────────────────────
    # Erik E.
    "erik":          ("G3", "Erik E."),
    "erik e.":       ("G3", "Erik E."),
    "erik e":        ("G3", "Erik E."),
    # Artem T.
    "artem":         ("G3", "Artem T."),
    "artem t.":      ("G3", "Artem T."),
    "artem t":       ("G3", "Artem T."),
    # Finn T.
    "finn g3":       ("G3", "Finn T."),
    "finn gross":    ("G3", "Finn T."),
    "finn gro\xdf":  ("G3", "Finn T."),  # "finn groß"
    "finn t.":       ("G3", "Finn T."),
    "finn t":        ("G3", "Finn T."),
    # Ben F.
    "ben g3":        ("G3", "Ben F."),
    "ben gross":     ("G3", "Ben F."),
    "ben gro\xdf":   ("G3", "Ben F."),   # "ben groß"
    "ben f.":        ("G3", "Ben F."),
    "ben f":         ("G3", "Ben F."),
    # Michael K.
    "michael":       ("G3", "Michael K."),
    "michael k.":    ("G3", "Michael K."),
    "michael k":     ("G3", "Michael K."),
    # ── G4 ──────────────────────────────────────────────────────
    # Felix L.
    "felix g4":      ("G4", "Felix L."),
    "felix l.":      ("G4", "Felix L."),
    "felix l":       ("G4", "Felix L."),
    # Anton K.
    "anton":         ("G4", "Anton K."),
    "anton k.":      ("G4", "Anton K."),
    "anton k":       ("G4", "Anton K."),
    # Mika W.
    "mika":          ("G4", "Mika W."),
    "mika w.":       ("G4", "Mika W."),
    "mika w":        ("G4", "Mika W."),
    "mika werling":  ("G4", "Mika W."),
    # Jamie G.
    "jamie":         ("G4", "Jamie G."),
    "jamie g.":      ("G4", "Jamie G."),
    "jamie g":       ("G4", "Jamie G."),
}

# Mehrdeutige Namen (brauchen Gruppenangabe)
AMBIGUOUS_NAMES = {"felix", "finn", "ben"}

# Alle Namen in der richtigen Reihenfolge für den WC-Script (neues Format)
WC_GRUPPEN_TEMPLATE = {
    "G1": ["Felix E.", "Finn M.", "Sinan Y.", "Ilyas E.", "Jonathan S.", "Hannes G.", "Ben B."],
    "G2": ["Henry K.", "Matti G.", "Levent K.", "Caius C."],
    "G3": ["Erik E.", "Artem T.", "Finn T.", "Ben F.", "Michael K."],
    "G4": ["Felix L.", "Anton K.", "Mika W.", "Jamie G."],
}

# Initiale Vorpunkte (Stand nach Woche 06.06–09.06.2026, neues Name-Format)
DEFAULT_VORPUNKTE = {
    "G1_Felix E.": 8,    "G1_Finn M.": 3,    "G1_Sinan Y.": 0,
    "G1_Ilyas E.": 8,    "G1_Jonathan S.": 7, "G1_Hannes G.": 5, "G1_Ben B.": 4,
    "G2_Henry K.": 8,    "G2_Matti G.": 8,   "G2_Levent K.": 8,  "G2_Caius C.": 8,
    "G3_Erik E.": 3,     "G3_Artem T.": 2,   "G3_Finn T.": 4,    "G3_Ben F.": 2, "G3_Michael K.": 3,
    "G4_Felix L.": 3,    "G4_Anton K.": 3,   "G4_Mika W.": 21,   "G4_Jamie G.": 0,
}

# Wochentag-Mapping (Deutsch → Python weekday, Mo=0 … So=6)
WOCHENTAGE = {
    "montag": 0, "mo": 0,
    "dienstag": 1, "di": 1,
    "mittwoch": 2, "mi": 2,
    "donnerstag": 3, "do": 3,
    "freitag": 4, "fr": 4,
    "samstag": 5, "sa": 5,
    "sonntag": 6, "so": 6,
}

# ════════════════════════════════════════════════════════════════
#  SFTP-VERBINDUNG
# ════════════════════════════════════════════════════════════════

def get_sftp():
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    client.connect(SSH_HOST, port=SSH_PORT, username=SSH_USER,
                   password=SSH_PASSWORD, timeout=20)
    return client, client.open_sftp()

def load_wc_state(sftp):
    try:
        f     = sftp.open("wc_state_auto.json", "r")
        state = json.loads(f.read().decode("utf-8"))
        f.close()
        print("wc_state_auto.json geladen.")
        return state
    except Exception:
        print("wc_state_auto.json nicht vorhanden – verwende Default.")
        return {"last_processed_week_start": None, "alle_vorpunkte": DEFAULT_VORPUNKTE.copy()}

def save_wc_state(sftp, state):
    data = json.dumps(state, indent=2, ensure_ascii=False).encode("utf-8")
    f    = sftp.open("wc_state_auto.json", "wb")
    f.write(data)
    f.close()
    print("[OK] wc_state_auto.json aktualisiert.")

# ════════════════════════════════════════════════════════════════
#  WHATSAPP (CallMeBot)
# ════════════════════════════════════════════════════════════════

def send_whatsapp(text):
    """Sendet eine WhatsApp-Nachricht via CallMeBot (kostenlos)."""
    if not WHATSAPP_PHONE or not CALLMEBOT_APIKEY:
        print(f"[WA-TEST] {text[:200]}")
        return
    try:
        encoded = urllib.parse.quote(text)
        url = f"https://api.callmebot.com/whatsapp.php?phone={WHATSAPP_PHONE}&text={encoded}&apikey={CALLMEBOT_APIKEY}"
        with urllib.request.urlopen(url, timeout=15) as r:
            print(f"[OK] WhatsApp gesendet (HTTP {r.status}): {text[:80]}...")
    except Exception as e:
        print(f"[FEHLER] WhatsApp konnte nicht gesendet werden: {e}")

# ════════════════════════════════════════════════════════════════
#  IMAP
# ════════════════════════════════════════════════════════════════

def imap_connect():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    return mail

def get_email_body(msg):
    """Extrahiert Plaintext-Body aus email.message.Message."""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))
            if ct == "text/plain" and "attachment" not in cd:
                charset = part.get_content_charset() or "utf-8"
                return part.get_payload(decode=True).decode(charset, errors="replace")
    else:
        charset = msg.get_content_charset() or "utf-8"
        return msg.get_payload(decode=True).decode(charset, errors="replace")
    return ""

def fetch_latest_wc_mail():
    """
    Holt die neueste Mail mit 'WC' im Betreff.
    Gibt (mail_id_bytes, raw_body_text) zurück oder (None, None).
    """
    mail = imap_connect()
    mail.select("inbox")

    # Suche alle Mails mit WC im Betreff (auch bereits gelesene)
    status, data = mail.search(None, 'SUBJECT "WC"')
    if status != "OK" or not data[0]:
        mail.logout()
        return None, None

    mail_ids = data[0].split()
    # Neueste zuletzt → letzte ID
    latest_id = mail_ids[-1]

    status, msg_data = mail.fetch(latest_id, "(RFC822)")
    if status != "OK":
        mail.logout()
        return None, None

    raw = msg_data[0][1]
    msg = email_lib.message_from_bytes(raw)

    body = get_email_body(msg)
    mail.logout()
    return latest_id, body

# ════════════════════════════════════════════════════════════════
#  PARSER
# ════════════════════════════════════════════════════════════════

def strip_emojis(text):
    """Entfernt Emojis und Non-ASCII-Sonderzeichen (behält Umlauts)."""
    # Behalte: ASCII, Umlauts (ä,ö,ü,Ä,Ö,Ü,ß), Leerzeichen, Ziffern, Bindestriche
    return re.sub(r'[^\x00-\x7Fäöüÿ\xc0-\xff\s0-9\-]', '', text).strip()

def resolve_date_modifier(parsed_date, rest_text):
    """
    Gibt (angepasstes_datum, bereinigter_name) zurück.
    Behandelt: 'für gestern', Wochentagsnamen.
    """
    d    = parsed_date
    text = rest_text

    # "für gestern" / "fuer gestern"
    if re.search(r'f[uü]r\s+gestern', text, re.IGNORECASE):
        d    = d - timedelta(days=1)
        text = re.sub(r'f[uü]r\s+gestern', '', text, flags=re.IGNORECASE).strip()
        return d, text

    # Wochentag am Ende (z.B. "Felix Montag" → Felix trainierte am Montag dieser Woche)
    for wd_de, wd_py in WOCHENTAGE.items():
        if text.lower().endswith(' ' + wd_de) or text.lower() == wd_de:
            # Finde den wd_py-Wochentag in der gleichen Woche wie d
            days_diff = (wd_py - d.weekday()) % 7
            # Wenn Wochentag in der Zukunft liegt, gehe zurück eine Woche
            if days_diff > 3:
                days_diff -= 7
            adjusted = d + timedelta(days=days_diff)
            # Name = alles vor dem Wochentag
            name_part = text[:-(len(wd_de))].strip().rstrip()
            if not name_part:
                # Wochentag war der ganze Rest – kein Name, skip
                return d, text
            return adjusted, name_part

    return d, text

def parse_email_body(body):
    """
    Parst den Mail-Body nach Trainingseinträgen.
    Format: TAG.MONAT Name [Extras]

    Gibt zurück: list of (date_obj, group_str, wc_internal_name_str)
    Wirft ValueError bei unbekannten / mehrdeutigen Namen.
    """
    entries = []
    errors  = []

    line_pat = re.compile(r'^(\d{1,2})\.(\d{1,2})\s+(.+)', re.UNICODE)
    today = date.today()

    for raw_line in body.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        m = line_pat.match(line)
        if not m:
            continue

        day_str, month_str, rest = m.group(1), m.group(2), m.group(3).strip()
        day, month = int(day_str), int(month_str)

        # Jahr bestimmen (laufendes Jahr, bei Monat weit zurück → letztes Jahr)
        year = today.year
        # Wenn der Monat mehr als 6 Monate in der Zukunft liegt, nimm letztes Jahr
        if month > today.month + 6:
            year -= 1
        # Wenn der Monat mehr als 6 Monate in der Vergangenheit liegt, nimm nächstes Jahr
        elif month < today.month - 6:
            year += 1

        try:
            d = date(year, month, day)
        except ValueError:
            continue  # Ungültiges Datum überspringen

        # Emojis entfernen
        rest = strip_emojis(rest)
        if not rest:
            continue

        # Datumsmodifikator auflösen
        d, name_text = resolve_date_modifier(d, rest)
        name_text = name_text.strip()
        if not name_text:
            continue

        # Mehrfache Leerzeichen normieren
        name_text = re.sub(r'\s+', ' ', name_text).strip()

        # Name nachschlagen (case-insensitive)
        # Klammern normalisieren: "Ben (G1)" → "ben g1", "Finn (G3)" → "finn g3"
        name_lower = name_text.lower().strip()
        name_lower = re.sub(r'[()]', '', name_lower)
        name_lower = re.sub(r'\s+', ' ', name_lower).strip()

        if name_lower in NAME_MAP:
            group, wc_name = NAME_MAP[name_lower]
            entries.append((d, group, wc_name))
        elif name_lower in AMBIGUOUS_NAMES:
            errors.append(
                f"Mehrdeutiger Name '{name_text}' – bitte Gruppe angeben "
                f"(z.B. '{name_text} G1' oder '{name_text} G4')."
            )
        else:
            errors.append(f"Unbekannter Name: '{name_text}'")

    if errors:
        raise ValueError("Fehler beim Parsen des WC-Mails:\n\n" + "\n".join(errors))

    return entries

# ════════════════════════════════════════════════════════════════
#  GRUPPEN-DICT AUFBAUEN
# ════════════════════════════════════════════════════════════════

def calc_saturday(d):
    """Gibt den Samstag der Woche zurück die d enthält (Sa=Wochenstart)."""
    # weekday: Mon=0, Sat=5
    days_since_sat = (d.weekday() - 5) % 7
    return d - timedelta(days=days_since_sat)

def calc_friday(d):
    """Gibt den Freitag der Woche zurück die d enthält (Fr=Wochenstart)."""
    # weekday: Mon=0, Fri=4
    days_since_fri = (d.weekday() - 4) % 7
    return d - timedelta(days=days_since_fri)

def build_gruppen(entries, week_start, num_days):
    """Gibt GRUPPEN-Dict zurück (alle Namen, nur die Trainierten mit Tages-Indizes).
    num_days: 4 für Sa-Di, 5 für Fr-Di"""
    gruppen = {}
    for grp, names in WC_GRUPPEN_TEMPLATE.items():
        gruppen[grp] = {name: [] for name in names}

    for d, group, wc_name in entries:
        day_idx = (d - week_start).days
        if not (0 <= day_idx < num_days):
            print(f"[WARN] {wc_name} am {d} liegt ausserhalb Periode (idx={day_idx}), ignoriere.")
            continue
        current = gruppen.get(group, {}).get(wc_name)
        if current is None:
            print(f"[WARN] {wc_name} nicht in Gruppe {group} gefunden, ignoriere.")
            continue
        if day_idx not in current:
            current.append(day_idx)
            current.sort()

    return gruppen

# ════════════════════════════════════════════════════════════════
#  CONFIG-BLOCK GENERIEREN & ERSETZEN
# ════════════════════════════════════════════════════════════════

def build_config_block(start_datum, end_datum, tage_header_raw, gruppen, vorherige):
    """
    Erstellt den Python-Quellcode-String für den CONFIG-Block.
    tage_header_raw: list of strings like "Sa\n06.06." (mit echtem Newline-Char)
    """
    # TAGE_HEADER: echte Newlines als \\n-Escapes in der Quelldatei
    tage_parts = []
    for h in tage_header_raw:
        escaped = h.replace('\\', '\\\\').replace('\n', '\\n').replace('"', '\\"')
        tage_parts.append(f'"{escaped}"')
    tage_header_src = "[" + ", ".join(tage_parts) + "]"

    # GRUPPEN als Python-Dict (via JSON)
    gruppen_json = json.dumps(gruppen, indent=4, ensure_ascii=False)

    # VORHERIGE_PUNKTE als Python-Dict
    vorherige_json = json.dumps(vorherige, indent=4, ensure_ascii=False)

    return (
        f'START_DATUM = "{start_datum}"\n'
        f'END_DATUM   = "{end_datum}"\n'
        f'\n'
        f'TAGE_HEADER = {tage_header_src}\n'
        f'\n'
        f'GRUPPEN = {gruppen_json}\n'
        f'\n'
        f'VORHERIGE_PUNKTE = {vorherige_json}\n'
    )

def replace_config_block(template, new_config):
    """
    Ersetzt den CONFIG-Block im Template-Script.
    Sucht nach den exakten Marker-Zeilen.
    """
    sep          = "# ===================================================================\n"
    start_marker = sep + "#  CONFIG - NUR DIESEN BLOCK ANPASSEN\n" + sep
    end_marker   = sep + "#  FARBEN & DESIGN (nicht aendern)\n"

    start_idx = template.find(start_marker)
    if start_idx == -1:
        raise ValueError("CONFIG-Start-Marker nicht gefunden in create_wochenchallenge_v2.py")

    content_start = start_idx + len(start_marker)
    end_idx       = template.find(end_marker, content_start)
    if end_idx == -1:
        raise ValueError("FARBEN-Marker nicht gefunden in create_wochenchallenge_v2.py")

    return template[:content_start] + "\n" + new_config + "\n" + template[end_idx:]

# ════════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════════

def main():
    print(f"=== Auto-Wochenchallenge | {date.today()} ===\n")

    if not GMAIL_USER or not GMAIL_APP_PASSWORD:
        print("[ERROR] GMAIL_USER oder GMAIL_APP_PASSWORD nicht gesetzt.")
        sys.exit(1)

    # SFTP verbinden
    print("Verbinde SFTP...")
    ssh, sftp = get_sftp()
    wc_state  = load_wc_state(sftp)

    # Mail holen
    print("Suche WC-Mail...")
    mail_id, body = fetch_latest_wc_mail()

    if not mail_id or not body:
        print("Keine WC-Mail gefunden.")
        sftp.close(); ssh.close()
        return

    print(f"Mail gefunden (ID {mail_id.decode()}):")
    print("--- BODY PREVIEW ---")
    print(body[:400])
    print("---")

    # Parsen
    try:
        entries = parse_email_body(body)
    except ValueError as e:
        print(f"[ERROR] Parse-Fehler: {e}")
        send_whatsapp(f"Hi Noah, Cloude hier 🚨\n\nDie WC-Mail konnte nicht geparst werden:\n{str(e)[:300]}\n\nBitte WC-Mail nochmal mit korrektem Format schicken.")
        sftp.close(); ssh.close()
        return

    if not entries:
        print("Keine Trainingseintraege im Mail gefunden.")
        sftp.close(); ssh.close()
        return

    print(f"Geparst: {len(entries)} Eintraege")
    for d, grp, name in entries:
        print(f"  {d.strftime('%d.%m.%Y')} – {grp}: {name}")

    # Wochenstart berechnen (Fr oder Sa, je nach Flag oder Auto-Detect)
    min_date   = min(d for d, _, _ in entries)
    use_friday = wc_state.get("use_friday_start", False)

    if use_friday:
        week_start = calc_friday(min_date)
        num_days   = 5
        wday_abbr  = ["Fr", "Sa", "So", "Mo", "Di"]
        print("[INFO] use_friday_start=True → Periode Fr-Di (5 Tage)")
    elif min_date.weekday() == 4:
        # AUTO-DETECT: min_date ist ein Freitag → Fr-Di Periode (5 Tage)
        # calc_saturday(Freitag) würde sonst auf den VORHERIGEN Samstag zeigen
        # und eine Duplikat-Erkennung auslösen, obwohl es eine neue Woche ist.
        week_start = calc_friday(min_date)
        num_days   = 5
        wday_abbr  = ["Fr", "Sa", "So", "Mo", "Di"]
        print(f"[INFO] min_date {min_date} ist Freitag → automatisch Periode Fr-Di (5 Tage)")
    else:
        week_start = calc_saturday(min_date)
        num_days   = 4
        wday_abbr  = ["Sa", "So", "Mo", "Di"]

    end_date         = week_start + timedelta(days=num_days - 1)
    week_start_short = week_start.strftime("%d.%m.%y")
    start_datum      = week_start.strftime("%d.%m.%Y")
    end_datum        = end_date.strftime("%d.%m.%Y")

    print(f"\nWoche: {start_datum} – {end_datum} (ab {week_start_short})")

    # E-Mail-Hash berechnen (um neue Inhalte bei gleicher Woche zu erkennen)
    email_hash = hashlib.md5(body.encode("utf-8")).hexdigest()

    # Schon verarbeitet?
    last_week = wc_state.get("last_processed_week_start")
    last_hash = wc_state.get("last_email_hash", "")

    if last_week == week_start_short and email_hash == last_hash:
        print(f"Woche {week_start_short} + E-Mail-Hash unverändert. Fertig.")
        sftp.close(); ssh.close()
        return
    elif last_week == week_start_short:
        print(f"[INFO] Neue E-Mail für bereits verarbeitete Woche {week_start_short} → Aktualisiere WC-Liste...")
        # Alte Dateien auf Server löschen (werden gleich überschrieben, aber explizit löschen für sauberen State)
        for ext in ["pdf", "xlsx"]:
            old_remote = f"wochen-challenge/ab_{week_start_short}_Wochenchallenge.{ext}"
            try:
                sftp.remove(old_remote)
                print(f"  Gelöscht: {old_remote}")
            except Exception:
                pass  # Datei existierte noch nicht

    # GRUPPEN-Dict aufbauen
    gruppen = build_gruppen(entries, week_start, num_days)

    # VORHERIGE_PUNKTE aus State
    # WICHTIG: .get() mit Default hilft nur wenn Key fehlt, nicht wenn Value leer ({})
    all_vorpunkte = wc_state.get("alle_vorpunkte") or DEFAULT_VORPUNKTE.copy()
    if not wc_state.get("alle_vorpunkte"):
        print("[WARN] alle_vorpunkte leer/fehlend im State → verwende DEFAULT_VORPUNKTE")
    vorherige     = {}
    for grp, members in gruppen.items():
        for name in members:
            key = f"{grp}_{name}"
            vorherige[key] = all_vorpunkte.get(key, 0)

    # TAGE_HEADER
    tage_header = []
    for i, abbr in enumerate(wday_abbr):
        d_str = (week_start + timedelta(days=i)).strftime("%d.%m.")
        tage_header.append(f"{abbr}\n{d_str}")  # echtes Newline

    # Template laden und CONFIG ersetzen
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 "create_wochenchallenge_v2.py")
    if not os.path.exists(template_path):
        msg = f"create_wochenchallenge_v2.py nicht gefunden unter {template_path}"
        print(f"[ERROR] {msg}")
        send_whatsapp(f"Hi Noah, Cloude hier 🚨\n\nDas WC-Template (create_wochenchallenge_v2.py) fehlt auf dem Server!\n\nBitte manuell hochladen: C:\\Claude\\create_wochenchallenge_v2.py → via upload_now.bat")
        sftp.close(); ssh.close()
        return

    with open(template_path, "r", encoding="utf-8") as f:
        template = f.read()

    new_config = build_config_block(start_datum, end_datum, tage_header, gruppen, vorherige)
    try:
        modified = replace_config_block(template, new_config)
    except ValueError as e:
        print(f"[ERROR] {e}")
        send_whatsapp(f"Hi Noah, Cloude hier 🚨\n\nDas WC-Template konnte nicht angepasst werden:\n{str(e)[:300]}")
        sftp.close(); ssh.close()
        return

    # Schreibe modifiziertes Script nach /tmp/
    run_script = "/tmp/wc_run.py"
    with open(run_script, "w", encoding="utf-8") as f:
        f.write(modified)
    print(f"Script geschrieben: {run_script}")

    # Script ausführen
    print("Starte Wochenchallenge-Generierung...")
    result = subprocess.run(
        ["python3", run_script],
        capture_output=True, text=True, timeout=120
    )
    print("STDOUT:", result.stdout[:500])
    if result.returncode != 0:
        print("STDERR:", result.stderr[:500])
        send_whatsapp(
            f"Hi Noah, Cloude hier 🚨\n\nDas WC-Script ist abgestürzt (Code {result.returncode}):\n"
            f"{result.stderr[:300]}"
        )
        sftp.close(); ssh.close()
        return

    # PDF + XLSX finden
    datum_short = week_start.strftime("%d.%m.%y")
    ordner      = f"/tmp/Wochenchallenge/ab {datum_short}"
    pdf_path    = f"{ordner}/ab_{datum_short}_Wochenchallenge.pdf"
    xlsx_path   = f"{ordner}/ab_{datum_short}_Wochenchallenge.xlsx"

    if not os.path.exists(pdf_path):
        msg = f"PDF nicht gefunden: {pdf_path}\n\nScript-Output:\n{result.stdout[:500]}"
        print(f"[ERROR] {msg}")
        send_whatsapp(f"Hi Noah, Cloude hier 🚨\n\nDie WC-PDF wurde nicht generiert. Script lief durch aber Datei fehlt.\n{msg[:200]}")
        sftp.close(); ssh.close()
        return

    print(f"PDF gefunden: {pdf_path}")

    # PDF hochladen
    remote_pdf = f"wochen-challenge/ab_{datum_short}_Wochenchallenge.pdf"
    sftp.put(pdf_path, remote_pdf)
    print(f"[OK] Hochgeladen: {remote_pdf}")

    # XLSX hochladen (falls vorhanden)
    if os.path.exists(xlsx_path):
        remote_xlsx = f"wochen-challenge/ab_{datum_short}_Wochenchallenge.xlsx"
        sftp.put(xlsx_path, remote_xlsx)
        print(f"[OK] Hochgeladen: {remote_xlsx}")
    else:
        print(f"[WARN] XLSX nicht gefunden, nur PDF hochgeladen.")

    # Neue Vorpunkte berechnen (altes Total + diese Woche)
    for grp, members in gruppen.items():
        for name, tage in members.items():
            key = f"{grp}_{name}"
            old = all_vorpunkte.get(key, 0)
            all_vorpunkte[key] = old + len(tage)

    wc_state["last_processed_week_start"] = week_start_short
    wc_state["last_email_hash"]           = email_hash   # NEU: E-Mail-Hash speichern
    wc_state["alle_vorpunkte"]            = all_vorpunkte
    wc_state["use_friday_start"]          = False  # nach Verarbeitung immer zurücksetzen
    save_wc_state(sftp, wc_state)

    sftp.close()
    ssh.close()

    # Bestätigungs-WhatsApp
    trainees = [f"{grp} {name}: {len(tage)} Pkt"
                for grp, members in gruppen.items()
                for name, tage in members.items() if tage]
    send_whatsapp(
        f"Hi Noah, Cloude hier 🏆\n\n"
        f"Wochenchallenge {start_datum} – {end_datum} ist fertig und hochgeladen!\n\n"
        f"Punkte diese Woche:\n" +
        "\n".join(f"  {t}" for t in trainees) +
        f"\n\ntv-rheinzabern.e-websolutions.de"
    )

    print("\nFERTIG!")

if __name__ == "__main__":
    main()
