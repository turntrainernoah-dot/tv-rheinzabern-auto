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
Bsp:  13.6 Vorname G1
      14.6 Vorname
      14.6 Vorname G3 für gestern
      15.6 Vorname 💪
(reale Namen/Aliase stehen nicht mehr im Code, siehe NAME_MAP/apply_name_aliases())
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

# --- Wochenchallenge-Quelle: "email" (Mail) oder "chat" (admin.php WhatsApp-Chat + KI) ---
WC_SOURCE         = os.environ.get("WC_SOURCE", "email")
GH_MODELS_TOKEN   = os.environ.get("GH_MODELS_TOKEN", "") or os.environ.get("GITHUB_TOKEN", "")
CHAT_REMOTE       = os.environ.get("WC_CHAT_REMOTE", "wc_chat_input.json")

# ════════════════════════════════════════════════════════════════
#  NAME-MAPPING (Phase 2, 24.08.2026: aus dem oeffentlichen Repo ausgelagert)
#  Schlüssel: lowercase wie er im Mail vorkommen kann
#  Wert: (Gruppe, interner Name im WC-Script)
#
#  NAME_MAP/AMBIGUOUS_NAMES/WC_GRUPPEN_TEMPLATE/DEFAULT_VORPUNKTE enthielten
#  bisher die echten Klarnamen/Spitznamen aller Kinder hartkodiert im
#  oeffentlichen Repo. Seit Phase 2 (24.08.2026) sind sie LEER und werden zur
#  Laufzeit ausschliesslich aus dem geschuetzten Server-Ordner geladen
#  (apply_name_aliases()/apply_config_roster(), siehe unten). Es gibt bewusst
#  KEINEN Namens-Fallback mehr: kann die Datei nicht geladen werden, bricht
#  main() klar ab (siehe REQUIRE_SERVER_ROSTER-Check), statt mit einer leeren
#  oder veralteten Namensliste weiterzulaufen und z.B. WC-Mails mit "Unbekannter
#  Name" fuer JEDES Kind abzulehnen.
# ════════════════════════════════════════════════════════════════

NAME_MAP = {}

# Mehrdeutige Namen (brauchen Gruppenangabe) – wird aus name_aliases.json befuellt
AMBIGUOUS_NAMES = set()

# Alle Namen in der richtigen Reihenfolge für den WC-Script (neues Format)
# – wird aus config/config.json befuellt (apply_config_roster)
WC_GRUPPEN_TEMPLATE = {}

# Name-Normalisierung fuer alte abmeldungen.json-Formate (siehe apply_name_aliases);
# wird zusaetzlich in den generierten create_wochenchallenge_v2-Lauf injiziert.
NAME_NORMALIZE = {}

# Initiale Vorpunkte – NUR Legacy-Fallback, falls wc_state_auto.json auf dem
# Server jemals "alle_vorpunkte" leer/fehlend haette. Im Normalbetrieb nie
# aktiv, weil das WC-Punkte-Ledger (wc_punkte_ledger.json) seit 01.07.2026 die
# tatsaechliche Wahrheitsquelle ist (siehe Vault: WC-Punkte-Ledger-System).
# Bewusst leer statt mit den alten Stand-Juni-2026-Werten je Kind.
DEFAULT_VORPUNKTE = {}

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

# ════════════════════════════════════════════════════════════════
#  KEY-MIGRATION: alte Format-Keys → neue Format-Keys
#  Phase 2 (24.08.2026): aus dem oeffentlichen Repo ausgelagert. War eine
#  EINMALIGE Migrationstabelle fuer den Namensformat-Wechsel im Juni 2026
#  (z.B. "G1_Vorname" -> "G1_Vorname X."). Der Server-Zustand ist seit Monaten
#  durchgehend im neuen Format (siehe Vault: WC-Punkte-Ledger-System) -- diese
#  Tabelle ist im Normalbetrieb bereits seit langem ein No-Op. Wird jetzt
#  optional aus config/name_aliases.json ("old_key_map") geladen; bleibt sie
#  leer, migriert migrate_vorpunkte_keys() einfach nichts (unveraendertes
#  Verhalten fuer alle Keys, die ohnehin schon im neuen Format sind).
# ════════════════════════════════════════════════════════════════

OLD_KEY_MAP = {}

def migrate_vorpunkte_keys(alle_vorpunkte):
    """Migriert alte Vorpunkte-Keys (z.B. 'G1_Vorname') zu neuen Keys ('G1_Vorname X.')."""
    migrated = {}
    changed = 0
    for key, val in alle_vorpunkte.items():
        new_key = OLD_KEY_MAP.get(key, key)
        if new_key != key:
            changed += 1
        migrated[new_key] = migrated.get(new_key, 0) + val  # addieren falls doppelter Key
    if changed > 0:
        print(f"[MIGRATION] {changed} Vorpunkte-Keys auf neues Format migriert.")
    return migrated

def _load_server_ledger(sftp):
    import json
    try:
        f = sftp.open("wc_punkte_ledger.json", "rb"); d = json.loads(f.read().decode("utf-8")); f.close(); return d
    except Exception as e:
        print("[WARN] Ledger nicht ladbar:", e); return {"init_offset": {}, "weeks": {}, "order": []}

def _ledger_base_vor(ledger, cname, week_iso):
    tot = ledger.get("init_offset", {}).get(cname, 0)
    for iso, wd in ledger.get("weeks", {}).items():
        if iso < week_iso:
            tot += wd.get("results", {}).get(cname, 0)
    return tot

def load_wc_state(sftp):
    try:
        f     = sftp.open("wc_state_auto.json", "r")
        state = json.loads(f.read().decode("utf-8"))
        f.close()
        print("wc_state_auto.json geladen.")
        # Alte Keys migrieren falls nötig
        if state.get("alle_vorpunkte"):
            state["alle_vorpunkte"] = migrate_vorpunkte_keys(state["alle_vorpunkte"])
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

def send_email(text):
    import os, smtplib, ssl as _ssl
    from email.mime.text import MIMEText
    user = os.environ.get("GMAIL_USER", "")
    pw   = os.environ.get("GMAIL_APP_PASSWORD", "")
    to   = os.environ.get("EMAIL_TO", "") or user
    if not user or not pw:
        print("[MAIL] keine Gmail-Zugangsdaten - uebersprungen.")
        return
    try:
        lines = [l for l in text.strip().splitlines() if l.strip()]
        subj = "TV Rheinzabern: " + (lines[0][:80] if lines else "Info")
        msg = MIMEText(text, _charset="utf-8")
        msg["Subject"] = subj
        msg["From"] = user
        msg["To"] = to
        ctx = _ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ctx) as s:
            s.login(user, pw)
            s.sendmail(user, [to], msg.as_string())
        print("[MAIL] gesendet an", to)
    except Exception as e:
        print("[MAIL] Fehler:", e)


def send_whatsapp(text):
    """Sendet eine WhatsApp-Nachricht via CallMeBot (kostenlos)."""
    send_email(text)
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

    # Wochentag am Ende (z.B. "Vorname Montag" → Vorname trainierte am Montag dieser Woche)
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
        # Klammern normalisieren: "Vorname (G1)" → "vorname g1", "Vorname (G3)" → "vorname g3"
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

def build_config_block(start_datum, end_datum, tage_header_raw, gruppen, vorherige, name_normalize=None):
    """
    Erstellt den Python-Quellcode-String für den CONFIG-Block.
    tage_header_raw: list of strings like "Sa\n06.06." (mit echtem Newline-Char)

    name_normalize (Phase 2, 24.08.2026): wird zusaetzlich als NAME_NORMALIZE
    injiziert (frueher hartkodiert in create_wochenchallenge_v2.py, jetzt aus
    config/name_aliases.json geladen -- siehe apply_name_aliases()). Default
    None -> globales NAME_NORMALIZE dieses Moduls.
    """
    if name_normalize is None:
        name_normalize = NAME_NORMALIZE

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

    # NAME_NORMALIZE als Python-Dict (ersetzt die frueher hartkodierte Kopie
    # in create_wochenchallenge_v2.py::_load_abwesend())
    normalize_json = json.dumps(name_normalize, indent=4, ensure_ascii=False)

    return (
        f'START_DATUM = "{start_datum}"\n'
        f'END_DATUM   = "{end_datum}"\n'
        f'\n'
        f'TAGE_HEADER = {tage_header_src}\n'
        f'\n'
        f'GRUPPEN = {gruppen_json}\n'
        f'\n'
        f'VORHERIGE_PUNKTE = {vorherige_json}\n'
        f'\n'
        f'NAME_NORMALIZE = {normalize_json}\n'
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

def apply_config_roster(sftp):
    """Laedt config/config.json (von admin.php gepflegt) und befuellt
    WC_GRUPPEN_TEMPLATE. Seit Phase 2 (24.08.2026) gibt es KEINE hartkodierte
    Namensliste mehr im Repo -- bei jedem Fehler bleibt WC_GRUPPEN_TEMPLATE
    leer, und main() bricht danach ueber den REQUIRE_SERVER_ROSTER-Check klar
    ab, statt mit leerem/falschem Roster weiterzulaufen."""
    global WC_GRUPPEN_TEMPLATE
    try:
        f = sftp.open("config/config.json", "r")
        cfg = json.loads(f.read().decode("utf-8", errors="replace"))
        f.close()
    except Exception as e:
        print(f"[CONFIG] config.json nicht ladbar ({e!r}).")
        return
    try:
        # Seit dem Gruppenzeiten-Umbau (25.08.2026) ist "gruppen" ein Array
        # von {"name":..,"zeiten":{...}} statt Strings (siehe
        # auto_trainingsplan.py::apply_config_roster) -- hier reicht der
        # Name, die Zeiten braucht nur der Trainingsplan-Generator.
        gruppen_raw = cfg.get("gruppen") or ["G1", "G2", "G3", "G4"]
        gruppen = [g["name"] if isinstance(g, dict) else g for g in gruppen_raw]
        tmpl = {g: [] for g in gruppen}
        for p in cfg.get("personen", []):
            if p.get("rolle") != "turner":
                continue
            ni = p.get("name_intern") or p.get("anzeige")
            g  = p.get("gruppe")
            if ni and g is not None:
                tmpl.setdefault(g, []).append(ni)
        if not any(tmpl.values()):
            print("[CONFIG] config.json ohne Turner.")
            return
        WC_GRUPPEN_TEMPLATE = tmpl
        print(f"[CONFIG] WC-Roster aus config.json: {sum(len(v) for v in tmpl.values())} Turner.")
    except Exception as e:
        print(f"[CONFIG] Fehler beim Aufbau ({e!r}).")


def apply_name_aliases(sftp):
    """Laedt config/name_aliases.json (geschuetzter Server-Ordner, gleiche
    .htaccess wie config.json) und befuellt NAME_MAP, AMBIGUOUS_NAMES und
    NAME_NORMALIZE. Ersetzt die frueher hartkodierte NAME_MAP (Spitznamen/
    Schreibvarianten je Kind fuer die WC-Mail-/Chat-Erkennung). Kein
    Namens-Fallback mehr im Code -- bei Fehler bleiben alle drei Strukturen
    leer, main() bricht danach klar ab (REQUIRE_SERVER_ROSTER-Check)."""
    global NAME_MAP, AMBIGUOUS_NAMES, NAME_NORMALIZE, OLD_KEY_MAP
    try:
        f = sftp.open("config/name_aliases.json", "r")
        cfg = json.loads(f.read().decode("utf-8", errors="replace"))
        f.close()
    except Exception as e:
        print(f"[CONFIG] name_aliases.json nicht ladbar ({e!r}).")
        return
    try:
        aliases = cfg.get("aliases") or {}
        nm = {k.lower(): tuple(v) for k, v in aliases.items() if isinstance(v, (list, tuple)) and len(v) == 2}
        if not nm:
            print("[CONFIG] name_aliases.json ohne aliases.")
            return
        NAME_MAP        = nm
        AMBIGUOUS_NAMES = set(cfg.get("ambiguous") or [])
        NAME_NORMALIZE  = dict(cfg.get("normalize") or {})
        OLD_KEY_MAP     = dict(cfg.get("old_key_map") or {})
        print(f"[CONFIG] Name-Aliase aus name_aliases.json: {len(nm)} Eintraege, "
              f"{len(AMBIGUOUS_NAMES)} mehrdeutig, {len(NAME_NORMALIZE)} Normalize-Eintraege, "
              f"{len(OLD_KEY_MAP)} Alt-Format-Migrationen.")
    except Exception as e:
        print(f"[CONFIG] Fehler beim Aufbau der Name-Aliase ({e!r}).")


def require_server_roster():
    """Harter Abbruch, wenn Roster ODER Name-Aliase nicht vom Server geladen
    werden konnten. Seit Phase 2 gibt es keine hartkodierten Namen mehr im
    Code, die als Fallback einspringen koennten -- ein leerer/kein Server-
    Zustand darf NICHT stillschweigend zu einer falschen/leeren WC-Liste
    fuehren, sondern muss den Lauf klar abbrechen (siehe Vault: GitHub-Umzug,
    Phase 2)."""
    problems = []
    if not any(WC_GRUPPEN_TEMPLATE.values()):
        problems.append("WC_GRUPPEN_TEMPLATE leer (config/config.json nicht ladbar/leer)")
    if not NAME_MAP:
        problems.append("NAME_MAP leer (config/name_aliases.json nicht ladbar/leer)")
    if problems:
        msg = ("Wochenchallenge-Lauf abgebrochen: Roster/Namensliste konnte nicht vom "
               "Server geladen werden -- " + "; ".join(problems) + ". "
               "Es gibt keinen Namens-Fallback im Code mehr (Phase 2, 24.08.2026), "
               "damit die Automatik nie mit falschen/leeren Daten weiterlaeuft.")
        print(f"[FATAL] {msg}")
        try:
            send_whatsapp(f"Hi Noah, Cloude hier 🚨\n\n{msg}")
        except Exception:
            pass
        raise SystemExit(msg)


def fetch_chat_body(sftp):
    """Chat-Modus: liest den rohen WhatsApp-Chat (von admin.php abgelegt) vom
    Server und laesst ihn von der KI (GitHub Models) in das saubere
    parse_email_body-Format umwandeln. Gibt (clean_body, info) oder (None, None)."""
    try:
        f = sftp.open(CHAT_REMOTE, "r")
        raw = f.read().decode("utf-8"); f.close()
    except Exception as e:
        print(f"[CHAT] Keine {CHAT_REMOTE} auf dem Server ({e!r}).")
        return None, None
    try:
        payload = json.loads(raw)
        raw_chat = payload.get("chat", "") if isinstance(payload, dict) else str(raw)
    except Exception:
        raw_chat = raw
    if not raw_chat.strip():
        print("[CHAT] Chat-Datei leer."); return None, None
    try:
        with sftp.open("wc_last_run.txt", "w") as _rf:
            _rf.write(f"{date.today()} chat-mode erreicht, chat_len={len(raw_chat)}, "
                      f"token_gesetzt={bool(GH_MODELS_TOKEN)}")
    except Exception:
        pass
    # Erst KI-Auswertung, DANN als .processed markieren. Bei KI-Fehler bleibt der
    # Chat erhalten und kann erneut verarbeitet werden (kein Datenverlust).
    import wc_chat_parser
    try:
        body, info = wc_chat_parser.chat_to_clean_body(
            raw_chat, NAME_MAP, WC_GRUPPEN_TEMPLATE, token=GH_MODELS_TOKEN)
    except Exception as e:
        print(f"[CHAT] KI-Auswertung fehlgeschlagen: {e!r} -- Chat bleibt erhalten.")
        try:
            with sftp.open("wc_last_error.txt", "w") as _ef:
                _ef.write(f"{date.today()} KI-Fehler: {e!r}")
        except Exception:
            pass
        return None, None
    if not body or not body.strip():
        print("[CHAT] KI lieferte keine verwertbaren Zeilen -- Chat bleibt erhalten.")
        try:
            with sftp.open("wc_last_error.txt", "w") as _ef:
                _ef.write(f"{date.today()} LEERER Body. info={info!r}")
        except Exception:
            pass
        return None, None
    try:
        try: sftp.remove(CHAT_REMOTE + ".processed")
        except Exception: pass
        sftp.rename(CHAT_REMOTE, CHAT_REMOTE + ".processed")
    except Exception as e:
        print(f"[CHAT] Konnte {CHAT_REMOTE} nicht umbenennen: {e!r}")
    print("[CHAT] KI-bereinigter Body:\n" + body)
    return body, info


def main():
    print(f"=== Auto-Wochenchallenge | {date.today()} ===\n")

    print(f"[MODE] WC_SOURCE = {WC_SOURCE}")

    if WC_SOURCE != "chat" and (not GMAIL_USER or not GMAIL_APP_PASSWORD):
        print("[ERROR] GMAIL_USER oder GMAIL_APP_PASSWORD nicht gesetzt.")
        sys.exit(1)

    # SFTP verbinden
    print("Verbinde SFTP...")
    ssh, sftp = get_sftp()
    apply_config_roster(sftp)
    apply_name_aliases(sftp)
    require_server_roster()
    wc_state  = load_wc_state(sftp)

    if WC_SOURCE == "chat":
        print("Lese WhatsApp-Chat (admin.php) + KI-Auswertung (GitHub Models)...")
        body, chat_info = fetch_chat_body(sftp)
        if not body:
            print("Kein verwertbarer Chat gefunden.")
            sftp.close(); ssh.close()
            return
        mail_id = b"chat"
    else:
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
    ledger = _load_server_ledger(sftp)
    week_iso = week_start.strftime("%Y-%m-%d")
    if not wc_state.get("alle_vorpunkte"):
        print("[WARN] alle_vorpunkte leer/fehlend im State → verwende DEFAULT_VORPUNKTE")
    vorherige     = {}
    for grp, members in gruppen.items():
        for name in members:
            key = f"{grp}_{name}"
            vorherige[key] = _ledger_base_vor(ledger, name, week_iso)

    # TAGE_HEADER
    tage_header = []
    for i, abbr in enumerate(wday_abbr):
        d_str = (week_start + timedelta(days=i)).strftime("%d.%m.")
        tage_header.append(f"{abbr}\n{d_str}")  # echtes Newline

    # abmeldungen.json für Abwesenheits-Rot-Markierung herunterladen
    abm_local = "/tmp/abmeldungen.json"
    try:
        sftp.get("abmeldungen/abmeldungen.json", abm_local)
        print(f"[OK] abmeldungen.json heruntergeladen → {abm_local}")
    except Exception as e:
        print(f"[WARN] abmeldungen.json konnte nicht geladen werden: {e} → keine Abwesenheitsmarkierung")
        # Leere Datei erstellen damit _load_abwesend() nicht abstürzt
        with open(abm_local, "w", encoding="utf-8") as fj:
            fj.write("[]")

    # murmel_punkte_state.json fuer Insgesamt=aktuell-verdient (Board-Reset-Stand, Phase 2)
    try:
        sftp.get("murmel_punkte_state.json", "/tmp/murmel_punkte_state.json")
        print("[OK] murmel_punkte_state.json heruntergeladen -> /tmp")
    except Exception as e:
        print(f"[WARN] murmel_punkte_state.json fehlt: {e} -> Insgesamt = voller Ledger-Stand")

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

    # Woche ins zentrale LEDGER schreiben (idempotent) + hochladen -> Board/"letzte Woche"
    week_iso = week_start.strftime("%Y-%m-%d")
    _results = {}
    for grp, members in gruppen.items():
        for name, tage in members.items():
            _results[name] = len(tage)
    ledger.setdefault("weeks", {})[week_iso] = {
        "label": week_start.strftime("%d.%m"),
        "start": week_start.strftime("%d.%m.%Y"),
        "end": (week_start + timedelta(days=num_days - 1)).strftime("%d.%m.%Y"),
        "tage": num_days, "results": _results}
    ledger["order"] = sorted(ledger["weeks"].keys())
    import json as _json
    with open("/tmp/wc_punkte_ledger.json", "w", encoding="utf-8") as _lf:
        _json.dump(ledger, _lf, ensure_ascii=False, indent=1)
    sftp.put("/tmp/wc_punkte_ledger.json", "wc_punkte_ledger.json")
    print(f"[OK] Ledger aktualisiert + hochgeladen (Woche {week_iso})")

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
