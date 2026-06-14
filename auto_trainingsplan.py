#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TV Rheinzabern – Automatischer Trainingsplan Generator
=======================================================
Läuft via GitHub Actions alle 30 Minuten.
PC muss NICHT an sein.

Ablauf:
  check_quick.py → needs_update=true → DIESER SCRIPT läuft
  - Plan vorhanden + Abwesenheitsänderung → Plan aktualisieren (Trainer unverändert)
  - Kein Plan + Veröffentlichungsfenster (Mi/Fr 22:00 CEST) → Neuen Plan erstellen
  - Trainer-Wechsel → E-Mail, manuell prüfen
"""

import json, hashlib, os, re, sys, smtplib, tempfile, shutil, subprocess
from datetime import date, timedelta, datetime, timezone
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import paramiko
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

# ════════════════════════════════════════════════════════════════
#  STATISCHE KONFIGURATION
# ════════════════════════════════════════════════════════════════

ALLE_TURNER = {
    "G1": ["Felix G1", "Finn G1", "Sinan", "Ilyas", "Jonathan", "Hannes", "Ben G1"],
    "G2": ["Henry", "Matti", "Levent", "Caius"],
    "G3": ["Erik", "Artem", "Finn G3", "Ben G3", "Michael"],
    "G4": ["Felix G4", "Anton", "Mika", "Jamie"],
}
ALLE_TRAINER = ["Noah", "Andy", "Fabian", "Cassian", "Julian", "Torben"]

GERAETE_ROTATION = [
    ("Boden", "Barren"),
    ("Sprung", "Reck"),
    ("Seitpferd", "Ringe"),
]

ZEITSLOTS = [
    "17:00–17:30",
    "17:30–18:00",
    "18:00–18:15",
    "18:15–19:00",
    "19:00–19:30",
]

FARBEN = {
    "g1_blau":        "2471A3",
    "g1_gruen":       "1E8449",
    "g2_orange":      "CA6F1E",
    "g2_lila":        "7D3C98",
    "aufwaermen":     "A569BD",
    "aufbauen":       "717D7E",
    "springer":       "A0522D",
    "abwesend":       "D5D8DC",
    "sonder":         "E74C3C",
    "titel":          "2C3E50",
    "header":         "5D6D7E",
    "anwesend":       "1E8449",
    "abwesend_turner":"C0392B",
    "anmerkung":      "D6EAF8",
    "legende_bg":     "2C3E50",
}

# ════════════════════════════════════════════════════════════════
#  UMGEBUNGSVARIABLEN (GitHub Secrets)
# ════════════════════════════════════════════════════════════════

SSH_HOST      = os.environ.get("SSH_HOST",     "access-5017462830.webspace-host.com")
SSH_USER      = os.environ.get("SSH_USER",     "a2358459")
SSH_PASSWORD  = os.environ.get("SSH_PASSWORD", "")
SSH_PORT      = int(os.environ.get("SSH_PORT", "22"))
GMAIL_USER         = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
EMAIL_TO      = os.environ.get("EMAIL_TO", "turntrainernoah@gmail.com")

# ════════════════════════════════════════════════════════════════
#  STATE (gespeichert auf dem Server als state_auto.json)
# ════════════════════════════════════════════════════════════════

DEFAULT_STATE = {
    "last_training_date": "17.06.2026",
    "geraet_combo_index": 1,      # 0=Boden+Barren, 1=Sprung+Reck, 2=Seitpferd+Ringe
    "g1_starts_geraet2": True,
    "generated_plans":   ["10.06.26", "17.06.26"],
    "plan_data":         {},
}

def load_state(sftp):
    try:
        f    = sftp.open("state_auto.json", "r")
        data = json.loads(f.read().decode("utf-8"))
        f.close()
        print("State von Server geladen.")
        return data
    except Exception:
        print("Kein State auf Server – verwende Default.")
        return DEFAULT_STATE.copy()

def save_state(sftp, state):
    data = json.dumps(state, indent=2, ensure_ascii=False).encode("utf-8")
    f    = sftp.open("state_auto.json", "wb")
    f.write(data)
    f.close()
    print("State auf Server gespeichert.")

# ════════════════════════════════════════════════════════════════
#  HILFSFUNKTIONEN
# ════════════════════════════════════════════════════════════════

def compute_absences_hash(absences):
    return hashlib.md5(
        json.dumps(absences, sort_keys=True, ensure_ascii=False).encode("utf-8")
    ).hexdigest()

def is_publication_window():
    """Veröffentlichungsfenster: Mi oder Fr nach 20:00 UTC (= 22:00 CEST)."""
    now = datetime.now(timezone.utc)
    return now.weekday() in (2, 4) and now.hour >= 20

def next_training_date():
    """Nächster Trainingstag (Mi/Fr) ab morgen."""
    d = date.today() + timedelta(days=1)
    while d.weekday() not in (2, 4):
        d += timedelta(days=1)
    return d

def fmt_datum(d):
    return d.strftime("%d.%m.%Y")

def fmt_datum_kurz(d):
    return d.strftime("%d.%m.%y")

def wochentag_name(d):
    return "Mittwoch" if d.weekday() == 2 else "Freitag"

# ════════════════════════════════════════════════════════════════
#  SFTP
# ════════════════════════════════════════════════════════════════

def get_sftp():
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    client.connect(SSH_HOST, port=SSH_PORT, username=SSH_USER,
                   password=SSH_PASSWORD, timeout=20)
    return client, client.open_sftp()

def plan_exists(sftp, datum_kurz):
    try:
        sftp.stat(f"trainingspläne/{datum_kurz}_Trainingsplan.pdf")
        return True
    except FileNotFoundError:
        return False

def normalize_name(name):
    return re.sub(r'\s*\(G(\d+)\)$', r' G\1', name.strip())

def read_abmeldungen(sftp):
    f    = sftp.open("abmeldungen/abmeldungen.json", "r")
    raw  = f.read()
    f.close()
    data = json.loads(raw.decode("utf-8", errors="replace"))
    raw_hash = hashlib.md5(raw).hexdigest()
    return data, raw_hash

def read_anmerkungen_server(sftp):
    """Liest ungelesene Trainer-Anmerkungen vom Server."""
    try:
        f    = sftp.open("anmerkungen/anmerkungen.json", "r")
        data = json.loads(f.read().decode("utf-8", errors="replace"))
        f.close()
        return [a for a in data if not a.get("gelesen", False)]
    except Exception:
        return []

def mark_anmerkungen_gelesen(sftp, ids_to_mark):
    """Markiert angegebene Anmerkungen als gelesen auf dem Server."""
    if not ids_to_mark:
        return
    try:
        f    = sftp.open("anmerkungen/anmerkungen.json", "r")
        data = json.loads(f.read().decode("utf-8", errors="replace"))
        f.close()
        for entry in data:
            if entry.get("id") in ids_to_mark:
                entry["gelesen"] = True
        f = sftp.open("anmerkungen/anmerkungen.json", "w")
        f.write(json.dumps(data, indent=2, ensure_ascii=False).encode("utf-8"))
        f.close()
        print(f"[OK] {len(ids_to_mark)} Anmerkung(en) als gelesen markiert.")
    except Exception as e:
        print(f"[WARN] Anmerkungen konnten nicht markiert werden: {e}")

def upload_pdf(sftp, local_path, datum_kurz):
    remote = f"trainingspläne/{datum_kurz}_Trainingsplan.pdf"
    sftp.put(local_path, remote)
    print(f"[OK] Hochgeladen: {remote}")

def upload_xlsx(sftp, local_path, datum_kurz):
    remote = f"trainingspläne/{datum_kurz}_Trainingsplan.xlsx"
    sftp.put(local_path, remote)
    print(f"[OK] Hochgeladen: {remote}")

# ════════════════════════════════════════════════════════════════
#  ABWESENHEITEN AUSWERTEN
# ════════════════════════════════════════════════════════════════

def get_absences(abmeldungen, training_date):
    target    = training_date.strftime("%Y-%m-%d")
    absences  = {"G1": [], "G2": [], "G3": [], "G4": [], "Trainer": []}
    late_notes = []

    for entry in abmeldungen:
        if entry.get("datum") != target:
            continue
        name   = normalize_name(entry.get("name", "").strip())
        gruppe = entry.get("gruppe", "").strip()
        notiz  = entry.get("notiz", "").strip()

        if notiz and any(k in notiz.lower() for k in ["später", "verspät", "kommt", "geht"]):
            late_notes.append(f"• {name}: {notiz}")
            continue

        if gruppe in absences:
            absences[gruppe].append(name)

    return absences, late_notes

# ════════════════════════════════════════════════════════════════
#  KOMPLEXE FÄLLE ERKENNEN
# ════════════════════════════════════════════════════════════════

ALLE_BEKANNTEN_NAMES = set()
for _g, _lst in ALLE_TURNER.items():
    ALLE_BEKANNTEN_NAMES.update(_lst)
ALLE_BEKANNTEN_NAMES.update(ALLE_TRAINER)

def detect_complex(absences, late_notes):
    issues = []
    anwesend_trainer = [t for t in ALLE_TRAINER if t not in absences.get("Trainer", [])]
    n = len(anwesend_trainer)

    if n <= 2:
        issues.append(
            f"Nur {n} Trainer anwesend: {', '.join(anwesend_trainer) or '–'}. "
            "Automatische Einteilung nicht möglich."
        )

    for gruppe, names in absences.items():
        for name in names:
            if name not in ALLE_BEKANNTEN_NAMES:
                issues.append(f"Unbekannter Name: '{name}' (Gruppe: {gruppe})")

    for gruppe, turner in ALLE_TURNER.items():
        abw = absences.get(gruppe, [])
        if len(abw) >= len(turner):
            issues.append(f"Gruppe {gruppe} hat keine anwesenden Turner (alle {len(abw)} abwesend)!")

    return issues, anwesend_trainer

# ════════════════════════════════════════════════════════════════
#  TRAINER-EINTEILUNG
# ════════════════════════════════════════════════════════════════

def build_trainer_plan(absences, geraet_1, geraet_2, g1_starts_geraet2):
    abwesend  = absences.get("Trainer", [])
    available = [t for t in ALLE_TRAINER if t not in abwesend]
    n = len(available)

    merge_g23 = (n == 3)

    assignment = {}
    pool = list(available)

    if "Noah" in pool:
        assignment["Noah"] = "G1"
        pool.remove("Noah")
    else:
        assignment[pool.pop(0)] = "G1"

    if "Andy" in pool:
        g4_trainer = "Andy"
        pool.remove("Andy")
    elif pool:
        g4_trainer = pool.pop(-1)
    else:
        g4_trainer = None

    if g4_trainer:
        assignment[g4_trainer] = "G4"

    if merge_g23:
        if pool:
            assignment[pool.pop(0)] = "G2+G3"
    else:
        for gruppe in ("G2", "G3"):
            if pool:
                assignment[pool.pop(0)] = gruppe
        for t in pool:
            assignment[t] = "Springer"

    def farbe(gruppe, phase):
        if g1_starts_geraet2:
            if phase == 1:
                return {"G1":"g2_orange","G2":"g1_blau","G3":"g1_blau",
                        "G2+G3":"g1_blau"}.get(gruppe,"aufbauen")
            else:
                return {"G1":"g1_blau","G2":"g1_gruen","G3":"g2_orange",
                        "G2+G3":"g2_lila"}.get(gruppe,"aufbauen")
        else:
            if phase == 1:
                return {"G1":"g1_blau","G2":"g1_gruen","G3":"g2_orange",
                        "G2+G3":"g2_lila"}.get(gruppe,"aufbauen")
            else:
                return {"G1":"g2_orange","G2":"g2_lila","G3":"g1_blau",
                        "G2+G3":"g1_blau"}.get(gruppe,"aufbauen")

    G4_SLOT3 = "g1_gruen"
    G4_SLOT4 = "g2_orange"

    TRAINER_PLAN = {}
    for trainer in ALLE_TRAINER:
        if trainer in abwesend:
            TRAINER_PLAN[trainer] = None
            continue
        if trainer not in assignment:
            TRAINER_PLAN[trainer] = None
            continue

        grp = assignment[trainer]

        if grp == "G1":
            TRAINER_PLAN[trainer] = [
                ("AW G1",    "aufwaermen"),
                ("G1",       farbe("G1", 1)),
                ("G1",       farbe("G1", 1)),
                ("G1",       farbe("G1", 2)),
                ("Abbauen",  "aufbauen"),
            ]
        elif grp == "G2+G3":
            TRAINER_PLAN[trainer] = [
                ("AW G2+G3", "aufwaermen"),
                ("G2+G3",    farbe("G2+G3", 1)),
                ("G2+G3",    farbe("G2+G3", 1)),
                ("G2+G3",    farbe("G2+G3", 2)),
                ("Abbauen",  "aufbauen"),
            ]
        elif grp == "G2":
            TRAINER_PLAN[trainer] = [
                ("AW G2",    "aufwaermen"),
                ("G2",       farbe("G2", 1)),
                ("G2",       farbe("G2", 1)),
                ("G2",       farbe("G2", 2)),
                ("Abbauen",  "aufbauen"),
            ]
        elif grp == "G3":
            TRAINER_PLAN[trainer] = [
                ("AW G3",    "aufwaermen"),
                ("G3",       farbe("G3", 1)),
                ("G3",       farbe("G3", 1)),
                ("G3",       farbe("G3", 2)),
                ("Abbauen",  "aufbauen"),
            ]
        elif grp == "G4":
            TRAINER_PLAN[trainer] = [
                ("Aufbauen", "aufbauen"),
                ("AW G4",    "aufwaermen"),
                ("AW G4",    "aufwaermen"),
                ("G4",       G4_SLOT3),
                ("G4",       G4_SLOT4),
            ]
        elif grp == "Springer":
            TRAINER_PLAN[trainer] = [
                ("Aufbauen", "aufbauen"),
                ("Springer", "springer"),
                ("Springer", "springer"),
                ("Springer", "springer"),
                ("Abbauen",  "aufbauen"),
            ]

    if "Andy" in available and assignment.get("Andy") == "Springer":
        if TRAINER_PLAN.get("Andy"):
            TRAINER_PLAN["Andy"][4] = ("G4", G4_SLOT4)

    anmerkungen = []
    if "Barren" in (geraet_1, geraet_2):
        anmerkungen.append("• Barren G3: Kippe üben")

    return TRAINER_PLAN, {}, anmerkungen

# ════════════════════════════════════════════════════════════════
#  EXCEL AUFBAUEN
# ════════════════════════════════════════════════════════════════

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color="FFFFFF", size=9):
    return Font(name="Arial", bold=bold, color=color, size=size)

def align(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def set_cell(ws, row, col, value, hex_fill=None, fnt=None, aln=None):
    c = ws.cell(row=row, column=col, value=value)
    if hex_fill: c.fill = fill(hex_fill)
    if fnt:      c.font = fnt
    if aln:      c.alignment = aln
    return c

def merge_set(ws, row, cs, ce, value, hex_fill, fnt, aln=None):
    ws.merge_cells(start_row=row, start_column=cs, end_row=row, end_column=ce)
    c = ws.cell(row=row, column=cs, value=value)
    c.fill = fill(hex_fill)
    c.font = fnt
    c.alignment = aln or align()
    return c

def build_excel(datum, wochentag, geraet_1, geraet_2, abwesend,
                trainer_plan, sondertiming, anmerkungen):
    datum_kurz = datum[0:2] + "." + datum[3:5] + "." + datum[8:10]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Trainingsplan"

    for col in "ABCDEFGHI":
        ws.column_dimensions[col].width = 13

    row = 1

    merge_set(ws, row, 2, 8,
        f"TRAININGSPLAN | {wochentag}, {datum}",
        FARBEN["titel"], font(bold=True, size=13), align(h="left"))
    row += 1

    merge_set(ws, row, 2, 8,
        f"Gerät 1 = {geraet_1} | Gerät 2 = {geraet_2}",
        FARBEN["header"], font(size=10), align(h="left"))
    row += 1

    merge_set(ws, row, 2, 8, "ANWESENHEITEN",
        FARBEN["titel"], font(bold=True, size=10))
    row += 1

    for col_idx, label in enumerate(["G1","G2","G3","G4","TRAINER"], start=3):
        set_cell(ws, row, col_idx, label, FARBEN["header"], font(bold=True))
    row += 1

    max_len = max(max(len(v) for v in ALLE_TURNER.values()), len(ALLE_TRAINER))
    alt = ["EBF5FB","FFFFFF"]

    for i in range(max_len):
        a = alt[i % 2]
        for col_bg in [2, 8]:
            ws.cell(row=row, column=col_bg).fill = fill("FFFFFF")

        for ci, gruppe in enumerate(["G1","G2","G3","G4"], start=3):
            tlist = ALLE_TURNER[gruppe]
            if i < len(tlist):
                name = tlist[i]
                ab   = name in abwesend.get(gruppe, [])
                set_cell(ws, row, ci,
                         f"{'x' if ab else 'v'} {name}",
                         FARBEN["abwesend_turner"] if ab else FARBEN["anwesend"],
                         font(size=9, color="FFFFFF"), align(h="left"))
            else:
                ws.cell(row=row, column=ci).fill = fill(a)

        if i < len(ALLE_TRAINER):
            name = ALLE_TRAINER[i]
            ab   = name in abwesend.get("Trainer", [])
            set_cell(ws, row, 7,
                     f"{'x' if ab else 'v'} {name}",
                     FARBEN["abwesend_turner"] if ab else FARBEN["anwesend"],
                     font(size=9, color="FFFFFF"), align(h="left"))
        else:
            ws.cell(row=row, column=7).fill = fill(a)
        row += 1

    row += 1

    merge_set(ws, row, 2, 8, "GERAETE-LEGENDE",
        FARBEN["legende_bg"], font(bold=True, size=10))
    row += 1
    items = [
        (geraet_1, FARBEN["g1_blau"]),
        (geraet_1, FARBEN["g1_gruen"]),
        (geraet_2, FARBEN["g2_orange"]),
        (geraet_2, FARBEN["g2_lila"]),
        ("Aufwaermen", FARBEN["aufwaermen"]),
        ("Aufbauen/Abbauen", FARBEN["aufbauen"]),
        ("Springer", FARBEN["springer"]),
    ]
    for ci, (label, hex_c) in enumerate(items, start=2):
        set_cell(ws, row, ci, label, hex_c, font(bold=True, size=9), align(wrap=True))
    row += 1

    row += 1

    merge_set(ws, row, 2, 8, "TRAINER-EINTEILUNG",
        FARBEN["legende_bg"], font(bold=True, size=10))
    row += 1

    # Nur anwesende Trainer in der Einteilung zeigen
    anwesende_trainer = [t for t in ALLE_TRAINER if t not in abwesend.get("Trainer", [])]

    set_cell(ws, row, 2, "Zeit", FARBEN["header"], font(bold=True))
    for ci, trainer in enumerate(anwesende_trainer, start=3):
        set_cell(ws, row, ci, trainer, FARBEN["header"], font(bold=True))
    row += 1

    for slot_idx, slot_label in enumerate(ZEITSLOTS):
        ws.row_dimensions[row].height = 38
        set_cell(ws, row, 2, slot_label, FARBEN["header"], font(bold=True))

        for ci, trainer in enumerate(anwesende_trainer, start=3):
            plan = trainer_plan.get(trainer)
            if plan is None or slot_idx >= len(plan):
                continue

            text, fk = plan[slot_idx]
            if slot_idx == 0 and trainer in sondertiming:
                text = sondertiming[trainer]
                fk   = "sonder"

            tc = "555555" if fk == "abwesend" else "FFFFFF"
            set_cell(ws, row, ci, text, FARBEN.get(fk, "FFFFFF"),
                     font(bold=True, color=tc), align(wrap=True))
        row += 1

    if anmerkungen:
        merge_set(ws, row, 2, 8, "ANMERKUNGEN",
            FARBEN["anmerkung"], font(bold=True, size=9, color="1A5276"))
        row += 1
        for anm in anmerkungen:
            merge_set(ws, row, 2, 8, anm,
                FARBEN["anmerkung"], font(size=9, color="1A5276"), align(h="left"))
            row += 1

    # Zurück zur Website
    row += 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
    lc = ws.cell(row=row, column=2, value="Zurück zur Website: tv-rheinzabern.e-websolutions.de")
    lc.hyperlink = "https://tv-rheinzabern.e-websolutions.de/"
    lc.font = Font(name="Arial", size=9, color="0563C1", underline="single")
    lc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 16

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    out_dir   = "/tmp/trainingsplan"
    os.makedirs(out_dir, exist_ok=True)
    xlsx_path = os.path.join(out_dir, f"{datum_kurz}_Trainingsplan.xlsx")
    pdf_path  = os.path.join(out_dir, f"{datum_kurz}_Trainingsplan.pdf")
    wb.save(xlsx_path)
    print(f"[OK] Excel: {xlsx_path}")

    tmp_dir = tempfile.mkdtemp()
    result  = subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf",
         "--outdir", tmp_dir, xlsx_path],
        capture_output=True, text=True, timeout=90
    )
    if result.returncode == 0:
        tmp_pdf = os.path.join(tmp_dir,
            os.path.basename(xlsx_path).replace(".xlsx", ".pdf"))
        if os.path.exists(tmp_pdf):
            shutil.copy2(tmp_pdf, pdf_path)
            print(f"[OK] PDF: {pdf_path}")
        else:
            raise RuntimeError("LibreOffice PDF nicht gefunden")
    else:
        raise RuntimeError(f"LibreOffice Fehler: {result.stderr[:300]}")

    return xlsx_path, pdf_path

# ════════════════════════════════════════════════════════════════
#  E-MAIL
# ════════════════════════════════════════════════════════════════

def send_email(subject, body, to=None):
    recipient = to or EMAIL_TO
    if not GMAIL_USER or not GMAIL_APP_PASSWORD:
        print(f"[EMAIL-TEST] → {recipient}: {subject}")
        return
    msg = MIMEMultipart()
    msg["From"]    = GMAIL_USER
    msg["To"]      = recipient
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        s.send_message(msg)
    print(f"[OK] E-Mail gesendet an {recipient}: {subject}")

# ════════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════════

def main():
    print(f"=== TV Rheinzabern Auto-Trainingsplan | {date.today()} ===\n")

    print("Verbinde mit Server...")
    ssh, sftp = get_sftp()

    state              = load_state(sftp)
    abmeldungen, raw_abm_hash = read_abmeldungen(sftp)
    anmerkungen_server = read_anmerkungen_server(sftp)
    print(f"Abmeldungen geladen: {len(abmeldungen)} Eintraege")
    print(f"Ungelesene Anmerkungen: {len(anmerkungen_server)}")
    print(f"Abmeldungen-Hash (raw): {raw_abm_hash[:8]}...")

    today         = date.today()
    training_date = next_training_date()
    datum         = fmt_datum(training_date)
    datum_kurz    = fmt_datum_kurz(training_date)
    wtag          = wochentag_name(training_date)
    days_away     = (training_date - today).days

    print(f"\nNaechstes Training: {wtag} {datum} ({days_away} Tage)\n")

    # Abwesenheiten und Hash berechnen
    absences, late_notes = get_absences(abmeldungen, training_date)
    # raw_abm_hash für konsistenten Vergleich mit check_quick.py (beide nutzen raw JSON hash)
    new_hash = raw_abm_hash

    if plan_exists(sftp, datum_kurz):
        # ── Plan vorhanden: prüfe ob Update nötig ──────────────
        plan_data   = state.get("plan_data", {}).get(datum_kurz, {})
        stored_hash = plan_data.get("absences_hash", "")

        if not plan_data or not stored_hash:
            # Plan ohne gespeicherten Hash (manuell erstellt oder erste Initialisierung)
            # → Plan mit AKTUELLEN Abwesenheiten NEU ERSTELLEN (nicht nur Hash speichern)
            print(f"Plan ohne Hash fuer {datum_kurz} → Plan wird mit aktuellen Abwesenheiten erstellt.")
            combo_idx = state.get("geraet_combo_index", 0)
            geraet_combo = GERAETE_ROTATION[combo_idx % len(GERAETE_ROTATION)]
            # plan_data mit Defaults initialisieren und stored_hash="" setzen
            # damit die Update-Logik unten sicher ausgeführt wird
            plan_data = {
                "absences_hash":    "",   # leer → abs_changed=True → Plan wird generiert
                "trainer_absences": list(absences.get("Trainer", [])),
                "stored_absences":  {},
                "geraet_1":         geraet_combo[0],
                "geraet_2":         geraet_combo[1],
                "g1_starts_geraet2": state.get("g1_starts_geraet2", False),
            }
            state.setdefault("plan_data", {})[datum_kurz] = plan_data
            if datum_kurz not in state.get("generated_plans", []):
                state.setdefault("generated_plans", []).append(datum_kurz)
            stored_hash = ""   # explizit leer → Update-Logik greift

        has_new_anm = len(anmerkungen_server) > 0
        abs_changed = (new_hash != stored_hash)

        if not abs_changed and not has_new_anm:
            print("Plan aktuell, keine Aenderungen → nichts zu tun.")
            sftp.close(); ssh.close()
            return

        if abs_changed:
            print(f"Abmeldungsaenderung erkannt (Hash {stored_hash[:8]}... → {new_hash[:8]}...)")
        if has_new_anm:
            print(f"Neue Trainer-Anmerkungen: {len(anmerkungen_server)}")

        # Trainer-Absences prüfen
        stored_trainer_abs = set(plan_data.get("trainer_absences", []))
        new_trainer_abs    = set(absences.get("Trainer", []))

        if stored_trainer_abs != new_trainer_abs:
            added   = new_trainer_abs - stored_trainer_abs
            removed = stored_trainer_abs - new_trainer_abs
            body = (
                f"Hallo Noah,\n\n"
                f"die Trainer-Abwesenheiten fuer {wtag}, {datum} haben sich geaendert.\n"
                f"Der Plan kann NICHT automatisch aktualisiert werden – die Trainer-Einteilung\n"
                f"wuerde sich aendern und das muss manuell geprueft werden.\n\n"
                f"Bisher abwesend: {', '.join(sorted(stored_trainer_abs)) or 'Alle da'}\n"
                f"Jetzt abwesend: {', '.join(sorted(new_trainer_abs)) or 'Alle da'}\n"
                f"Neu abwesend: {', '.join(sorted(added)) or '–'}\n"
                f"Nicht mehr abwesend: {', '.join(sorted(removed)) or '–'}\n\n"
                f"Bitte erstelle den Plan manuell in Claude.\n\nGrueße, Auto-Bot"
            )
            send_email(f"Trainer-Aenderung fuer {datum} – manuell pruefen", body)
            print("Trainer-Aenderung! E-Mail gesendet.")
            sftp.close(); ssh.close()
            return

        # UPDATE: gleiche Trainer, neue Abwesenheiten
        print("UPDATE: Trainer unveraendert, erstelle Plan neu...")

        geraet_1     = plan_data["geraet_1"]
        geraet_2     = plan_data["geraet_2"]
        g1_starts_g2 = plan_data["g1_starts_geraet2"]

        issues, anwesend_trainer = detect_complex(absences, late_notes)
        if issues:
            body = (
                f"Hallo Noah,\n\nKonnte Plan fuer {datum} nicht aktualisieren:\n\n"
                + "\n".join(f"  - {i}" for i in issues)
                + f"\n\nBitte manuell pruefen.\n\nGrueße, Auto-Bot"
            )
            send_email(f"Plan-Update {datum} – Fehler", body)
            print("Fehler beim Update! E-Mail gesendet.")
            sftp.close(); ssh.close()
            return

        trainer_plan, sondertiming, anmerkungen = build_trainer_plan(
            absences, geraet_1, geraet_2, g1_starts_g2
        )

        # Verspätungen / frühes Gehen als Hinweis eintragen
        for note in late_notes:
            anmerkungen.append(note)

        # Trainer-Anmerkungen eintragen (Name: Text, ohne Datum)
        for anm in anmerkungen_server:
            trainer_name = anm.get("trainer", "")
            notiz        = anm.get("notiz", "").strip()
            if notiz:
                anmerkungen.append(f"• {trainer_name}: {notiz}")

        xlsx_path, pdf_path = build_excel(
            datum=datum, wochentag=wtag,
            geraet_1=geraet_1, geraet_2=geraet_2,
            abwesend=absences,
            trainer_plan=trainer_plan,
            sondertiming=sondertiming,
            anmerkungen=anmerkungen,
        )
        upload_pdf(sftp, pdf_path, datum_kurz)
        upload_xlsx(sftp, xlsx_path, datum_kurz)

        ids_gelesen = [a["id"] for a in anmerkungen_server if a.get("id")]
        mark_anmerkungen_gelesen(sftp, ids_gelesen)

        # Late notes deduplication: nur neue Hinweise senden
        already_sent_notes = set(plan_data.get("late_notes_sent", []))
        new_late_notes     = [n for n in late_notes if n not in already_sent_notes]

        # State: Hash + late_notes_sent aktualisieren
        plan_data["absences_hash"]    = new_hash
        plan_data["stored_absences"]  = absences
        if new_late_notes:
            plan_data["late_notes_sent"] = list(already_sent_notes | set(new_late_notes))
        save_state(sftp, state)

        # Notification: Trainer-Anmerkungen vorhanden → bitte prüfen
        if anmerkungen_server:
            anm_lines = "\n".join(
                f"  • {a.get('trainer','')}: {a.get('notiz','').strip()}"
                for a in anmerkungen_server if a.get("notiz","").strip()
            )
            send_email(
                f"Trainer-Anmerkung fuer {datum} – bitte pruefen",
                f"Hallo Noah,\n\n"
                f"Es liegt eine neue Trainer-Anmerkung fuer das Training {wtag}, {datum} vor:\n\n"
                f"{anm_lines}\n\n"
                f"Die Anmerkung wurde bereits in den Plan eingebaut und hochgeladen.\n"
                f"Bitte prüfe, ob alles korrekt eingetragen ist.\n\nGrueße, Auto-Bot",
                to="turntrainernoah@gmail.com"
            )

        # Notification: Verspätungen / frühes Gehen → nur NEUE Hinweise
        if new_late_notes:
            send_email(
                f"Verspätung/Frühes Gehen fuer {datum} – bitte pruefen",
                f"Hallo Noah,\n\n"
                f"Fuer das Training {wtag}, {datum} gibt es folgende NEUE Hinweise:\n\n"
                + "\n".join(f"  {n}" for n in new_late_notes) +
                f"\n\nDiese wurden im Plan als Hinweis eingetragen.\n\nGrueße, Auto-Bot",
                to="turntrainernoah@gmail.com"
            )

        # Zusammenfassung E-Mail
        anm_text = ""
        if anmerkungen_server:
            anm_text = f"\nEingebaute Trainer-Anmerkungen:\n" + "\n".join(
                f"  • {a.get('trainer','')}: {a.get('notiz','').strip()}"
                for a in anmerkungen_server if a.get("notiz","").strip()
            ) + "\n"
        send_email(
            f"Plan {datum} aktualisiert",
            f"Hallo Noah,\n\n"
            f"der Trainingsplan fuer {wtag}, {datum} wurde automatisch aktualisiert.\n\n"
            f"Geraete: {geraet_1} + {geraet_2}\n\n"
            f"Abwesend:\n"
            f"  Trainer: {', '.join(absences.get('Trainer','')) or 'Alle da'}\n"
            f"  G1: {', '.join(absences.get('G1','')) or 'Alle da'}\n"
            f"  G2: {', '.join(absences.get('G2','')) or 'Alle da'}\n"
            f"  G3: {', '.join(absences.get('G3','')) or 'Alle da'}\n"
            f"  G4: {', '.join(absences.get('G4','')) or 'Alle da'}\n"
            f"{anm_text}\n"
            f"Grueße, Auto-Bot"
        )
        print(f"UPDATE FERTIG fuer {datum}.")

    else:
        # ── Kein Plan vorhanden: erstelle neuen ────────────────
        # Erlaubt wenn Training ≤5 Tage entfernt (= genug Zeit, aber nicht zu früh)
        now_utc   = datetime.now(timezone.utc)
        days_away = (training_date - date.today()).days
        if days_away > 5 and not is_publication_window():
            print(
                f"Kein Plan vorhanden, Training in {days_away} Tagen "
                f"({now_utc.strftime('%H:%M')} UTC) → warte noch."
            )
            sftp.close(); ssh.close()
            return

        print(f"Kein Plan vorhanden, starte Generierung... (Training in {days_away} Tagen, {now_utc.strftime('%H:%M')} UTC)")

        new_combo_idx = (state.get("geraet_combo_index", 2) + 1) % 3
        geraet_1, geraet_2 = GERAETE_ROTATION[new_combo_idx]
        g1_starts_g2 = not state.get("g1_starts_geraet2", True)

        issues, anwesend_trainer = detect_complex(absences, late_notes)

        if issues:
            body = (
                f"Hallo Noah,\n\n"
                f"der automatische Trainingsplan fuer {wtag}, {datum} "
                f"konnte nicht erstellt werden.\n\n"
                f"Geraete (geplant): {geraet_1} + {geraet_2}\n\n"
                f"Probleme:\n" +
                "\n".join(f"  - {i}" for i in issues) +
                f"\n\nAnwesende Trainer: {', '.join(anwesend_trainer) or '–'}\n"
                f"Abwesend Trainer: {', '.join(absences.get('Trainer', [])) or '–'}\n\n"
                f"Bitte oeffne Claude und erstelle den Plan manuell.\n\nGrueße, Auto-Bot"
            )
            send_email(f"Trainingsplan {datum} – manuelle Pruefung noetig", body)
            print("Komplex! E-Mail gesendet.")
            sftp.close(); ssh.close()
            return

        trainer_plan, sondertiming, anmerkungen = build_trainer_plan(
            absences, geraet_1, geraet_2, g1_starts_g2
        )

        # Verspätungen / frühes Gehen als Hinweis eintragen
        for note in late_notes:
            anmerkungen.append(note)

        # Trainer-Anmerkungen eintragen (Name: Text, ohne Datum)
        for anm in anmerkungen_server:
            trainer_name = anm.get("trainer", "")
            notiz        = anm.get("notiz", "").strip()
            if notiz:
                anmerkungen.append(f"• {trainer_name}: {notiz}")

        xlsx_path, pdf_path = build_excel(
            datum=datum,
            wochentag=wtag,
            geraet_1=geraet_1,
            geraet_2=geraet_2,
            abwesend=absences,
            trainer_plan=trainer_plan,
            sondertiming=sondertiming,
            anmerkungen=anmerkungen,
        )

        upload_pdf(sftp, pdf_path, datum_kurz)
        upload_xlsx(sftp, xlsx_path, datum_kurz)

        ids_gelesen = [a["id"] for a in anmerkungen_server if a.get("id")]
        mark_anmerkungen_gelesen(sftp, ids_gelesen)

        # State aktualisieren
        state["last_training_date"] = datum
        state["geraet_combo_index"] = new_combo_idx
        state["g1_starts_geraet2"]  = g1_starts_g2
        state.setdefault("generated_plans", []).append(datum_kurz)
        state.setdefault("plan_data", {})[datum_kurz] = {
            "absences_hash":    new_hash,
            "trainer_absences": list(absences.get("Trainer", [])),
            "stored_absences":  absences,
            "geraet_1":         geraet_1,
            "geraet_2":         geraet_2,
            "g1_starts_geraet2": g1_starts_g2,
            "late_notes_sent":  list(late_notes),  # track sent late notes
        }
        save_state(sftp, state)

        # Notification: Trainer-Anmerkungen vorhanden → bitte prüfen
        if anmerkungen_server:
            anm_lines = "\n".join(
                f"  • {a.get('trainer','')}: {a.get('notiz','').strip()}"
                for a in anmerkungen_server if a.get("notiz","").strip()
            )
            send_email(
                f"Trainer-Anmerkung fuer {datum} – bitte pruefen",
                f"Hallo Noah,\n\n"
                f"Es liegt eine neue Trainer-Anmerkung fuer das Training {wtag}, {datum} vor:\n\n"
                f"{anm_lines}\n\n"
                f"Die Anmerkung wurde bereits in den Plan eingebaut und hochgeladen.\n"
                f"Bitte prüfe, ob alles korrekt eingetragen ist.\n\nGrueße, Auto-Bot",
                to="turntrainernoah@gmail.com"
            )

        # Notification: Verspätungen / frühes Gehen → bitte prüfen
        if late_notes:
            send_email(
                f"Verspätung/Frühes Gehen fuer {datum} – bitte pruefen",
                f"Hallo Noah,\n\n"
                f"Fuer das Training {wtag}, {datum} gibt es folgende Hinweise:\n\n"
                + "\n".join(f"  {n}" for n in late_notes) +
                f"\n\nDiese wurden im Plan als Hinweis eingetragen.\n\nGrueße, Auto-Bot",
                to="turntrainernoah@gmail.com"
            )

        anm_text = ""
        if anmerkungen_server:
            anm_text = f"\nEingebaute Trainer-Anmerkungen:\n" + "\n".join(
                f"  • {a.get('trainer','')}: {a.get('notiz','').strip()}"
                for a in anmerkungen_server if a.get("notiz","").strip()
            ) + "\n"
        send_email(
            f"Trainingsplan {datum} automatisch erstellt",
            f"Hallo Noah,\n\n"
            f"der Trainingsplan fuer {wtag}, {datum} wurde automatisch erstellt und hochgeladen.\n\n"
            f"Geraete: {geraet_1} + {geraet_2}\n\n"
            f"Abwesend:\n"
            f"  Trainer: {', '.join(absences.get('Trainer','')) or 'Alle da'}\n"
            f"  G1: {', '.join(absences.get('G1','')) or 'Alle da'}\n"
            f"  G2: {', '.join(absences.get('G2','')) or 'Alle da'}\n"
            f"  G3: {', '.join(absences.get('G3','')) or 'Alle da'}\n"
            f"  G4: {', '.join(absences.get('G4','')) or 'Alle da'}\n"
            f"{anm_text}\n"
            f"Der Plan ist jetzt auf der Website.\n"
            f"Falls etwas nicht stimmt, oeffne Claude.\n\nGrueße, Auto-Bot"
        )
        print(f"FERTIG! Plan fuer {datum} hochgeladen.")

    sftp.close()
    ssh.close()

if __name__ == "__main__":
    main()
