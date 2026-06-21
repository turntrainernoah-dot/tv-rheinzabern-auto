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

import json, hashlib, os, re, sys, tempfile, shutil, subprocess, urllib.request, urllib.parse
from datetime import date, timedelta, datetime, timezone

import paramiko
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

# ════════════════════════════════════════════════════════════════
#  STATISCHE KONFIGURATION
# ════════════════════════════════════════════════════════════════

ALLE_TURNER = {
    "G1": ["Felix E.", "Finn M.", "Sinan Y.", "Ilyas E.", "Jonathan S.", "Hannes G.", "Ben B."],
    "G2": ["Henry K.", "Matti G.", "Levent K.", "Caius C."],
    "G3": ["Erik E.", "Artem T.", "Finn T.", "Ben F.", "Michael K."],
    "G4": ["Felix L.", "Anton K.", "Mika W.", "Jamie G."],
}
ALLE_TRAINER = ["Noah W.", "Andy K.", "Fabian G.", "Cassian P.", "Julian K.", "Torben W."]

# Mapping: Website-Format (nach normalize) → Anzeigename
WEBSITE_TO_DISPLAY = {
    # G1
    "Felix G1": "Felix E.", "Finn G1": "Finn M.", "Sinan": "Sinan Y.",
    "Ilyas": "Ilyas E.", "Jonathan": "Jonathan S.", "Hannes": "Hannes G.",
    "Ben G1": "Ben B.",
    # G2
    "Henry": "Henry K.", "Matti": "Matti G.", "Levent": "Levent K.", "Caius": "Caius C.",
    # G3
    "Erik": "Erik E.", "Artem": "Artem T.", "Finn G3": "Finn T.",
    "Ben G3": "Ben F.", "Michael": "Michael K.",
    # G4
    "Felix G4": "Felix L.", "Anton": "Anton K.", "Mika": "Mika W.", "Jamie": "Jamie G.",
    # Trainer
    "Noah": "Noah W.", "Andy": "Andy K.", "Fabian": "Fabian G.",
    "Cassian": "Cassian P.", "Julian": "Julian K.", "Torben": "Torben W.",
}

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
WHATSAPP_PHONE    = os.environ.get("WHATSAPP_PHONE", "")
CALLMEBOT_APIKEY  = os.environ.get("CALLMEBOT_APIKEY", "")

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

def get_gear_from_plan_data(state):
    """
    Berechnet den korrekten Geräte-Combo-Index aus plan_data (nicht aus geraet_combo_index).
    Sucht den neuesten Eintrag in plan_data mit Gerätedaten und leitet den nächsten Index ab.
    Gibt zurück: (letzter_combo_idx, g1_starts_geraet2, letztes_datum) oder (None, None, None).
    """
    plan_data = state.get("plan_data", {})
    if not plan_data:
        return None, None, None

    dated_plans = []
    for datum_kurz, pdata in plan_data.items():
        if pdata and "geraet_1" in pdata and "geraet_2" in pdata:
            try:
                d = datetime.strptime(datum_kurz, "%d.%m.%y").date()
                dated_plans.append((d, datum_kurz, pdata))
            except Exception:
                pass

    if not dated_plans:
        return None, None, None

    dated_plans.sort(key=lambda x: x[0])
    latest_date, latest_key, latest = dated_plans[-1]

    g1 = latest.get("geraet_1")
    g2 = latest.get("geraet_2")
    for idx, (r1, r2) in enumerate(GERAETE_ROTATION):
        if r1 == g1 and r2 == g2:
            print(f"[ROTATION] Letzter Plan: {latest_key} = {g1}+{g2} (Combo {idx}) → nächster: {(idx+1)%3}")
            return idx, latest.get("g1_starts_geraet2", False), latest_date

    print(f"[ROTATION-WARN] Kombo {g1}+{g2} nicht in GERAETE_ROTATION gefunden – Fallback auf State.")
    return None, None, None

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
    """Website-Format → interner Anzeigename (z.B. 'Felix (G1)' → 'Felix E.')."""
    # Schritt 1: "(G1)" Suffix entfernen → "Felix G1"
    intermediate = re.sub(r'\s*\(G(\d+)\)$', r' G\1', name.strip())
    # Schritt 2: Mapping auf neues Anzeigeformat
    return WEBSITE_TO_DISPLAY.get(intermediate, WEBSITE_TO_DISPLAY.get(name.strip(), intermediate))

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

def read_fixed_entries(sftp):
    """Liest fixed_entries.json vom Server (falls vorhanden).
    Struktur: { "DD.MM.JJ": { "lock_trainer_plan": bool,
                               "fixed_absences": { "Gruppe": ["Name", ...] },
                               "fixed_trainer_plan": { "Trainer": [[text,farbe],...] } } }
    """
    try:
        f    = sftp.open("fixed_entries.json", "r")
        data = json.loads(f.read().decode("utf-8", errors="replace"))
        f.close()
        print(f"[FIXED] fixed_entries.json geladen: {list(data.keys())}")
        return data
    except Exception:
        return {}

def apply_fixed_absences(absences, fixed_entry):
    """Merged fixed_absences in die berechneten Abwesenheiten (addiert, entfernt nichts)."""
    for gruppe, namen in fixed_entry.get("fixed_absences", {}).items():
        existing = absences.get(gruppe, [])
        for name in namen:
            if name not in existing:
                existing.append(name)
                print(f"[FIXED] {name} ({gruppe}) als fix-abwesend hinzugefügt.")
        absences[gruppe] = existing
    return absences

# ════════════════════════════════════════════════════════════════
#  TRAININGSENTFALL (Training abgesagt – von Trainer auf der Website markiert)
# ════════════════════════════════════════════════════════════════

def read_trainingsentfall(sftp):
    """Liest /abmeldungen/trainingsentfall.json (Liste von Y-m-d Strings).
    Single Source of Truth für abgesagte Trainings – wird auf der Website gepflegt."""
    try:
        f    = sftp.open("abmeldungen/trainingsentfall.json", "r")
        data = json.loads(f.read().decode("utf-8", errors="replace"))
        f.close()
        if isinstance(data, list):
            dates = [d for d in data if isinstance(d, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", d)]
            print(f"[ENTFALL] trainingsentfall.json geladen: {dates}")
            return dates
    except Exception:
        pass
    return []

def build_entfall_pdf(datum, datum_kurz, wochentag):
    """Erzeugt den minimalen Trainingsentfall-Hinweis (xlsx+pdf), identisch zum
    manuellen 12.06-Hinweis. Gibt (xlsx_path, pdf_path) zurück."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Trainingsplan"
    for col in "ABCDEFGHI":
        ws.column_dimensions[col].width = 13.0
    ws.merge_cells("B1:H1"); ws.merge_cells("B2:H2"); ws.merge_cells("B3:H3")
    c = ws["B1"]; c.value = f"TRAININGSPLAN | {wochentag}, {datum}"
    c.fill = fill("2C3E50"); c.font = font(bold=True, color="FFFFFF", size=14); c.alignment = align()
    ws.row_dimensions[1].height = 30
    c = ws["B2"]; c.value = "⚠  TRAININGSENTFALL"
    c.fill = fill("C0392B"); c.font = font(bold=True, color="FFFFFF", size=22); c.alignment = align()
    ws.row_dimensions[2].height = 50
    c = ws["B3"]; c.value = f"Das Training am {datum} ist ausgefallen."
    c.fill = fill("FAD7A0"); c.font = font(bold=False, color="000000", size=12); c.alignment = align()
    ws.row_dimensions[3].height = 28
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = "A1:I4"

    out_dir = "/tmp/trainingsplan"; os.makedirs(out_dir, exist_ok=True)
    xlsx_path = os.path.join(out_dir, f"{datum_kurz}_Trainingsplan.xlsx")
    pdf_path  = os.path.join(out_dir, f"{datum_kurz}_Trainingsplan.pdf")
    wb.save(xlsx_path)
    tmp_dir = tempfile.mkdtemp()
    result  = subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmp_dir, xlsx_path],
        capture_output=True, text=True, timeout=90)
    tmp_pdf = os.path.join(tmp_dir, os.path.basename(xlsx_path).replace(".xlsx", ".pdf"))
    if result.returncode == 0 and os.path.exists(tmp_pdf):
        shutil.copy2(tmp_pdf, pdf_path)
    else:
        raise RuntimeError(f"LibreOffice Fehler (Entfall): {result.stderr[:300]}")
    return xlsx_path, pdf_path

def publish_entfall(sftp, datum, datum_kurz, wochentag):
    """Veröffentlicht den Entfall-Hinweis: ersetzt einen evtl. vorhandenen normalen
    Plan durch den Trainingsentfall-Hinweis und aktualisiert das Widget-JSON."""
    xlsx_path, pdf_path = build_entfall_pdf(datum, datum_kurz, wochentag)
    upload_pdf(sftp, pdf_path, datum_kurz)
    upload_xlsx(sftp, xlsx_path, datum_kurz)
    tag, monat, jahr = datum.split(".")
    datum_iso = f"{jahr}-{monat}-{tag}"
    aktuell = {
        "datum": datum, "datum_iso": datum_iso, "wochentag": wochentag,
        "trainingsentfall": True,
        "pdf_url": f"https://tv-rheinzabern.e-websolutions.de/trainingspläne/{datum_kurz}_Trainingsplan.pdf",
        "einteilung": {},
    }
    upload_aktuell_json(sftp, aktuell)
    print(f"[ENTFALL] Entfall-Hinweis für {datum} veröffentlicht.")

def remove_plan_files(sftp, datum_kurz):
    """Entfernt PDF/XLSX eines Plans vom Server (z.B. wenn Entfall aufgehoben wird,
    damit ein frischer normaler Plan erzeugt wird)."""
    for ext in ("pdf", "xlsx"):
        try:
            sftp.remove(f"trainingspläne/{datum_kurz}_Trainingsplan.{ext}")
            print(f"[ENTFALL] Entfernt: trainingspläne/{datum_kurz}_Trainingsplan.{ext}")
        except Exception:
            pass

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

def build_aktuell_json(datum, datum_kurz, wochentag, geraet_1, geraet_2, trainer_plan, abwesend):
    """Erstellt trainingsplan_aktuell.json für iOS Widgets."""
    tag, monat, jahr = datum.split(".")
    datum_iso = f"{jahr}-{monat}-{tag}"
    pdf_url = f"https://tv-rheinzabern.e-websolutions.de/trainingspläne/{datum_kurz}_Trainingsplan.pdf"
    einteilung = {}
    for trainer, plan in trainer_plan.items():
        if plan is None:
            continue
        if trainer in abwesend.get("Trainer", []):
            continue
        slots = []
        for slot_idx, slot_time in enumerate(ZEITSLOTS):
            if slot_idx < len(plan):
                text, _ = plan[slot_idx]
                slots.append({"zeit_start": slot_time.split("–")[0], "aufgabe": text})
        einteilung[trainer] = slots
    return {
        "datum": datum, "datum_iso": datum_iso, "wochentag": wochentag,
        "geraet1": geraet_1, "geraet2": geraet_2, "pdf_url": pdf_url,
        "einteilung": einteilung,
    }

def upload_aktuell_json(sftp, json_data):
    """Lädt trainingsplan_aktuell.json auf Server hoch (für iOS Widgets)."""
    data = json.dumps(json_data, indent=2, ensure_ascii=False).encode("utf-8")
    f = sftp.open("trainingspläne/trainingsplan_aktuell.json", "wb")
    f.write(data)
    f.close()
    print("[OK] Hochgeladen: trainingspläne/trainingsplan_aktuell.json")

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
    # Es wird IMMER ein Plan erstellt. Staffing-Probleme sind nur soft_warnings
    # (Hinweis per WhatsApp/Mail) -> autonomer Plan auch bei wenig Trainern/Kindern.
    issues        = []
    soft_warnings = []
    anwesend_trainer = [t for t in ALLE_TRAINER if t not in absences.get("Trainer", [])]
    n = len(anwesend_trainer)
    if n <= 3:
        soft_warnings.append((f"low_trainer_{n}",
            f"Nur {n} Trainer anwesend ({', '.join(anwesend_trainer) or '-'}). "
            f"Gruppen automatisch zusammengelegt, alle trainieren 17:00-19:00."))
    for gruppe, names in absences.items():
        for name in names:
            if name not in ALLE_BEKANNTEN_NAMES:
                soft_warnings.append((f"unknown_{gruppe}_{name}",
                    f"Unbekannter Name: '{name}' (Gruppe {gruppe})"))
    for gruppe, turner in ALLE_TURNER.items():
        abw = absences.get(gruppe, [])
        anwesend = len(turner) - len(abw)
        if len(abw) >= len(turner):
            soft_warnings.append((f"empty_{gruppe}", f"Gruppe {gruppe}: alle Turner abwesend."))
        elif anwesend <= 2:
            soft_warnings.append((f"low_turner_{gruppe}_{anwesend}",
                f"Gruppe {gruppe} hat nur {anwesend} Turner anwesend."))
    return issues, anwesend_trainer, soft_warnings

# ════════════════════════════════════════════════════════════════
#  TRAINER-EINTEILUNG
# ════════════════════════════════════════════════════════════════

def _build_lowstaff_plan(available, abwesend):
    """Notbesetzung (<=3 Trainer): Gruppen zusammenlegen, alle trainieren 17:00-19:00.
    3 Trainer -> G1, G2, G3+G4 | 2 -> G1+G2, G3+G4 | 1 -> alle zusammen."""
    pool = list(available)
    def pop_by_first(lst, fn):
        for i, t in enumerate(lst):
            if t.startswith(fn):
                return lst.pop(i)
        return None
    ordered = []
    noah = pop_by_first(pool, "Noah")
    if noah:
        ordered.append(noah)
    ordered += pool
    n = len(ordered)
    tpl = {
        "G1":    [("AW G1", "aufwaermen"), ("G1", "g1_blau"), ("G1", "g1_blau"), ("G1", "g1_blau"), ("Abbauen", "aufbauen")],
        "G2":    [("AW G2", "aufwaermen"), ("G2", "g1_gruen"), ("G2", "g1_gruen"), ("G2", "g1_gruen"), ("Abbauen", "aufbauen")],
        "G1+G2": [("AW G1+G2", "aufwaermen"), ("G1+G2", "g1_blau"), ("G1+G2", "g1_blau"), ("G1+G2", "g1_blau"), ("Abbauen", "aufbauen")],
        "G3+G4": [("AW G3+G4", "aufwaermen"), ("G3+G4", "g2_orange"), ("G3+G4", "g2_orange"), ("G3+G4", "g2_orange"), ("Abbauen", "aufbauen")],
        "Alle":  [("AW Alle", "aufwaermen"), ("Alle Gruppen", "g1_gruen"), ("Alle Gruppen", "g1_gruen"), ("Alle Gruppen", "g1_gruen"), ("Abbauen", "aufbauen")],
    }
    if n == 3:
        labels = ["G1", "G2", "G3+G4"]
    elif n == 2:
        labels = ["G1+G2", "G3+G4"]
    elif n == 1:
        labels = ["Alle"]
    else:
        labels = []
    assign = {}
    for i, t in enumerate(ordered):
        if i < len(labels):
            assign[t] = labels[i]
    TRAINER_PLAN = {}
    for t in ALLE_TRAINER:
        TRAINER_PLAN[t] = [tuple(x) for x in tpl[assign[t]]] if t in assign else None
    anmerkungen = ["Alle Gruppen trainieren von 17:00-19:00 Uhr."]
    if n == 0:
        anmerkungen.append("ACHTUNG: Kein Trainer anwesend - bitte dringend klaeren!")
    return TRAINER_PLAN, {}, anmerkungen



def build_trainer_plan(absences, geraet_1, geraet_2, g1_starts_geraet2):
    abwesend  = absences.get("Trainer", [])
    available = [t for t in ALLE_TRAINER if t not in abwesend]
    n = len(available)

    if n <= 3:
        return _build_lowstaff_plan(available, abwesend)

    merge_g23 = (n == 3)

    assignment = {}
    pool = list(available)

    # Suche nach Vorname (robust gegen Namensformat-Änderungen)
    def pop_by_first(lst, first_name):
        for i, t in enumerate(lst):
            if t.startswith(first_name):
                return lst.pop(i)
        return None

    noah = pop_by_first(pool, "Noah")
    if noah:
        assignment[noah] = "G1"
    else:
        assignment[pool.pop(0)] = "G1"

    andy = pop_by_first(pool, "Andy")
    if andy:
        g4_trainer = andy
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

    andy_full = next((t for t in available if t.startswith("Andy")), None)
    if andy_full and assignment.get(andy_full) == "Springer":
        if TRAINER_PLAN.get(andy_full):
            TRAINER_PLAN[andy_full][4] = ("G4", G4_SLOT4)

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
                         f"{'✗' if ab else '✓'} {name}",
                         FARBEN["abwesend_turner"] if ab else FARBEN["anwesend"],
                         font(size=9, color="FFFFFF"), align(h="left"))
            else:
                ws.cell(row=row, column=ci).fill = fill(a)

        if i < len(ALLE_TRAINER):
            name = ALLE_TRAINER[i]
            ab   = name in abwesend.get("Trainer", [])
            set_cell(ws, row, 7,
                     f"{'✗' if ab else '✓'} {name}",
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
#  MAIN
# ════════════════════════════════════════════════════════════════

def apply_config_roster(sftp):
    """Laedt config/config.json (von admin.php gepflegt) und ueberschreibt
    ALLE_TURNER, ALLE_TRAINER, WEBSITE_TO_DISPLAY. Bei jedem Fehler bleibt die
    Hardcodierung aktiv -> der Plan wird NIE durch eine fehlerhafte config kaputt."""
    global ALLE_TURNER, ALLE_TRAINER, WEBSITE_TO_DISPLAY
    try:
        f = sftp.open("config/config.json", "r")
        cfg = json.loads(f.read().decode("utf-8", errors="replace"))
        f.close()
    except Exception as e:
        print(f"[CONFIG] config.json nicht ladbar ({e!r}) - nutze Hardcodierung.")
        return
    try:
        gruppen = cfg.get("gruppen") or ["G1", "G2", "G3", "G4"]
        turner  = {g: [] for g in gruppen}
        trainer = []
        w2d     = {}
        for p in cfg.get("personen", []):
            ni = p.get("name_intern") or p.get("anzeige")
            if not ni:
                continue
            anz = (p.get("anzeige") or ni).strip()
            key = re.sub(r'\s*\(G(\d+)\)$', r' G\1', anz)
            w2d[key] = ni
            if p.get("rolle") == "trainer":
                trainer.append(ni)
            elif p.get("rolle") == "turner":
                g = p.get("gruppe")
                turner.setdefault(g, []).append(ni)
        if not trainer or not any(turner.values()):
            print("[CONFIG] config.json unvollstaendig - nutze Hardcodierung.")
            return
        ALLE_TURNER        = turner
        ALLE_TRAINER       = trainer
        WEBSITE_TO_DISPLAY = w2d
        print(f"[CONFIG] Roster aus config.json: "
              f"{sum(len(v) for v in turner.values())} Turner, {len(trainer)} Trainer.")
    except Exception as e:
        print(f"[CONFIG] Fehler beim Aufbau ({e!r}) - nutze Hardcodierung.")


def build_admin_trainer_plan(absences, geraet_1, geraet_2, g1_starts_geraet2, partial):
    """Admin-Plan: manuell gesetzte Trainer-Zellen sind feste Vorgaben.
    Manuell belegte Trainer werden aus der Auto-Verteilung herausgenommen, damit
    ihre Gruppe (z.B. G4) automatisch von einem anderen Trainer uebernommen wird.
    Leere Randzeiten (erste/letzte Zeile) werden mit Aufbauen/Abbauen gefuellt."""
    partial = partial or {}
    committed = [t for t, s in partial.items()
                 if isinstance(s, list) and any((c and len(c) >= 2 and (c[0] or c[1])) for c in s)]
    nulled = [t for t, s in partial.items() if s is None]
    abs2 = {k: list(v) for k, v in absences.items()}
    abs2.setdefault("Trainer", [])
    for t in committed + nulled:
        if t not in abs2["Trainer"]:
            abs2["Trainer"].append(t)
    base, _s, _a = build_trainer_plan(abs2, geraet_1, geraet_2, g1_starts_geraet2)
    for t in committed:
        cells = partial[t]
        n = len(cells)
        out = []
        for i, c in enumerate(cells):
            if c and len(c) >= 2 and (c[0] or c[1]):
                out.append((c[0], c[1]))
            elif i == 0:
                out.append(("Aufbauen", "aufbauen"))
            elif i == n - 1:
                out.append(("Abbauen", "aufbauen"))
            else:
                out.append(("", ""))
        base[t] = out
    for t in nulled:
        base[t] = None
    return base



def main():
    print(f"=== TV Rheinzabern Auto-Trainingsplan | {date.today()} ===\n")

    print("Verbinde mit Server...")
    ssh, sftp = get_sftp()
    apply_config_roster(sftp)

    state              = load_state(sftp)
    abmeldungen, raw_abm_hash = read_abmeldungen(sftp)
    anmerkungen_server = read_anmerkungen_server(sftp)
    fixed_entries      = read_fixed_entries(sftp)
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

    # Fixed entries anwenden (erzwingt Abwesenheiten, die das Auto-System nicht ändern darf)
    fixed_for_date = fixed_entries.get(datum_kurz, {})
    lock_trainer   = fixed_for_date.get("lock_trainer_plan", False)
    fixed_tp_raw   = fixed_for_date.get("fixed_trainer_plan")  # vorberechneter Trainer-Plan
    if fixed_for_date:
        absences = apply_fixed_absences(absences, fixed_for_date)
        if lock_trainer:
            print(f"[FIXED] Trainer-Plan für {datum_kurz} ist gesperrt – Auto-Berechnung deaktiviert.")

    # ── Trainingsentfall-Check (Training wurde auf der Website abgesagt) ──────────
    # trainingsentfall.json ist die Single Source of Truth. Ist das nächste Training
    # als Entfall markiert, wird KEIN normaler Plan erzeugt, sondern der Entfall-Hinweis
    # veröffentlicht. Wird der Entfall wieder aufgehoben, erzeugt das System einen frischen Plan.
    entfall_list      = read_trainingsentfall(sftp)
    datum_iso         = training_date.strftime("%Y-%m-%d")
    entfall_published = state.setdefault("entfall_published", [])

    if datum_iso in entfall_list:
        if datum_kurz in entfall_published and plan_exists(sftp, datum_kurz):
            print(f"[ENTFALL] {datum} bereits als Entfall veröffentlicht – nichts zu tun.")
            sftp.close(); ssh.close()
            return
        publish_entfall(sftp, datum, datum_kurz, wtag)
        if datum_kurz not in entfall_published:
            entfall_published.append(datum_kurz)
        if datum_kurz not in state.setdefault("generated_plans", []):
            state["generated_plans"].append(datum_kurz)
        state.get("plan_data", {}).pop(datum_kurz, None)  # alten Plan-Hash verwerfen
        save_state(sftp, state)
        send_whatsapp(
            f"Hi Noah, Cloude hier ⚠️\n\n"
            f"Das Training am {wtag}, {datum} ist als Trainingsentfall markiert.\n"
            f"Ich habe den Entfall-Hinweis veröffentlicht und erstelle KEINEN normalen Plan."
        )
        print(f"[ENTFALL] Entfall für {datum} verarbeitet.")
        sftp.close(); ssh.close()
        return
    elif datum_kurz in entfall_published:
        # Entfall wurde wieder aufgehoben → Hinweis entfernen, frischen Plan erzeugen
        print(f"[ENTFALL] Entfall für {datum} aufgehoben → normaler Plan wird neu erstellt.")
        entfall_published.remove(datum_kurz)
        remove_plan_files(sftp, datum_kurz)
        state.get("plan_data", {}).pop(datum_kurz, None)
        if datum_kurz in state.get("generated_plans", []):
            state["generated_plans"].remove(datum_kurz)
        save_state(sftp, state)

    # -- Admin-Editor (manuell_bearbeitet aus admin.php) anwenden --------------
    force_regen = False
    admin_fixed_hash = ""
    if fixed_for_date.get("manuell_bearbeitet"):
        import hashlib as _hl
        admin_fixed_hash = _hl.md5(json.dumps(fixed_for_date, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()
        _g1 = fixed_for_date.get("geraet_1") or "Boden"
        _g2 = fixed_for_date.get("geraet_2") or "Barren"
        _g1s = fixed_for_date.get("g1_starts_geraet2", state.get("g1_starts_geraet2", False))
        _base_tp = build_admin_trainer_plan(absences, _g1, _g2, _g1s, fixed_for_date.get("fixed_trainer_partial") or {})
        fixed_for_date["fixed_trainer_plan"] = _base_tp
        fixed_for_date["lock_trainer_plan"] = True
        lock_trainer = True
        fixed_tp_raw = _base_tp
        _notiz = (fixed_for_date.get("notiz") or "").strip()
        if _notiz:
            anmerkungen_server.append({"trainer": "Hinweis", "notiz": _notiz})
        _stored_fh = state.get("plan_data", {}).get(datum_kurz, {}).get("fixed_hash", "")
        if admin_fixed_hash != _stored_fh:
            print(f"[ADMIN] Manuelle Planbearbeitung {datum_kurz} (neu/geaendert) -> regeneriere.")
            remove_plan_files(sftp, datum_kurz)
            state.get("plan_data", {}).pop(datum_kurz, None)
            if datum_kurz in state.get("generated_plans", []):
                state["generated_plans"].remove(datum_kurz)
            force_regen = True

    if plan_exists(sftp, datum_kurz):
        # ── Geschützter Plan → nie überschreiben ───────────────
        if datum_kurz in state.get("protected_plans", []):
            print(f"Plan {datum_kurz} ist als geschützt markiert – wird nicht überschrieben.")
            sftp.close(); ssh.close()
            return

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
        fixed_trainer_abs  = set(fixed_for_date.get("fixed_absences", {}).get("Trainer", []))

        # Trainer-Änderung ignorieren wenn der Plan gesperrt ist (lock_trainer_plan)
        trainer_changed = (stored_trainer_abs != new_trainer_abs)
        if trainer_changed and lock_trainer:
            print(f"[FIXED] Trainer-Änderung ignoriert (Plan gesperrt): "
                  f"{stored_trainer_abs} → {new_trainer_abs}")
            trainer_changed = False

        if trainer_changed:
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
            send_whatsapp(
                f"Hi Noah, hier ist Cloude ✋\n\n"
                f"Beim Trainingsplan für {wtag}, {datum} hat sich bei den Trainern etwas geändert "
                f"und ich kann den Plan nicht automatisch anpassen.\n\n"
                f"Neu abwesend: {', '.join(sorted(added)) or '–'}\n"
                f"Nicht mehr abwesend: {', '.join(sorted(removed)) or '–'}\n\n"
                f"Bitte erstell den Plan kurz manuell in Claude. 🙏"
            )
            print("Trainer-Aenderung! WhatsApp gesendet.")
            sftp.close(); ssh.close()
            return

        # UPDATE: gleiche Trainer, neue Abwesenheiten
        print("UPDATE: Trainer unveraendert, erstelle Plan neu...")

        geraet_1     = plan_data["geraet_1"]
        geraet_2     = plan_data["geraet_2"]
        g1_starts_g2 = plan_data["g1_starts_geraet2"]

        issues, anwesend_trainer, soft_warnings = detect_complex(absences, late_notes)
        # Bei gesperrtem Trainer-Plan: Trainer-Anzahl-Fehler ignorieren
        if lock_trainer:
            issues = [i for i in issues if "Trainer anwesend" not in i]
        if issues:
            body = (
                f"Hallo Noah,\n\nKonnte Plan fuer {datum} nicht aktualisieren:\n\n"
                + "\n".join(f"  - {i}" for i in issues)
                + f"\n\nBitte manuell pruefen.\n\nGrueße, Auto-Bot"
            )
            send_whatsapp(
                f"Hi Noah, Cloude hier 🚨\n\n"
                f"Beim automatischen Update des Trainingsplans für {datum} ist ein Fehler aufgetreten:\n\n"
                + "\n".join(f"• {i}" for i in issues) +
                f"\n\nBitte kurz manuell prüfen. 🙏"
            )
            print("Fehler beim Update! WhatsApp gesendet.")
            sftp.close(); ssh.close()
            return

        if lock_trainer and fixed_tp_raw:
            # Gesperrter Trainer-Plan: direkt aus fixed_entries (UPDATE path)
            trainer_plan = {
                k: [tuple(s) for s in v] if v is not None else None
                for k, v in fixed_tp_raw.items()
            }
            sondertiming = {}
            anmerkungen  = []
            if "Barren" in (geraet_1, geraet_2):
                anmerkungen.append("• Barren G3: Kippe üben")
            print("[FIXED] Verwende gesperrten Trainer-Plan aus fixed_entries (UPDATE).")
        else:
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
        aktuell_json = build_aktuell_json(datum, datum_kurz, wtag, geraet_1, geraet_2, trainer_plan, absences)
        upload_aktuell_json(sftp, aktuell_json)

        ids_gelesen = [a["id"] for a in anmerkungen_server if a.get("id")]
        mark_anmerkungen_gelesen(sftp, ids_gelesen)

        # Dedup: late_notes
        already_sent_notes    = set(plan_data.get("late_notes_sent", []))
        new_late_notes        = [n for n in late_notes if n not in already_sent_notes]

        # Dedup: soft warnings (≤2 Turner in Gruppe)
        already_sent_warnings = set(plan_data.get("warnings_sent", []))
        new_soft_warnings     = [(k, t) for k, t in soft_warnings if k not in already_sent_warnings]

        # State: Hash + dedup-Listen aktualisieren
        plan_data["absences_hash"]   = new_hash
        if admin_fixed_hash:
            plan_data["fixed_hash"] = admin_fixed_hash
        plan_data["stored_absences"] = absences
        if new_late_notes:
            plan_data["late_notes_sent"] = list(already_sent_notes | set(new_late_notes))
        if new_soft_warnings:
            plan_data["warnings_sent"] = list(already_sent_warnings | {k for k, _ in new_soft_warnings})
        save_state(sftp, state)

        # Notification: Trainer-Anmerkungen vorhanden → bitte prüfen (einmalig; Anmerkungen werden als gelesen markiert)
        if anmerkungen_server:
            anm_lines = "\n".join(
                f"  • {a.get('trainer','')}: {a.get('notiz','').strip()}"
                for a in anmerkungen_server if a.get("notiz","").strip()
            )
            send_whatsapp(
                f"Hi Noah, Cloude hier 📋\n\n"
                f"Neue Trainer-Anmerkung für {wtag}, {datum} wurde eingebaut:\n\n"
                f"{anm_lines}\n\n"
                f"Bitte kurz prüfen ob alles stimmt ✅"
            )

        # Notification: Verspätungen / frühes Gehen → nur NEUE Hinweise
        if new_late_notes:
            send_whatsapp(
                f"Hi Noah, Cloude hier ⏰\n\n"
                f"Neue Verspätungs-/Frühgeh-Hinweise für {wtag}, {datum}:\n\n"
                + "\n".join(f"{n}" for n in new_late_notes) +
                f"\n\nSind im Plan eingetragen."
            )

        # Notification: ≤2 Turner in Gruppe → nur neue Warnungen (einmalig pro Plan)
        if new_soft_warnings:
            warn_text = "\n".join(f"• {t}" for _, t in new_soft_warnings)
            send_whatsapp(
                f"Hi Noah, Cloude hier ⚠️\n\n"
                f"Hinweis für {wtag}, {datum}:\n\n"
                f"{warn_text}\n\n"
                f"Plan wurde trotzdem aktualisiert – nur zur Info."
            )

        # KEIN Update-Summary mehr (keine Routine-Benachrichtigung bei normalen Abwesenheits-Updates)
        print(f"UPDATE FERTIG fuer {datum}.")

    else:
        # ── Kein Plan vorhanden: erstelle neuen ────────────────
        # NUR im Publikationsfenster (Mi oder Fr nach 22:00 CEST / 20:00 UTC) erstellen
        now_utc   = datetime.now(timezone.utc)
        days_away = (training_date - date.today()).days

        if not is_publication_window() and not force_regen:
            print(
                f"Kein Plan vorhanden, aber außerhalb Publikationsfenster "
                f"({now_utc.strftime('%H:%M')} UTC, Wochentag {now_utc.weekday()}) → warte bis Mi/Fr 22:00 CEST."
            )
            sftp.close(); ssh.close()
            return

        if days_away > 5 and not force_regen:
            print(
                f"Kein Plan vorhanden, Training in {days_away} Tagen "
                f"({now_utc.strftime('%H:%M')} UTC) → noch zu weit entfernt."
            )
            sftp.close(); ssh.close()
            return

        print(f"Kein Plan vorhanden, starte Generierung... (Training in {days_away} Tagen, {now_utc.strftime('%H:%M')} UTC)")

        # Geräte-Rotation aus plan_data berechnen (robuster als geraet_combo_index)
        last_idx, last_g1_starts, last_date = get_gear_from_plan_data(state)
        if last_idx is not None:
            new_combo_idx = (last_idx + 1) % 3
            g1_starts_g2  = not last_g1_starts
            print(f"[ROTATION] Aus plan_data berechnet: letzter Combo {last_idx} → neu {new_combo_idx}")
        else:
            new_combo_idx = (state.get("geraet_combo_index", 2) + 1) % 3
            g1_starts_g2  = not state.get("g1_starts_geraet2", True)
            print(f"[ROTATION] Fallback auf geraet_combo_index: {new_combo_idx}")

        geraet_1, geraet_2 = GERAETE_ROTATION[new_combo_idx]

        # Bei gesperrtem Trainer-Plan: Geräte aus fixed_entries übernehmen (falls vorhanden)
        if lock_trainer and fixed_for_date.get("geraet_1"):
            geraet_1     = fixed_for_date["geraet_1"]
            geraet_2     = fixed_for_date["geraet_2"]
            g1_starts_g2 = fixed_for_date.get("g1_starts_geraet2", g1_starts_g2)
            print(f"[FIXED] Geräte aus fixed_entries: {geraet_1} + {geraet_2}, G1 starts G2: {g1_starts_g2}")

        issues, anwesend_trainer, soft_warnings = detect_complex(absences, late_notes)
        # Bei gesperrtem Trainer-Plan: Trainer-Anzahl-Fehler ignorieren (NEW path)
        if lock_trainer:
            issues = [i for i in issues if "Trainer anwesend" not in i]

        if issues:
            # Dedup: Nur einmal WA senden, nicht bei jedem 30-Min-Run
            plan_data_current = state.get("plan_data", {}).get(datum_kurz, {})
            if not plan_data_current.get("complex_warning_sent"):
                send_whatsapp(
                    f"Hi Noah, Cloude hier 🚨\n\n"
                    f"Den Trainingsplan für {wtag}, {datum} konnte ich leider nicht automatisch erstellen.\n\n"
                    f"Probleme:\n" +
                    "\n".join(f"• {i}" for i in issues) +
                    f"\n\nAnwesende Trainer: {', '.join(anwesend_trainer) or '–'}\n"
                    f"Abwesend: {', '.join(absences.get('Trainer', [])) or '–'}\n\n"
                    f"Bitte kurz manuell in Claude erstellen. 🙏"
                )
                print("Komplex! WhatsApp gesendet.")
                # Dedup-Flag setzen
                state.setdefault("plan_data", {})[datum_kurz] = plan_data_current
                state["plan_data"][datum_kurz]["complex_warning_sent"] = True
                save_state(sftp, state)
            else:
                print(f"Komplexe Situation für {datum_kurz} bereits gemeldet – kein erneutes WA.")
            sftp.close(); ssh.close()
            return

        if lock_trainer and fixed_tp_raw:
            # Gesperrter Trainer-Plan: direkt aus fixed_entries (NEW path)
            trainer_plan = {
                k: [tuple(s) for s in v] if v is not None else None
                for k, v in fixed_tp_raw.items()
            }
            sondertiming = {}
            anmerkungen  = []
            if "Barren" in (geraet_1, geraet_2):
                anmerkungen.append("• Barren G3: Kippe üben")
            print("[FIXED] Verwende gesperrten Trainer-Plan aus fixed_entries (NEW).")
        else:
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
        aktuell_json = build_aktuell_json(datum, datum_kurz, wtag, geraet_1, geraet_2, trainer_plan, absences)
        upload_aktuell_json(sftp, aktuell_json)

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
            "late_notes_sent":  list(late_notes),            # bereits gesendet beim Erstellen
            "warnings_sent":    [k for k, _ in soft_warnings],  # bereits gesendet beim Erstellen
            "fixed_hash":       admin_fixed_hash,
        }
        save_state(sftp, state)

        # Notification: Trainer-Anmerkungen vorhanden → bitte prüfen (einmalig)
        if anmerkungen_server:
            anm_lines = "\n".join(
                f"• {a.get('trainer','')}: {a.get('notiz','').strip()}"
                for a in anmerkungen_server if a.get("notiz","").strip()
            )
            send_whatsapp(
                f"Hi Noah, Cloude hier 📋\n\n"
                f"Neue Trainer-Anmerkung für {wtag}, {datum} wurde eingebaut:\n\n"
                f"{anm_lines}\n\n"
                f"Bitte kurz prüfen ob alles stimmt ✅"
            )

        # Notification: Verspätungen / frühes Gehen (einmalig beim Erstellen)
        if late_notes:
            send_whatsapp(
                f"Hi Noah, Cloude hier ⏰\n\n"
                f"Verspätungs-/Frühgeh-Hinweise für {wtag}, {datum}:\n\n"
                + "\n".join(f"{n}" for n in late_notes) +
                f"\n\nSind im Plan eingetragen."
            )

        # Notification: ≤2 Turner in Gruppe (einmalig beim Erstellen)
        if soft_warnings:
            warn_text = "\n".join(f"• {t}" for _, t in soft_warnings)
            send_whatsapp(
                f"Hi Noah, Cloude hier ⚠️\n\n"
                f"Hinweis für {wtag}, {datum}:\n\n"
                f"{warn_text}\n\n"
                f"Plan wurde trotzdem erstellt – nur zur Info."
            )

        # Haupt-Notification: Plan fertig (immer)
        anm_text = ""
        if anmerkungen_server:
            anm_text = f"\nAnmerkungen:\n" + "\n".join(
                f"• {a.get('trainer','')}: {a.get('notiz','').strip()}"
                for a in anmerkungen_server if a.get("notiz","").strip()
            ) + "\n"
        send_whatsapp(
            f"Hi Noah, Cloude hier 🎉\n\n"
            f"Trainingsplan für {wtag}, {datum} ist fertig und auf der Website!\n\n"
            f"Geräte: {geraet_1} + {geraet_2}\n"
            f"Abwesend:\n"
            f"  Trainer: {', '.join(absences.get('Trainer','')) or 'Alle da'}\n"
            f"  G1: {', '.join(absences.get('G1','')) or 'Alle da'}\n"
            f"  G2: {', '.join(absences.get('G2','')) or 'Alle da'}\n"
            f"  G3: {', '.join(absences.get('G3','')) or 'Alle da'}\n"
            f"  G4: {', '.join(absences.get('G4','')) or 'Alle da'}"
            f"{anm_text}"
            f"\ntv-rheinzabern.e-websolutions.de"
        )
        print(f"FERTIG! Plan fuer {datum} hochgeladen.")

    sftp.close()
    ssh.close()

if __name__ == "__main__":
    main()
# (Trainingsentfall-Support 19.06.2026)
