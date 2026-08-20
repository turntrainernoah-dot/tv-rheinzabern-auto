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
    "G4": ["Felix L.", "Anton K.", "Mika W."],
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
    "Felix G4": "Felix L.", "Anton": "Anton K.", "Mika": "Mika W.",
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

# ────────────────────────────────────────────────────────────────
#  STRUKTUR-TOGGLE: G3+G4 Dauer-Zusammenlegung (Noah, 18.08.2026)
# ────────────────────────────────────────────────────────────────
# G3 und G4 trainieren ab sofort IMMER zusammen als eine feste Einheit
# "G3+G4" (17:00-19:00 Uhr statt G4 vorher separat 17:30-19:30 Uhr), mit
# nur noch 1 Trainer statt vorher 2 (plus ganz normal Springer als Backup,
# wenn genug Trainer da sind - unveraendert wie bei jeder anderen Gruppe).
# Turner-Listen/Anwesenheit/Wochenchallenge bleiben pro G3 und G4 GETRENNT -
# nur Trainingszeit + Trainer-Einteilung werden zusammengelegt.
#
# Technisch nutzt dies exakt den Zusammenlegungs-Pfad (tpl_merged), der
# vorher schon situativ bei Trainer-Engpass/kleinen Gruppen lief - jetzt als
# feste Basis-Einheit statt als Ausnahme. Siehe build_trainer_plan() und
# build_ki_einteilung().
#
# REVERT ("mach alles wie zuvor, trenne G3+G4 auf"): einfach auf False
# setzen. Stellt die exakte Alt-Logik wieder her (G4 wieder eigenstaendig
# 17:30-19:30 mit eigenem Trainer; G3+G4-Zusammenlegung wieder nur situativ
# bei Engpass/kleinen Gruppen, wie vor dem 18.08.2026).
# Siehe Vault-Notiz [[G3+G4 Dauer-Zusammenlegung]] fuer Details.
G3_G4_PERMANENT_MERGE = True

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


def get_next_gear(state, exclude_date=None, fixed_entries=None):
    """Bestimmt die Geraete fuers naechste Training aus den letzten (bis zu) 8
    ECHTEN Trainingsplaenen in plan_data. Ausgefallene Trainings
    (entfall_published) und Eintraege ohne gueltige Geraetekombination werden
    ignoriert. Vom juengsten gueltigen Plan wird genau EIN Schritt weitergerueckt.
    Reihenfolge (Zyklus): Sprung+Reck -> Seitpferd+Ringe -> Boden+Barren.
    Gibt (combo_idx, g1_starts_geraet2) zurueck. geraet_combo_index wird bewusst
    NICHT mehr verwendet (vermeidet Drift / uebersprungene Geraete)."""
    plan_data = state.get("plan_data", {})
    entfall   = set(state.get("entfall_published", []))
    fixed_entries = fixed_entries or {}
    valid = []
    for datum_kurz, pdata in plan_data.items():
        if exclude_date and datum_kurz == exclude_date:
            continue
        if datum_kurz in entfall or not pdata:
            continue
        _fe = fixed_entries.get(datum_kurz) or {}
        g1 = (_fe.get("geraet_1") if isinstance(_fe, dict) else None) or pdata.get("geraet_1")
        g2 = (_fe.get("geraet_2") if isinstance(_fe, dict) else None) or pdata.get("geraet_2")
        idx = None
        for i, (r1, r2) in enumerate(GERAETE_ROTATION):
            if r1 == g1 and r2 == g2:
                idx = i; break
        if idx is None:
            continue
        try:
            d = datetime.strptime(datum_kurz, "%d.%m.%y").date()
        except Exception:
            continue
        valid.append((d, datum_kurz, idx, pdata.get("g1_starts_geraet2", False)))

    if not valid:
        print("[ROTATION] Kein gueltiger Geraete-Verlauf in plan_data -> Default Sprung+Reck.")
        return 1, True   # Sprung+Reck als sinnvoller Startpunkt

    valid.sort(key=lambda x: x[0])
    last8 = valid[-8:]
    _, last_key, last_idx, last_g1s = last8[-1]
    next_idx = (last_idx + 1) % len(GERAETE_ROTATION)
    verlauf = ", ".join(f"{k}={GERAETE_ROTATION[i][0]}+{GERAETE_ROTATION[i][1]}"
                        for _, k, i, _ in last8)
    print(f"[ROTATION] Letzte {len(last8)} echten Trainings: {verlauf}")
    print(f"[ROTATION] Juengster echter Plan {last_key} = "
          f"{GERAETE_ROTATION[last_idx][0]}+{GERAETE_ROTATION[last_idx][1]} "
          f"-> naechster: {GERAETE_ROTATION[next_idx][0]}+{GERAETE_ROTATION[next_idx][1]}")
    return next_idx, (not last_g1s)

def active_training_date():
    """Aktuell relevanter Trainingstag.
    - Trainingstag (Mi/Fr) TAGSUEBER: HEUTE (damit ein am Trainingstag markierter
      Entfall noch veroeffentlicht wird und Updates auf den heutigen Plan wirken).
    - Trainingstag ABENDS im Veroeffentlichungsfenster (>=22:00 CEST, Training
      vorbei): der NAECHSTE Trainingstag -> Mi-Abend erzeugt bereits den
      Freitagsplan, Fr-Abend den naechsten Mittwochsplan.
    - Sonst: der naechste Mi/Fr."""
    t = date.today()
    if t.weekday() in (2, 4) and not is_publication_window():
        return t
    return next_training_date()


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

def build_entfall_pdf(datum, datum_kurz, wochentag, notiz=""):
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
    notiz = (notiz or "").strip()
    b3_text = notiz if notiz else f"Das Training am {datum} ist ausgefallen."
    c = ws["B3"]; c.value = b3_text
    c.fill = fill("FAD7A0"); c.font = font(bold=False, color="000000", size=12); c.alignment = align(wrap=bool(notiz))
    # Zeilenhoehe an Anmerkungs-Laenge anpassen (mehrzeilige Anmerkungen)
    _lines = b3_text.count("\n") + 1 + (len(b3_text) // 55)
    ws.row_dimensions[3].height = max(28, _lines * 20)
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

def publish_entfall(sftp, datum, datum_kurz, wochentag, notiz=""):
    """Veröffentlicht den Entfall-Hinweis: ersetzt einen evtl. vorhandenen normalen
    Plan durch den Trainingsentfall-Hinweis und aktualisiert das Widget-JSON.
    notiz = im Admin-Bereich (fixed_entries) eingetragene Anmerkung -> wird Inhalt."""
    xlsx_path, pdf_path = build_entfall_pdf(datum, datum_kurz, wochentag, notiz)
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

SLOT_START_MIN = [17*60+0, 17*60+30, 18*60+0, 18*60+15, 19*60+0]   # 17:00,17:30,18:00,18:15,19:00
SLOT_END_MIN   = [17*60+30, 18*60+0, 18*60+15, 19*60+0, 19*60+30]

def _extract_time_min(notiz):
    """Findet eine Uhrzeit (HH:MM, HH.MM, 'HH Uhr') -> Minuten seit 0:00 oder None."""
    s = (notiz or "").lower()
    m = re.search(r'(\d{1,2})[:.](\d{2})', s)
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
        if 15 <= h <= 21 and 0 <= mi < 60:
            return h*60+mi
    m = re.search(r'(\d{1,2})\s*uhr', s)
    if m:
        h = int(m.group(1))
        if 15 <= h <= 21:
            return h*60
    return None

def parse_trainer_timing(notiz):
    """Richtung ('spaet'=kommt spaeter / 'frueh'=geht frueher / None) + Uhrzeit-Min."""
    s = (notiz or "").lower()
    t = _extract_time_min(s)
    frueh_kw = ["früher", "frueher", "geht", "gehe", "gehen", "weg", "raus", "verlass",
                "eher", "nur bis", "muss los", "muss weg", "muss gehen", "vorzeitig"]
    spaet_kw = ["später", "spaeter", "verspät", "verspaet", "komme", "kommt", "erst",
                "etwas spät", "bisschen spät", "verspätung", "verzögert"]
    is_frueh = any(k in s for k in frueh_kw)
    is_spaet = any(k in s for k in spaet_kw)
    if "bis" in s and t is not None:
        return "frueh", t
    if is_frueh and not is_spaet:
        return "frueh", t
    if is_spaet and not is_frueh:
        return "spaet", t
    if is_spaet and is_frueh:
        if "komme" in s or "kommt" in s or "erst" in s:
            return "spaet", t
        return "frueh", t
    return None, t

def is_timing_note(notiz, kind, tmin):
    """True = Verspaetung/frueher-gehen (Trainer bleibt da). Echte Abwesenheiten
    ('Urlaub','krank','Spätschicht') liefern kind=None -> False."""
    if kind is None:
        return False
    if tmin is not None:
        return True
    s = (notiz or "").lower()
    strong = ["später", "spaeter", "früher", "frueher", "verspät", "verspaet",
              "kommt erst", "komme erst", "geht früh", "geh früh", "nur bis"]
    return any(k in s for k in strong)

def compute_blocked_slots(kind, time_min):
    """Set geblockter Slot-Indizes (0..4). spaet: bis Ankunft; frueh: ab Weggang."""
    if kind == "spaet":
        if time_min is None:
            return {0}
        return {i for i in range(5) if time_min > SLOT_START_MIN[i]}
    if kind == "frueh":
        if time_min is None:
            return {4}
        return {i for i in range(5) if time_min < SLOT_END_MIN[i]}
    return set()

def upload_anmerkungen_auto(sftp, datum_kurz, lines):
    """Schreibt die generierten Auto-Anmerkungen je Datum nach anmerkungen_auto.json
    (Quelle fuer die Vorbefuellung des Notizfeldes in admin.php)."""
    try:
        try:
            f = sftp.open("anmerkungen_auto.json", "r")
            data = json.loads(f.read().decode("utf-8"))
            f.close()
            if not isinstance(data, dict):
                data = {}
        except Exception:
            data = {}
        data[datum_kurz] = list(lines)
        f = sftp.open("anmerkungen_auto.json", "wb")
        f.write(json.dumps(data, indent=2, ensure_ascii=False).encode("utf-8"))
        f.close()
        print(f"[ANM] anmerkungen_auto.json aktualisiert ({datum_kurz}: {len(lines)} Zeilen).")
    except Exception as e:
        print(f"[ANM] anmerkungen_auto.json nicht geschrieben: {e}")

def timing_annotations(trainer_timing):
    """Anmerkungs-Zeilen: '(Name) kommt später (HH:MM): Notiz'."""
    lines = []
    for name, info in (trainer_timing or {}).items():
        label = "kommt später" if info.get("kind") == "spaet" else "geht früher"
        ts    = info.get("time_str")
        notiz = (info.get("notiz") or "").strip()
        head  = f"{name} {label}" + (f" ({ts})" if ts else "")
        lines.append(f"• {head}: {notiz}" if notiz else f"• {head}")
    return lines

def _find_free_coverer(trainer_plan, trainer_timing, slot, exclude):
    """Freien Trainer fuer 'slot' finden: erst Springer, dann Auf-/Abbau.
    Wer in diesem Slot selbst spaet/frueh geblockt ist, scheidet aus."""
    def blocked_at(t):
        info = trainer_timing.get(t)
        return bool(info) and slot in info.get("blocked", [])
    for want in ("springer", "aufbauen"):
        for t, plan in trainer_plan.items():
            if t == exclude or not plan or slot >= len(plan):
                continue
            if blocked_at(t):
                continue
            _txt, ck = plan[slot]
            if ck == want:
                return t
    return None

def apply_timing_coverage(trainer_plan, trainer_timing):
    """Vertretung: geht ein Trainer frueher / kommt spaeter, uebernimmt fuer die
    fehlenden Slots ein freier Trainer (bevorzugt der Springer) dessen Gruppe.
    Der Springer verlaesst dafuer kurz seine Julian-Bereitschaft (laut Noah ok).
    MUSS VOR apply_timing_blocks laufen (liest die Original-Gruppenzelle)."""
    if not trainer_timing:
        return trainer_plan
    GRUPPEN_COLORS = {"aufwaermen", "g1_blau", "g1_gruen", "g2_orange", "g2_lila"}
    for name, info in trainer_timing.items():
        plan = trainer_plan.get(name)
        if not plan:
            continue
        for i in sorted(info.get("blocked", [])):
            if not (0 <= i < len(plan)):
                continue
            text, ck = plan[i]
            if ck not in GRUPPEN_COLORS:
                continue
            coverer = _find_free_coverer(trainer_plan, trainer_timing, i, exclude=name)
            if coverer:
                p = list(trainer_plan[coverer]); p[i] = (text, ck); trainer_plan[coverer] = p
    return trainer_plan


def apply_timing_blocks(trainer_plan, trainer_timing):
    """Ueberschreibt geblockte Slots verspaeteter/frueher gehender Trainer ROT mit
    'kommt später'/'geht früher'. Trainer bleibt eingeteilt (NICHT abwesend)."""
    if not trainer_timing:
        return trainer_plan
    for name, info in trainer_timing.items():
        plan = trainer_plan.get(name)
        if not plan:
            continue
        label = "kommt später" if info.get("kind") == "spaet" else "geht früher"
        ts    = info.get("time_str")
        cell  = f"{label}\n{ts}" if ts else label
        new   = list(plan)
        for i in info.get("blocked", []):
            if 0 <= i < len(new):
                new[i] = (cell, "sonder")
        trainer_plan[name] = new
    return trainer_plan

def get_absences(abmeldungen, training_date):
    target    = training_date.strftime("%Y-%m-%d")
    absences  = {"G1": [], "G2": [], "G3": [], "G4": [], "Trainer": []}
    late_notes = []
    trainer_timing = {}

    for entry in abmeldungen:
        if entry.get("datum") != target:
            continue
        name   = normalize_name(entry.get("name", "").strip())
        gruppe = entry.get("gruppe", "").strip()
        notiz  = entry.get("notiz", "").strip()

        # Trainer mit Verspaetung / frueher-gehen: ANWESEND, nur Zeitbloecke blocken.
        if name in ALLE_TRAINER:
            kind, tmin = parse_trainer_timing(notiz)
            is_versp = bool(entry.get("verspaetung", False))
            if is_versp or is_timing_note(notiz, kind, tmin):
                if kind is None:
                    kind = "spaet"
                blocked = compute_blocked_slots(kind, tmin)
                tstr = f"{tmin//60:02d}:{tmin%60:02d}" if tmin is not None else None
                trainer_timing[name] = {
                    "kind": kind, "time_min": tmin, "time_str": tstr,
                    "notiz": notiz, "blocked": sorted(blocked),
                }
                continue

        if gruppe in absences:
            absences[gruppe].append(name)

    return absences, late_notes, trainer_timing

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
            f"Gruppen werden automatisch sinnvoll zusammengelegt."))
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



# ════════════════════════════════════════════════════════════════
#  GERAETE-FARBEN: EINZIGE Quelle fuer beide Trainer-Plan-Builder
#  (Standard-Pfad build_trainer_plan + KI-Pfad build_ki_einteilung).
#  Bugfix 12.08.2026 (Noah): vorher hatte jeder Builder seine eigene Kopie
#  dieser Zuordnung, und im g1_starts_geraet2=True-Zweig kollidierten G2 und
#  G3 auf "g1_blau" (dieselbe Station gleichzeitig doppelt belegt), waehrend
#  G2 dadurch zweimal dasselbe Geraet turnte statt beide Geraete zu durchlaufen.
#  Jetzt zentral hier definiert, damit ein Fix nicht mehr an zwei Stellen
#  gepflegt werden muss und nicht wieder auseinanderlaufen kann.
# ════════════════════════════════════════════════════════════════

def geraet_farbe(gruppe, phase, g1_starts_geraet2):
    """Farbe/Station fuer G1/G2/G3 je Phase (1 = Slots 2+3, 2 = Slot 4).
    Regel (siehe Vault [[Gruppen, Zeiten und Geräte-Rotation]]): G1+G2 turnen
    IMMER denselben Geraetetyp gleichzeitig (auf den zwei verschiedenen Farben
    dieses Geraets), G3 IMMER den jeweils anderen Typ. Dadurch ist innerhalb
    einer Phase jede Farbe nur genau einer Gruppe zugeteilt, und jede Gruppe
    durchlaeuft ueber beide Phasen beide Geraete (nie zweimal dasselbe)."""
    if g1_starts_geraet2:
        m = ({"G1": "g2_orange", "G2": "g2_lila",  "G3": "g1_blau"}   if phase == 1
             else {"G1": "g1_blau",  "G2": "g1_gruen", "G3": "g2_orange"})
    else:
        m = ({"G1": "g1_blau",  "G2": "g1_gruen", "G3": "g2_orange"} if phase == 1
             else {"G1": "g2_orange", "G2": "g2_lila",  "G3": "g1_blau"})
    return m.get(gruppe, "aufbauen")

def geraet_farbe_g4(g1_starts_geraet2):
    """G4-Farben fuer Slot 4 (18:15-19:00) und Slot 5 (19:00-19:30) -- immer
    die Farbe, die von G1/G2/G3 in dem Moment gerade frei ist (G4 turnt laut
    Regel 'das gerade freie Geraet'). Muss wie geraet_farbe() vom
    g1_starts_geraet2-Flag abhaengen, sonst kollidiert G4 in Slot 4 mit G2,
    sobald G1+G2 in Phase 2 beide Faerbungen von Geraet 1 belegen."""
    return ("g2_lila", "g1_gruen") if g1_starts_geraet2 else ("g1_gruen", "g2_orange")

def _merged_ref_group(groups):
    """Referenz-Gruppe fuer die Farbwahl einer zusammengelegten Einheit (z.B.
    ["G3","G4"] fuer die G3+G4-Dauer-Einheit oder ein situativer Engpass-Merge).
    Bugfix 19.08.2026 (Noah, Regression nach G3+G4-Dauer-Zusammenlegung 18.08.2026):
    vorher hatten tpl_merged() (Standard-Pfad) und _cells() (KI-/Anmerkungs-Pfad)
    JEWEILS eine eigene, von g1_starts_geraet2 UNABHAENGIGE Hartverdrahtung
    ("g1_blau"/"g2_orange" nach Einheiten-Index-Paritaet). Bei g1_starts_geraet2=False
    ergab das fuer eine gerade Einheiten-Position exakt dieselben Farben wie G1 ->
    zwei Gruppen gleichzeitig auf "Geraet 1 blau" (z.B. 19.08.2026: G1 und G3+G4
    beide blau). Fix: die zusammengelegte Einheit uebernimmt die Farbreihe einer
    ihrer Original-Gruppen aus der bereits kollisionsfreien geraet_farbe()-Tabelle
    (G1/G2/G3 belegen dort garantiert unterschiedliche Farben je Phase) - dadurch
    kann eine Einheit nie mit einer separat gebliebenen Einzelgruppe kollidieren,
    da jede Gruppen-Identitaet immer nur in genau einer Einheit vorkommt."""
    for g in ("G1", "G2", "G3"):
        if g in groups:
            return g
    return "G3"   # Fallback (z.B. Einheit aus nur ["G4"] sollte hier nie ankommen)

# ────────────────────────────────────────────────────────────────
#  ROLLEN-ROTATION (Bug-Fix 19.08.2026, Noah)
# ────────────────────────────────────────────────────────────────
# Vorher: der Backup-Pool (Fabian/Cassian/Torben) wurde immer in fester
# ALLE_TRAINER-Reihenfolge iteriert - dadurch war Torben (letzter Eintrag)
# quasi Dauer-Springer, sobald Julian ausfiel und einer aus dem Pool eine
# Gruppe uebernehmen musste, kam immer nur Fabian dran. Jetzt: wer im
# LETZTEN Plan Springer war, bekommt beim naechsten Mal Prioritaet
# fuer eine Gruppe; wer eine Gruppe hatte, wird nach hinten sortiert.
# Speicherung: plan_data[datum_kurz]["trainer_roles"] = {Trainer: "Springer"|"Gruppe"|None}

def _extract_trainer_roles(trainer_plan):
    """Ermittelt aus einem generierten trainer_plan-Dict die Rolle jedes
    Trainers: 'Springer' | 'Gruppe' | None. 'Springer' = das Springer-Template
    (nur Aufbauen/Springer/Abbauen). Alles andere mit Zellen = 'Gruppe'."""
    roles = {}
    for t in ALLE_TRAINER:
        cells = (trainer_plan or {}).get(t)
        if not cells:
            roles[t] = None
            continue
        labs = {str(c[0] or "").strip().lower() for c in cells if isinstance(c, (list, tuple))}
        if labs and labs <= {"aufbauen", "springer", "abbauen"}:
            roles[t] = "Springer"
        else:
            roles[t] = "Gruppe"
    return roles

def _load_previous_trainer_roles(state, exclude_date=None):
    """Sucht den JUENGSTEN plan_data-Eintrag (ausser exclude_date) mit
    gespeicherten trainer_roles. Gibt {} zurueck, wenn nichts vorhanden ist
    (z.B. erste Woche nach dem Rollout dieser Rotations-Logik)."""
    pd = (state or {}).get("plan_data", {}) or {}
    entries = []
    for dk, d in pd.items():
        if exclude_date and dk == exclude_date:
            continue
        if not isinstance(d, dict):
            continue
        roles = d.get("trainer_roles")
        if not roles:
            continue
        try:
            entries.append((datetime.strptime(dk, "%d.%m.%y"), roles))
        except Exception:
            pass
    if not entries:
        return {}
    entries.sort(key=lambda x: x[0])
    return entries[-1][1] or {}


def build_trainer_plan(absences, geraet_1, geraet_2, g1_starts_geraet2, trainer_timing=None, prev_trainer_roles=None):
    abwesend  = absences.get("Trainer", [])
    available = [t for t in ALLE_TRAINER if t not in abwesend]
    n = len(available)
    trainer_timing = trainer_timing or {}
    prev_trainer_roles = prev_trainer_roles or {}

    # anwesende Kinder je Gruppe
    present = {g: max(0, len(ALLE_TURNER[g]) - len(absences.get(g, []))) for g in ("G1","G2","G3","G4")}
    active  = [g for g in ("G1","G2","G3","G4") if present[g] >= 1]

    anmerkungen = []

    if not active:
        anmerkungen.insert(0, "ACHTUNG: Alle Turner abwesend - kein Training.")
        return {t: None for t in ALLE_TRAINER}, {}, anmerkungen
    if n == 0:
        anmerkungen.insert(0, "ACHTUNG: Kein Trainer anwesend - bitte dringend klaeren!")
        return {t: None for t in ALLE_TRAINER}, {}, anmerkungen

    # Vollzeit-Trainer (halten eine Gruppe komplett) vs. partielle (frueh/spaet).
    # Ein frueh gehender / spaet kommender Trainer darf keine Gruppe ALLEIN halten
    # -> zaehlt nicht als Halter, wird Springer; die Vollzeit-Trainer uebernehmen.
    partial_set  = {t for t in available if t in trainer_timing}
    full_holders = [t for t in available if t not in partial_set]
    holders_pool = full_holders if full_holders else list(available)   # Notlage-Fallback
    n_hold = len(holders_pool)

    MIN_KIDS = 3   # jede Gruppe braucht >= 3 anwesende Turner, sonst zusammenlegen

    # ---- Einheiten bilden ----
    if G3_G4_PERMANENT_MERGE:
        # G3+G4 sind eine feste Basis-Einheit (siehe Toggle oben) statt zwei
        # separater Einzel-Einheiten. Alle nachfolgenden Schritte (Mindest-
        # groesse, Trainerzahl-Reduktion, Julian-Springer) arbeiten unveraendert
        # auf dieser Einheit weiter - identisch zum bisherigen tpl_merged-Pfad.
        units = []
        if present["G1"] >= 1: units.append(["G1"])
        if present["G2"] >= 1: units.append(["G2"])
        if present["G3"] >= 1 or present["G4"] >= 1: units.append(["G3", "G4"])
    else:
        units = [[g] for g in active]
    def ucount(u): return sum(present[g] for g in u)
    def crosses(a, b): return a[-1] == "G2" and b[0] == "G3"   # G2|G3-Grenze = last resort
    def _merge_at(i, j):
        lo, hi = sorted((i, j))
        units[lo] = units[lo] + units[hi]
        del units[hi]
    def best_merge():
        best = None
        for i in range(len(units)-1):
            combined = ucount(units[i]) + ucount(units[i+1])
            score = combined + (1000 if crosses(units[i], units[i+1]) else 0)
            if best is None or score < best[0]:
                best = (score, i)
        return best

    # 1) Mindestgroesse: zu kleine Gruppe (<3) mit bestem Nachbarn (gleiche Seite bevorzugt) mergen
    def _merge_undersized():
        for i, u in enumerate(units):
            if ucount(u) < MIN_KIDS and len(units) > 1:
                cands = []
                if i > 0: cands.append(i-1)
                if i < len(units)-1: cands.append(i+1)
                best = None
                for jj in cands:
                    a, b = (units[jj], units[i]) if jj < i else (units[i], units[jj])
                    sc = ucount(a) + ucount(b) + (1000 if crosses(a, b) else 0)
                    if best is None or sc < best[0]:
                        best = (sc, jj)
                _merge_at(i, best[1])
                return True
        return False
    while _merge_undersized():
        pass

    # 2) Nicht mehr Gruppen als Vollzeit-Trainer -> weiter zusammenlegen (kleinste Paarung)
    while len(units) > max(1, n_hold):
        bm = best_merge()
        if not bm: break
        _merge_at(bm[1], bm[1]+1)

    # 3) Julian-Springer-Prinzip (nur mit Vollzeit-Haltern, zusammengelegte Gruppe <= 8)
    julian_present = any(t.startswith("Julian") for t in holders_pool)
    if julian_present and (n_hold - len(units)) == 0 and len(units) >= 2:
        bm = best_merge()
        if bm is not None and (ucount(units[bm[1]]) + ucount(units[bm[1]+1])) <= 8:
            _merge_at(bm[1], bm[1]+1)

    # G4-Zeit-Hinweis, wenn G4 in einer Zusammenlegung ist (nur bei situativer
    # Zusammenlegung relevant - bei der permanenten G3+G4-Einheit ist das der
    # neue Normalfall und braucht keinen Extra-Hinweis mehr).
    if not G3_G4_PERMANENT_MERGE and any(len(u) >= 2 and "G4" in u for u in units):
        anmerkungen.append("Gruppe 4 hat zwischen 17:00-19:00 Training")

    # ---- Trainer den Einheiten zuordnen: nur Vollzeit-Halter bekommen Gruppen ----
    pool = list(holders_pool)
    def pop_first(name):
        for i, t in enumerate(pool):
            if t.startswith(name): return pool.pop(i)
        return None
    def unit_with(g):
        for idx, u in enumerate(units):
            if g in u: return idx
        return None
    noah = pop_first("Noah"); andy = pop_first("Andy"); julian = pop_first("Julian")

    # ── Rollen-Rotation (Bug-Fix 19.08.2026, Noah) ─────────────────────────
    # Sortiere den restlichen Pool (typischerweise Fabian/Cassian/Torben) so,
    # dass Trainer, die im LETZTEN Plan Springer waren, ZUERST kommen (=hoehere
    # Chance auf eine Gruppe). Trainer, die eine Gruppe hatten, wandern nach
    # hinten (=hoehere Chance, jetzt Springer zu werden). Wenn kein vorheriger
    # Plan mit Rollen bekannt ist, bleibt die urspruengliche Reihenfolge -
    # kein Regressionsrisiko fuer bestehende Zuteilungen.
    def _prev_role_key(t):
        r = prev_trainer_roles.get(t)
        if r == "Springer":
            return 0     # vorher Springer -> jetzt Prio auf Gruppe
        if r == "Gruppe":
            return 2     # vorher Gruppe -> jetzt nach hinten (eher Springer)
        return 1         # unbekannt/neu -> Mitte
    pool.sort(key=lambda t: (_prev_role_key(t), ALLE_TRAINER.index(t) if t in ALLE_TRAINER else 99))

    order = [t for t in (noah, andy, julian) if t] + pool
    pref = {}
    if noah: pref[noah] = unit_with("G1")
    if andy: pref[andy] = unit_with("G4")

    taken = set(); assigned = {}; springers = []
    for t in order:
        idx = pref.get(t)
        if idx is None or idx in taken:
            idx = next((j for j in range(len(units)) if j not in taken), None)
        if idx is None:
            springers.append(t)
        else:
            assigned[t] = idx; taken.add(idx)
    # partielle (frueh/spaet) + uebrige Vollzeit-Trainer -> Springer
    for t in available:
        if t not in assigned and t not in springers:
            springers.append(t)

    # ---- Farben / Templates ----
    def farbe(gruppe, phase):
        return geraet_farbe(gruppe, phase, g1_starts_geraet2)
    G4_SLOT3, G4_SLOT4 = geraet_farbe_g4(g1_starts_geraet2)

    def label_of(u):
        return "Alle Gruppen" if len(u) == 4 else "+".join(u)
    def tpl_single(g):
        return [("AW "+g,"aufwaermen"),(g,farbe(g,1)),(g,farbe(g,1)),(g,farbe(g,2)),("Abbauen","aufbauen")]
    def tpl_g4():
        return [("Aufbauen","aufbauen"),("AW G4","aufwaermen"),("AW G4","aufwaermen"),("G4",G4_SLOT3),("G4",G4_SLOT4)]
    def tpl_merged(label, groups):
        ref = _merged_ref_group(groups)
        c1, c2 = farbe(ref, 1), farbe(ref, 2)
        return [("AW "+label,"aufwaermen"),(label,c1),(label,c1),(label,c2),("Abbauen","aufbauen")]
    def tpl_springer():
        return [("Aufbauen","aufbauen"),("Springer","springer"),("Springer","springer"),("Springer","springer"),("Abbauen","aufbauen")]

    TRAINER_PLAN = {}
    for t in ALLE_TRAINER:
        if t not in assigned and t not in springers:
            TRAINER_PLAN[t] = None; continue
        if t in springers:
            TRAINER_PLAN[t] = tpl_springer(); continue
        u = units[assigned[t]]
        if u == ["G4"]:
            TRAINER_PLAN[t] = tpl_g4()
        elif len(u) == 1:
            TRAINER_PLAN[t] = tpl_single(u[0])
        else:
            TRAINER_PLAN[t] = tpl_merged(label_of(u), u)

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


def build_admin_trainer_plan(absences, geraet_1, geraet_2, g1_starts_geraet2, partial, ki=None):
    """Admin-Plan: manuell im Admin-Bereich gesetzte Trainer-Zellen (fixed_trainer_partial)
    sind HARTE Vorgaben. Manuell belegte Trainer werden aus der Auto-Verteilung
    herausgenommen. Der REST wird um diese Vorgabe herum geplant - und zwar inklusive
    KI-/Kommando-Anweisungen (ki.assign/merges aus Anmerkungen), falls vorhanden, statt
    sie stillschweigend zu verwerfen.
    Bugfix 19.08.2026 (Noah: "im Admin-Bereich festgelegte Bereiche werden ignoriert,
    die KI soll aussenrum planen"): vorher wurde bei vorhandenem fixed_trainer_partial
    IMMER nur der reine Auto-Builder (build_trainer_plan, ohne KI) fuer den Rest benutzt
    - jede KI-Zuteilung/Zusammenlegung aus einer gleichzeitigen Anmerkung ging verloren.
    Jetzt: von Admin fest belegte Gruppen (aus dem Zellentext der committed Trainer
    erkannt) werden fuer den Rest-Builder als 'voll abwesend' markiert, damit sie nicht
    doppelt vergeben werden - und der Rest nutzt build_ki_einteilung, wenn ki.assign/
    merges vorhanden sind. Leere Randzeiten (erste/letzte Zeile) werden mit
    Aufbauen/Abbauen gefuellt."""
    partial = partial or {}
    ki = ki or {}
    committed = [t for t, s in partial.items()
                 if isinstance(s, list) and any((c and len(c) >= 2 and (c[0] or c[1])) for c in s)]
    nulled = [t for t, s in partial.items() if s is None]

    # Von den fest zugewiesenen Trainern belegte Gruppen ermitteln (aus dem
    # Zellentext, z.B. "G3+G4" oder "G1"), damit der Rest-Builder sie nicht
    # nochmal vergibt (sonst doppelte Gruppenbelegung moeglich).
    covered_groups = set()
    for t in committed:
        labels = {str(c[0]).strip() for c in partial[t]
                  if isinstance(c, (list, tuple)) and len(c) >= 2 and c[0]}
        for lab in labels:
            norm = _ki_norm_merge(lab) or (lab if lab in ("G1", "G2", "G3", "G4") else None)
            if norm:
                covered_groups.update(norm.split("+"))

    abs2 = {k: list(v) for k, v in absences.items()}
    abs2.setdefault("Trainer", [])
    for t in committed + nulled:
        if t not in abs2["Trainer"]:
            abs2["Trainer"].append(t)
    for g in covered_groups & {"G1", "G2", "G3", "G4"}:
        existing = abs2.get(g, [])
        for name in ALLE_TURNER.get(g, []):
            if name not in existing:
                existing.append(name)
        abs2[g] = existing

    if ki.get("assign") or ki.get("merges"):
        try:
            base, _s, _a = build_ki_einteilung(abs2, ki, geraet_1, geraet_2, g1_starts_geraet2)
        except Exception as _e:
            print(f"[ADMIN] KI-Einteilung um Admin-Fix herum fehlgeschlagen ({_e!r}), Fallback Standard-Builder.")
            base, _s, _a = build_trainer_plan(abs2, geraet_1, geraet_2, g1_starts_geraet2)
    else:
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



# ════════════════════════════════════════════════════════════════
#  MINI-KI: freie Trainer-Anmerkungen -> Planaenderungen (GitHub Models)
# ════════════════════════════════════════════════════════════════
GH_MODELS_ENDPOINT = os.environ.get("GH_MODELS_ENDPOINT", "https://models.github.ai/inference/chat/completions")
GH_MODELS_MODEL    = os.environ.get("GH_MODELS_MODEL", "openai/gpt-4o")

def _ki_token():
    return os.environ.get("GH_MODELS_TOKEN") or os.environ.get("GITHUB_TOKEN") or ""

def _ki_call(messages, token, timeout=60):
    body = json.dumps({"model": GH_MODELS_MODEL, "messages": messages,
                       "temperature": 0, "response_format": {"type": "json_object"}}).encode("utf-8")
    req = urllib.request.Request(GH_MODELS_ENDPOINT, data=body, method="POST")
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Content-Type", "application/json")
    req.add_header("Accept", "application/json")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))["choices"][0]["message"]["content"]

_KI_PROMPT = """Du wandelst eine freie Trainer-Anmerkung fuer einen Kinder-Turntrainingsplan in striktes JSON um.

Kontext:
- Gruppen: <GR>
- Trainer: <TR>
- Geraete: <GE>
- Heute/naechstes Training: <NEXT>
- Diese Anmerkung stammt vom Trainer: "<WRITER>". "ich"/"mich"/"mir"/"meine" bezieht sich immer auf diesen Trainer.

Aktuelle Anmerkungen je Datum — NUR Referenz fuer Text-Aenderungen, leite daraus NIEMALS Aktionen (gruppe_entfall/timing/zuteilung/geraete/merges) ab: <CURRENT>

Gib GENAU dieses JSON zurueck (keine weiteren Felder, kein Text drumherum):
{"ziel": {"typ": "naechstes | datum | naechstes_geraet", "datum": "TT.MM.JJJJ oder null", "geraet": "Geraetename oder null"}, "notiz": "reiner Hinweistext oder null", "geraete": {"g1": "Geraet", "g2": "Geraet"} , "gruppe_entfall": ["G3"], "abwesend_kinder": [], "wieder_da_kinder": [], "abwesend_trainer": [], "einteilung": {}, "timing": [{"trainer": "Name", "richtung": "spaet | frueh", "uhrzeit": "HH:MM oder null"}], "merges": [["G3","G4"]], "zuteilung": [{"trainer": "Name", "gruppe": "z.B. G1, G2+G3 oder Springer"}], "notiz_neu": null, "reset": false, "unsicher": true}

Regeln:
- ziel.typ: KEIN Datum/Zeitpunkt genannt -> "naechstes" (gilt NUR fuer genau das naechste Training, NIE Dauerauftrag). Konkretes Datum (z.B. "am 14.10") -> "datum" mit datum. "naechstes Mal <Geraet>" / "wenn wir <Geraet> turnen" -> "naechstes_geraet" mit geraet. Wenn ziel unklar -> typ "naechstes".
- "kommt spaeter"/"erst um X" -> timing richtung "spaet"; "geht frueher"/"muss um X weg/los" -> "frueh". uhrzeit als HH:MM oder null.
- "GX entfaellt"/"faellt aus" -> gruppe_entfall ["GX"].
- Geraetewahl "wir turnen X und Y" -> geraete {"g1":X,"g2":Y}, sonst geraete null.
- Gruppen zusammenlegen OHNE genannten Trainer (z.B. "G3 und G4 zusammenlegen") -> merges [["G3","G4"]]. MIT genanntem Trainer ("Fabian macht G1+G2") -> zuteilung {trainer, gruppe:"G1+G2"}. "ich mache ..." -> trainer = der Schreiber. gruppe darf "Springer" oder "GX+GY" sein.
- Reiner Hinweis ("schreibe/bitte ...") -> nur notiz.
- Nur Namen aus der Trainerliste, nur Gruppen aus der Gruppenliste, nur Geraete aus der Geraeteliste.
- Keine Aktion in einem Feld: leere Liste [] bzw. null. Bei Unklarheit unsicher=true und notiz=Originaltext.
- Ist ein Trainer NICHT namentlich genannt (z.B. "der andere Trainer", "jemand", "ein Trainer macht G2+G3") -> als merges eintragen (ohne Trainer), NICHT als zuteilung.
- "loesche/entferne alle (bisherigen) Anmerkungen", "setze zurueck", "reset", "mach alles rueckgaengig" -> reset: true (alle bisherigen Vorgaben fuer das Ziel-Datum werden entfernt).
- notiz_neu: Setze es bei JEDER Aenderung des Anmerkungs-TEXTES — hinzufuegen ("fuege hinzu ..."), loeschen ("loesche die Zeile ..."), umformulieren oder ersetzen ("aendere X zu Y"). Gib dann IMMER den KOMPLETTEN neuen Anmerkungstext des Ziel-Datums zurueck: bestehende Zeilen aus der Referenz uebernehmen und die Aenderung anwenden. Bei reinen Struktur-/Plan-Aktionen ohne Text-Aenderung notiz_neu = null.
- WICHTIG: gruppe_entfall/timing/zuteilung/merges/geraete/reset NUR aus der AKTUELLEN Trainer-Anmerkung ableiten, NIEMALS aus der Referenz der aktuellen Anmerkungen.
- abwesend_kinder / wieder_da_kinder: einzelne Kinder ab-/wieder anmelden (nur Namen aus der Kinderliste). abwesend_trainer: Trainer die fehlen.
- einteilung: KOMPLETTE Zuordnung {Trainer: Label}, wenn die Anweisung Trainer den Gruppen zuordnet (z.B. "Andy G1, Noah G2, Cassian G4, Rest Springer"). Label = G1..G4, "GX+GY", "Springer" oder eigener Kurztext. Sonst {} = automatische Einteilung.
- Antworte ausschliesslich mit dem JSON-Objekt."""

def _ki_system_prompt(writer, next_date_str, current_notes=None):
    gr = "; ".join(f"{g}: {', '.join(ALLE_TURNER[g])}" for g in ALLE_TURNER)
    tr = ", ".join(ALLE_TRAINER)
    ge = ", ".join(sorted({g for c in GERAETE_ROTATION for g in c}))
    cur = "(keine)"
    if current_notes:
        cur = " | ".join(f"{d}: {(t or '').replace(chr(10), ' / ')}" for d, t in current_notes.items()) or "(keine)"
    return (_KI_PROMPT.replace("<GR>", gr).replace("<TR>", tr).replace("<GE>", ge)
            .replace("<NEXT>", next_date_str).replace("<WRITER>", writer or "?").replace("<CURRENT>", cur))

def ki_analyze(notiz, writer, next_date_str, token, current_notes=None):
    content = _ki_call([{"role": "system", "content": _ki_system_prompt(writer, next_date_str, current_notes)},
                        {"role": "user", "content": notiz}], token)
    return json.loads(content)

def _ki_full_name(short):
    s = (short or "").strip()
    if s in ALLE_TRAINER:
        return s
    return WEBSITE_TO_DISPLAY.get(s, s)

def _ki_time_min(s):
    m = re.search(r'(\d{1,2})[:.](\d{2})', s or "")
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
        if 0 <= h < 24 and 0 <= mi < 60:
            return h * 60 + mi
    return None

def ki_timing_to_dict(timing_list):
    out = {}
    for tm in (timing_list or []):
        name = _ki_full_name(tm.get("trainer", ""))
        if name not in ALLE_TRAINER:
            continue
        kind = "spaet" if tm.get("richtung") == "spaet" else "frueh"
        t = _ki_time_min(tm.get("uhrzeit"))
        blocked = compute_blocked_slots(kind, t)
        out[name] = {"kind": kind, "time_min": t,
                     "time_str": (f"{t//60:02d}:{t%60:02d}" if t is not None else None),
                     "notiz": "", "blocked": sorted(blocked)}
    return out

def _ki_next_training_after(d):
    d = d + timedelta(days=1)
    while d.weekday() not in (2, 4):
        d += timedelta(days=1)
    return d

def ki_resolve_target(ziel, state):
    ziel = ziel or {}
    typ = ziel.get("typ") or "naechstes"
    if typ == "datum" and ziel.get("datum"):
        for fmt in ("%d.%m.%Y", "%d.%m.%y"):
            try:
                return datetime.strptime(ziel["datum"], fmt).date().strftime("%d.%m.%y")
            except Exception:
                pass
    if typ == "naechstes_geraet" and ziel.get("geraet"):
        g = ziel["geraet"]
        try:
            idx, _ = get_next_gear(state, exclude_date=None)
        except Exception:
            idx = state.get("geraet_combo_index", 0)
        d = active_training_date()
        for _ in range(24):
            if g in GERAETE_ROTATION[idx % len(GERAETE_ROTATION)]:
                return d.strftime("%d.%m.%y")
            d = _ki_next_training_after(d)
            idx += 1
    return active_training_date().strftime("%d.%m.%y")

# ---- KI-Einteilung: Zusammenlegen / Zuteilen im Raster ----
_KI_GRPCOL = {"G1": "g1_blau", "G2": "g1_gruen", "G3": "g2_orange", "G4": "g2_lila"}

def _ki_label_cells(label):
    if not label:
        return None
    lab = str(label).strip()
    if lab.lower() == "springer":
        return [("Aufbauen", "aufbauen"), ("Springer", "springer"), ("Springer", "springer"),
                ("Springer", "springer"), ("Abbauen", "aufbauen")]
    first = re.split(r'[+/ ]', lab)[0]
    col = _KI_GRPCOL.get(first, "g1_blau")
    return [(f"AW {lab}", "aufwaermen"), (lab, col), (lab, col), (lab, col), ("Abbauen", "aufbauen")]

def _ki_norm_merge(label):
    parts = [p for p in re.split(r'\+', str(label or "").replace(" ", "")) if p in ("G1", "G2", "G3", "G4")]
    return "+".join(sorted(parts, key=lambda x: int(x[1]))) if len(parts) >= 2 else None

def _merge_small_singletons(groups, present, min_kids=3):
    """Legt zu kleine (<min_kids anwesende Kinder) benachbarte Einzelgruppen zusammen.
    Bevorzugt gleiche Geraeteseite; die G2|G3-Grenze ist letzter Ausweg.
    (Noah 16.07.2026: auch im KI-/Anmerkungs-Pfad automatisch mergen, z.B. G3<3 -> G3+G4.)"""
    units = [[g] for g in groups]
    def cnt(u): return sum(present.get(g, 0) for g in u)
    def crosses(a, b): return a[-1] == "G2" and b[0] == "G3"
    changed = True
    while changed:
        changed = False
        for i, u in enumerate(units):
            if cnt(u) < min_kids and len(units) > 1:
                cands = []
                if i > 0: cands.append(i - 1)
                if i < len(units) - 1: cands.append(i + 1)
                best = None
                for j in cands:
                    a, b = (units[j], units[i]) if j < i else (units[i], units[j])
                    sc = cnt(units[j]) + cnt(u) + (1000 if crosses(a, b) else 0)
                    if best is None or sc < best[0]:
                        best = (sc, j)
                if best:
                    j = best[1]; lo, hi = sorted((i, j))
                    units[lo] = units[lo] + units[hi]; del units[hi]
                    changed = True; break
    return units


def build_ki_einteilung(absences, ki, geraet_1, geraet_2, g1_starts_geraet2):
    """Isolierter Builder aus KI-Anweisungen (cancel/merges/assign) -> vollstaendiges Raster.
    Entfallene Gruppen -> kein eigener Trainer (Trainer wird Springer)."""
    abwesend = set(absences.get("Trainer", []))
    available = [t for t in ALLE_TRAINER if t not in abwesend]
    cancel = set(g for g in (ki.get("cancel") or []) if g in ("G1", "G2", "G3", "G4"))
    for g in ("G1", "G2", "G3", "G4"):
        tl = ALLE_TURNER.get(g, [])
        if tl and all(t in absences.get(g, []) for t in tl):
            cancel.add(g)
    base_groups = [g for g in ("G1", "G2", "G3", "G4") if g not in cancel]

    assign_raw = ki.get("assign") or []
    used_groups = set()
    merge_units = []
    for a in assign_raw:
        tr = _ki_full_name(a.get("trainer", ""))
        lab = _ki_norm_merge(a.get("gruppe") or a.get("label") or "")
        if tr in available and lab:
            grps = set(lab.split("+")) & set(base_groups)
            if len(grps) >= 2 and not (grps & used_groups):
                merge_units.append((grps, tr)); used_groups |= grps
    for mg in (ki.get("merges") or []):
        grps = set(g for g in (mg or []) if g in base_groups) - used_groups
        if len(grps) >= 2:
            merge_units.append((grps, None)); used_groups |= grps

    # G3+G4-Dauer-Zusammenlegung (siehe Toggle G3_G4_PERMANENT_MERGE oben):
    # gilt auch im KI-/Anmerkungs-Pfad als Standard, AUSSER eine Anmerkung
    # nennt G3 oder G4 explizit einzeln (z.B. "Fabian macht G3 alleine" oder
    # ein cancel/merge, das G3/G4 schon anders verplant hat).
    if G3_G4_PERMANENT_MERGE and {"G3", "G4"} <= set(base_groups) and not ({"G3", "G4"} & used_groups):
        explicit_single = {
            (a.get("gruppe") or a.get("label") or "").strip()
            for a in assign_raw
            if _ki_full_name(a.get("trainer", "")) in available
            and "+" not in (a.get("gruppe") or a.get("label") or "")
        }
        if "G3" not in explicit_single and "G4" not in explicit_single:
            merge_units.append(({"G3", "G4"}, None))
            used_groups |= {"G3", "G4"}

    forced = {}
    for a in assign_raw:
        tr = _ki_full_name(a.get("trainer", ""))
        rawlab = (a.get("gruppe") or a.get("label") or "").strip()
        if tr not in available or "+" in rawlab:
            continue
        if rawlab.lower() == "springer":
            forced[tr] = "Springer"
        elif rawlab in base_groups and rawlab not in used_groups:
            forced[tr] = rawlab; used_groups.add(rawlab)

    units = []
    for grps, tr in merge_units:
        units.append(("+".join(sorted(grps, key=lambda x: int(x[1]))), tr))
    # Uebrig gebliebene Einzelgruppen: zu kleine Gruppen (<3 anwesende Kinder) automatisch
    # mit bestem Nachbarn zusammenlegen (wie im Standard-Pfad build_trainer_plan).
    present = {g: max(0, len(ALLE_TURNER.get(g, [])) - len(absences.get(g, [])))
               for g in ("G1", "G2", "G3", "G4")}
    leftover = [g for g in base_groups if g not in used_groups]
    for grp_list in _merge_small_singletons(leftover, present):
        lbl = "+".join(sorted(grp_list, key=lambda x: int(x[1])))
        units.append((lbl, None))
        for g in grp_list:
            used_groups.add(g)

    assign = dict(forced)
    pool = [t for t in available if t not in assign]
    open_units = []
    for lab, tr in units:
        if tr and tr in pool:
            assign[tr] = lab; pool.remove(tr)
        else:
            open_units.append(lab)
    # Zu wenige Trainer fuer die offenen Gruppen -> benachbart zusammenlegen (volle Abdeckung)
    def _gnum(x):
        return int(x[1]) if len(x) > 1 and x[1].isdigit() else 9
    while len(open_units) > max(len(pool), 1) and len([u for u in open_units if "+" not in u]) >= 2:
        _s = sorted([u for u in open_units if "+" not in u], key=_gnum)
        _a, _b = _s[0], _s[1]
        open_units.remove(_a); open_units.remove(_b)
        open_units.append("+".join(sorted([_a, _b], key=_gnum)))
    g4u = next((u for u in open_units if "G4" in u), None)
    andy = next((t for t in pool if t.startswith("Andy")), None)
    if andy and g4u:
        assign[andy] = g4u; pool.remove(andy); open_units.remove(g4u)
    noah = next((t for t in pool if t.startswith("Noah")), None)
    ordered = ([noah] if noah else []) + [t for t in pool if t is not noah]
    for lab in list(open_units):
        if ordered:
            tr = ordered.pop(0); assign[tr] = lab; open_units.remove(lab)
    for t in ordered:
        assign.setdefault(t, "Springer")

    # Zellen mit korrekter Geraete-Rotation (Phasen-Farben wie Standard-Builder)
    def _farbe(g, phase):
        return geraet_farbe(g, phase, g1_starts_geraet2)
    def _cells(label):
        lab = str(label).strip()
        if lab.lower() == "springer":
            return [("Aufbauen","aufbauen"),("Springer","springer"),("Springer","springer"),("Springer","springer"),("Abbauen","aufbauen")]
        if lab == "G4":
            _g4s3, _g4s4 = geraet_farbe_g4(g1_starts_geraet2)
            return [("Aufbauen","aufbauen"),("AW G4","aufwaermen"),("AW G4","aufwaermen"),("G4",_g4s3),("G4",_g4s4)]
        if "+" not in lab and lab in ("G1","G2","G3"):
            return [("AW "+lab,"aufwaermen"),(lab,_farbe(lab,1)),(lab,_farbe(lab,1)),(lab,_farbe(lab,2)),("Abbauen","aufbauen")]
        # Zusammengelegte Einheit (z.B. "G1+G2", "G3+G4"): Farbe von der Referenz-
        # Gruppe uebernehmen statt von einem Listen-Index abhaengigen Hartcode
        # (Bugfix 19.08.2026 - siehe _merged_ref_group()/Vault "Geräte-Farben-Kollision").
        _parts = [p for p in re.split(r'[+/]', lab) if p in ("G1", "G2", "G3", "G4")]
        _ref = _merged_ref_group(_parts) if _parts else "G3"
        c1, c2 = _farbe(_ref, 1), _farbe(_ref, 2)
        return [("AW "+lab,"aufwaermen"),(lab,c1),(lab,c1),(lab,c2),("Abbauen","aufbauen")]
    TRAINER_PLAN = {t: (_cells(assign[t]) if t in assign else None) for t in ALLE_TRAINER}
    anm = []
    if open_units:
        anm.append("ACHTUNG: kein Trainer mehr fuer: " + ", ".join(open_units))
    return TRAINER_PLAN, {}, anm

def _ki_kid(name):
    disp = normalize_name((name or "").strip())
    for g, lst in ALLE_TURNER.items():
        if disp in lst:
            return g, disp
    return None

def _merge_ki_assign(existing, new):
    """Fuegt neue Zuteilungen zu bestehenden hinzu, statt sie zu ueberschreiben.
    Spaetere Eintraege fuer denselben Trainer gewinnen; andere Trainer bleiben erhalten.
    (Noah 16.07.2026: sonst warf eine zweite Anmerkung die erste raus, z.B. 'Noah Springer'.)"""
    by = {}
    order = []
    for a in list(existing or []) + list(new or []):
        if not isinstance(a, dict):
            continue
        tr = (a.get("trainer") or "").strip()
        gr = a.get("gruppe") or a.get("label")
        if not tr or not gr:
            continue
        if tr not in by:
            order.append(tr)
        by[tr] = gr
    return [{"trainer": t, "gruppe": by[t]} for t in order]


# ════════════════════════════════════════════════════════════════
#  DETERMINISTISCHER ANMERKUNGS-KOMMANDO-PARSER (Bug-Fix 19.08.2026)
# ════════════════════════════════════════════════════════════════
# Trainer schreiben Kurz-Kommandos wie "Noah Springer" oder "lösche alle
# Anmerkungen" in die Anmerkungs-Zeile - vorher landete das 1:1 als Text im
# naechsten Plan (statt als Anweisung ausgefuehrt zu werden), weil die
# KI-Interpretation entweder keinen GH_MODELS_TOKEN hatte oder scheiterte.
# Dieser Parser erkennt eine feste Liste kanonischer Muster deterministisch
# UND OHNE KI, wendet die Aktion direkt auf fixed_entries[Ziel-Datum] an und
# konsumiert die Anmerkung (markiert sie als gelesen). Nur wenn ALLE Zeilen
# einer Anmerkung sicher matchen, wird sie hier verarbeitet - sonst bleibt
# sie der KI/dem Copy-Paste-Fallback ueberlassen.
#
# Erkannte Muster (case-insensitive, jeweils EINE Zeile):
#   - "lösche/loesche/entferne (alle) (bisherigen) Anmerkungen"
#     "alle Anmerkungen löschen" / "reset" / "zurücksetzen"
#     -> reset: alle Vorgaben fuer das Ziel-Datum werden entfernt
#   - "<Trainer-Vorname> Springer"       (z.B. "Noah Springer")
#   - "<Trainer-Vorname> G1" / "G2" / "G3" / "G4"
#   - "<Trainer-Vorname> G1+G2"          (auch "G3+G4", "Gruppe 1+2", ...)
#   - "<Trainer-Vorname> abwesend/nicht da/fehlt"
#   - "<Trainer-Vorname> anwesend/wieder da/kommt"
#   - "GX entfaellt" / "Gruppe X entfaellt"
#   - "GX und GY zusammenlegen" / "GX+GY zusammenlegen"
#
# Erweiterbar: neues Regex + Handler unten in COMMAND_PATTERNS ergaenzen.

_TRAINER_FIRSTNAMES = ["Noah", "Andy", "Fabian", "Cassian", "Julian", "Torben"]

def _first_to_full(first):
    for full in ALLE_TRAINER:
        if full.split()[0].lower() == first.lower():
            return full
    return None

_RE_RESET = re.compile(
    r'^\s*(?:'
    r'(?:l(?:oe|ö)sch(?:e|t|en)?|entfern(?:e|t|en)?)\s+(?:alle\s+)?(?:bisherigen\s+)?anmerkungen'
    r'|alle\s+anmerkungen\s+(?:l(?:oe|ö)schen|entfernen)'
    r'|reset|zur(?:ue|ü)cksetzen|mach\s+alles\s+r(?:ue|ü)ckg(?:ae|ä)ngig'
    r')\s*[.!]?\s*$',
    re.IGNORECASE,
)

_RE_TRAINER_ROLE = re.compile(
    r'^\s*(?P<name>' + '|'.join(_TRAINER_FIRSTNAMES) + r')'
    r'\s+(?:macht\s+|ist\s+|:\s*)?'
    r'(?P<lab>Springer|Gruppe\s*[1-4](?:\s*\+\s*(?:Gruppe\s*)?[1-4])?|G[1-4](?:\s*\+\s*G?[1-4])?)'
    r'\s*[.!]?\s*$',
    re.IGNORECASE,
)

_RE_TRAINER_ABSENT = re.compile(
    r'^\s*(?P<name>' + '|'.join(_TRAINER_FIRSTNAMES) + r')'
    r'\s+(?:ist\s+)?(?:abwesend|nicht\s+da|fehlt|kann\s+nicht|f(?:ae|ä)llt\s+aus)'
    r'\s*[.!]?\s*$',
    re.IGNORECASE,
)

_RE_TRAINER_PRESENT = re.compile(
    r'^\s*(?P<name>' + '|'.join(_TRAINER_FIRSTNAMES) + r')'
    r'\s+(?:ist\s+)?(?:anwesend|wieder\s+da|kommt(?:\s+doch)?|doch\s+da)'
    r'\s*[.!]?\s*$',
    re.IGNORECASE,
)

_RE_GRUPPE_ENTFALL = re.compile(
    r'^\s*(?:Gruppe\s*)?G?(?P<g>[1-4])\s+(?:entf(?:ae|ä)llt|f(?:ae|ä)llt\s+aus)\s*[.!]?\s*$',
    re.IGNORECASE,
)

_RE_MERGE = re.compile(
    r'^\s*(?:Gruppe\s*)?G?(?P<a>[1-4])\s*(?:\+|und|u\.)\s*(?:Gruppe\s*)?G?(?P<b>[1-4])'
    r'\s+(?:zusammenlegen|zusammen|gemeinsam)\s*[.!]?\s*$',
    re.IGNORECASE,
)

def _norm_group_label(lab):
    """'g1', 'Gruppe 1', 'G3+G4', 'gruppe 3+4', 'Springer' -> kanonisch."""
    s = lab.strip()
    if re.match(r'^\s*springer\s*$', s, re.IGNORECASE):
        return "Springer"
    parts = re.split(r'\s*\+\s*', s)
    out = []
    for p in parts:
        m = re.match(r'^\s*(?:gruppe\s*)?g?\s*([1-4])\s*$', p, re.IGNORECASE)
        if not m:
            return None
        out.append("G" + m.group(1))
    if not out:
        return None
    if len(out) == 1:
        return out[0]
    return "+".join(sorted(out, key=lambda x: int(x[1])))

def _parse_command_line(line):
    """Versucht, EINE Anmerkungs-Zeile als bekanntes Kommando zu erkennen.
    Gibt ein dict {typ, ...} zurueck oder None."""
    s = (line or "").strip().rstrip(",;")
    if not s:
        return None
    if _RE_RESET.match(s):
        return {"typ": "reset"}
    m = _RE_TRAINER_ROLE.match(s)
    if m:
        full = _first_to_full(m.group("name"))
        lab = _norm_group_label(m.group("lab"))
        if full and lab:
            return {"typ": "assign", "trainer": full, "gruppe": lab}
    m = _RE_TRAINER_ABSENT.match(s)
    if m:
        full = _first_to_full(m.group("name"))
        if full:
            return {"typ": "trainer_abwesend", "trainer": full}
    m = _RE_TRAINER_PRESENT.match(s)
    if m:
        full = _first_to_full(m.group("name"))
        if full:
            return {"typ": "trainer_anwesend", "trainer": full}
    m = _RE_GRUPPE_ENTFALL.match(s)
    if m:
        return {"typ": "gruppe_entfall", "gruppe": "G" + m.group("g")}
    m = _RE_MERGE.match(s)
    if m:
        a, b = "G" + m.group("a"), "G" + m.group("b")
        if a != b:
            return {"typ": "merge", "gruppen": sorted([a, b], key=lambda x: int(x[1]))}
    return None

def _apply_command(cmd, fe):
    """Wendet ein geparstes Kommando auf einen fixed_entries[dk]-Eintrag an.
    fe wird in-place mutiert. Setzt manuell_bearbeitet=True damit der Admin-
    Pfad im main() den Trainer-Plan aus diesen Vorgaben neu baut."""
    typ = cmd.get("typ")
    if typ == "reset":
        # Alle Vorgaben fuer dieses Datum entfernen (analog KI reset: true)
        fe.clear()
        return
    fe.setdefault("manuell_bearbeitet", True)
    fe["quelle"] = fe.get("quelle") or "cmd"
    ki = fe.setdefault("ki", {})
    if typ == "assign":
        existing = ki.get("assign") or []
        existing = [x for x in existing if x.get("trainer") != cmd["trainer"]]
        existing.append({"trainer": cmd["trainer"], "gruppe": cmd["gruppe"]})
        ki["assign"] = existing
    elif typ == "merge":
        ex = [tuple(sorted(m, key=lambda x: int(x[1]))) for m in (ki.get("merges") or [])]
        new = tuple(cmd["gruppen"])
        if new not in ex:
            ex.append(new)
        ki["merges"] = [list(m) for m in ex]
    elif typ == "gruppe_entfall":
        ki["cancel"] = sorted(set((ki.get("cancel") or []) + [cmd["gruppe"]]))
        fa = fe.setdefault("fixed_absences", {})
        g = cmd["gruppe"]
        fa.setdefault(g, [])
        for t in ALLE_TURNER.get(g, []):
            if t not in fa[g]:
                fa[g].append(t)
    elif typ == "trainer_abwesend":
        fa = fe.setdefault("fixed_absences", {})
        fa.setdefault("Trainer", [])
        if cmd["trainer"] not in fa["Trainer"]:
            fa["Trainer"].append(cmd["trainer"])
    elif typ == "trainer_anwesend":
        fa = fe.setdefault("fixed_absences", {})
        if cmd["trainer"] in fa.get("Trainer", []):
            fa["Trainer"].remove(cmd["trainer"])

def preprocess_deterministic_annotations(sftp, anmerkungen_server, fixed_entries):
    """Erkennt deterministisch parsbare Kommando-Anmerkungen, wendet sie auf
    fixed_entries[naechstes_Datum] an und konsumiert sie (markiert gelesen).
    Laeuft VOR ki_process_annotations - so klappt die Interpretation der
    einfachen Faelle auch ohne GH_MODELS_TOKEN. Alles was nicht sauber matcht,
    bleibt fuer die KI/den Copy-Paste-Pfad liegen."""
    if not anmerkungen_server:
        return
    dk = active_training_date().strftime("%d.%m.%y")
    changed = False
    processed_ids = []
    consumed = []
    for a in list(anmerkungen_server):
        notiz = (a.get("notiz") or "").strip()
        if not notiz:
            continue
        lines = [ln for ln in notiz.splitlines() if ln.strip()]
        if not lines:
            continue
        parsed = [_parse_command_line(ln) for ln in lines]
        if any(p is None for p in parsed):
            continue  # gemischt -> lieber der KI ueberlassen
        # Alle Zeilen sind Kommandos -> anwenden
        fe = fixed_entries.setdefault(dk, {})
        for p in parsed:
            _apply_command(p, fe)
        # Nach reset() kann fe leer sein -> Eintrag komplett verwerfen
        if not fe:
            fixed_entries.pop(dk, None)
        changed = True
        consumed.append(a)
        if a.get("id"):
            processed_ids.append(a["id"])
        writer = a.get("trainer", "?")
        print(f"[CMD] '{notiz[:60]}' ({writer}) -> {dk}  {[p['typ'] for p in parsed]}")
    for a in consumed:
        try:
            anmerkungen_server.remove(a)
        except ValueError:
            pass
    if changed:
        try:
            f = sftp.open("fixed_entries.json", "wb")
            f.write(json.dumps(fixed_entries, indent=2, ensure_ascii=False).encode("utf-8"))
            f.close()
            print("[CMD] fixed_entries.json aktualisiert.")
        except Exception as e:
            print(f"[CMD] fixed_entries Upload fehlgeschlagen: {e}")
    if processed_ids:
        try:
            mark_anmerkungen_gelesen(sftp, processed_ids)
        except Exception as e:
            print(f"[CMD] mark gelesen fehlgeschlagen: {e}")


def ki_process_annotations(sftp, anmerkungen_server, fixed_entries, state):
    """Analysiert jede ungelesene Anmerkung per KI und schreibt die Aenderung in
    fixed_entries[Zieldatum] (Admin-Mechanismus). Einmalig pro Datum."""
    token = _ki_token()
    if not token or not anmerkungen_server:
        return
    next_date_str = active_training_date().strftime("%d.%m.%Y")
    allowed_ger = {g for c in GERAETE_ROTATION for g in c}
    processed_ids = []
    changed = False
    current_notes = {d: (fe.get("notiz") or "") for d, fe in fixed_entries.items()
                     if isinstance(fe, dict) and (fe.get("notiz") or "").strip()}
    for a in list(anmerkungen_server):
        notiz = (a.get("notiz") or "").strip()
        writer = a.get("trainer", "")
        if not notiz:
            continue
        try:
            res = ki_analyze(notiz, writer, next_date_str, token, current_notes)
        except Exception as e:
            print(f"[KI] Analyse fehlgeschlagen ({writer}): {e}")
            continue
        try:
            dk = ki_resolve_target(res.get("ziel"), state)
            if res.get("reset"):
                fixed_entries.pop(dk, None)
                if a.get("id"):
                    processed_ids.append(a["id"])
                if a in anmerkungen_server:
                    anmerkungen_server.remove(a)
                changed = True
                print(f"[KI] RESET {dk} (alle Vorgaben entfernt)")
                continue
            fe = fixed_entries.setdefault(dk, {})
            fe["manuell_bearbeitet"] = True
            fe["quelle"] = "ki"
            ki = fe.setdefault("ki", {})
            _nn = res.get("notiz_neu")
            if _nn is not None:
                fe["notiz"] = str(_nn).strip()   # reine Anmerkungs-Textbearbeitung (Struktur bleibt)
                if a.get("id"):
                    processed_ids.append(a["id"])
                if a in anmerkungen_server:
                    anmerkungen_server.remove(a)
                changed = True
                print(f"[KI] Anmerkungs-Text bearbeitet -> {dk}")
                continue
            note_lines = []
            ge = res.get("geraete")
            if ge and ge.get("g1") in allowed_ger and ge.get("g2") in allowed_ger:
                fe["geraet_1"] = ge["g1"]; fe["geraet_2"] = ge["g2"]
            fa = fe.setdefault("fixed_absences", {})
            ent = [g for g in (res.get("gruppe_entfall") or []) if g in ALLE_TURNER]
            if ent:
                ki["cancel"] = sorted(set((ki.get("cancel") or []) + ent))
                for g in ent:
                    fa.setdefault(g, [])
                    for t in ALLE_TURNER[g]:
                        if t not in fa[g]:
                            fa[g].append(t)
                    note_lines.append(f"{g} entfaellt")
            if res.get("timing"):
                ki["timing"] = res["timing"]
                for tm in res["timing"]:
                    ri = "kommt spaeter" if tm.get("richtung") == "spaet" else "geht frueher"
                    uh = tm.get("uhrzeit")
                    note_lines.append(f"{tm.get('trainer','')} {ri}" + (f" {uh}" if uh else ""))
            if res.get("zuteilung"):
                ki["assign"] = _merge_ki_assign(ki.get("assign"), res["zuteilung"])
                for z in res["zuteilung"]:
                    if z.get("trainer") and z.get("gruppe"):
                        note_lines.append(f"{z['trainer']}: {z['gruppe']}")
            if res.get("merges"):
                ki["merges"] = res["merges"]
                for mg in res["merges"]:
                    if isinstance(mg, list) and len(mg) >= 2:
                        note_lines.append("+".join(mg) + " zusammen")
            for kn in (res.get("abwesend_kinder") or []):
                hit = _ki_kid(kn)
                if hit:
                    g, disp = hit
                    fa.setdefault(g, [])
                    if disp not in fa[g]:
                        fa[g].append(disp)
                    note_lines.append(f"{disp} abgemeldet")
            for kn in (res.get("wieder_da_kinder") or []):
                hit = _ki_kid(kn)
                if hit:
                    g, disp = hit
                    if disp in fa.get(g, []):
                        fa[g].remove(disp)
                    note_lines.append(f"{disp} wieder da")
            for tn in (res.get("abwesend_trainer") or []):
                d = normalize_name((tn or "").strip())
                if d in ALLE_TRAINER:
                    fa.setdefault("Trainer", [])
                    if d not in fa["Trainer"]:
                        fa["Trainer"].append(d)
                    note_lines.append(f"{d} nicht da")
            eint = res.get("einteilung") or {}
            if isinstance(eint, dict) and eint:
                val = []
                for tr, lab in eint.items():
                    d = normalize_name((tr or "").strip())
                    if d in ALLE_TRAINER and lab:
                        val.append({"trainer": d, "gruppe": str(lab)})
                if val:
                    ki["assign"] = _merge_ki_assign(ki.get("assign"), val)
            # Umgesetzte Anweisungen werden NICHT als Anmerkungs-Text in den Plan geschrieben
            # (Noah 16.07.2026). Nur echter Hinweistext ("In die Anmerkungen schreiben: ...")
            # landet als Notiz; note_lines dienen nur noch dem Log unten.
            base_note = (res.get("notiz") or "").strip()
            if base_note:
                prev = (fe.get("notiz") or "").strip()
                fe["notiz"] = (prev + "\n" if prev else "") + base_note
            if a.get("id"):
                processed_ids.append(a["id"])
            if a in anmerkungen_server:
                anmerkungen_server.remove(a)
            changed = True
            print(f"[KI] '{notiz[:45]}' ({writer}) -> {dk}")
        except Exception as e:
            print(f"[KI] Anwenden fehlgeschlagen: {e}")
    if changed:
        try:
            f = sftp.open("fixed_entries.json", "wb")
            f.write(json.dumps(fixed_entries, indent=2, ensure_ascii=False).encode("utf-8"))
            f.close()
            print("[KI] fixed_entries.json aktualisiert.")
        except Exception as e:
            print(f"[KI] fixed_entries Upload fehlgeschlagen: {e}")
    if processed_ids:
        try:
            mark_anmerkungen_gelesen(sftp, [i for i in processed_ids if i])
        except Exception as e:
            print(f"[KI] mark gelesen fehlgeschlagen: {e}")


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
    training_date = active_training_date()
    datum         = fmt_datum(training_date)
    datum_kurz    = fmt_datum_kurz(training_date)
    wtag          = wochentag_name(training_date)
    days_away     = (training_date - today).days

    print(f"\nNaechstes Training: {wtag} {datum} ({days_away} Tage)\n")

    # Abwesenheiten und Hash berechnen
    absences, late_notes, trainer_timing = get_absences(abmeldungen, training_date)
    late_notes = late_notes + timing_annotations(trainer_timing)
    # raw_abm_hash für konsistenten Vergleich mit check_quick.py (beide nutzen raw JSON hash)
    new_hash = raw_abm_hash

    # -- Deterministischer Kommando-Parser (Bug-Fix 19.08.2026): erkennt einfache
    # Kurz-Kommandos wie "Noah Springer" oder "lösche alle Anmerkungen" OHNE KI
    # und wendet sie direkt an. Was hier nicht matcht, bleibt fuer die KI. --
    try:
        preprocess_deterministic_annotations(sftp, anmerkungen_server, fixed_entries)
    except Exception as _cmde:
        print(f"[CMD] uebersprungen: {_cmde}")

    # -- Mini-KI: freie Trainer-Anmerkungen analysieren und in fixed_entries anwenden --
    try:
        ki_process_annotations(sftp, anmerkungen_server, fixed_entries, state)
    except Exception as _kie:
        print(f"[KI] uebersprungen: {_kie}")

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
        _ef_notiz = (fixed_for_date.get("notiz", "") or "").strip()
        _ef_nhash = hashlib.md5(_ef_notiz.encode("utf-8")).hexdigest()
        _notiz_store = state.setdefault("entfall_notiz_hash", {})
        _notiz_changed = _notiz_store.get(datum_kurz) != _ef_nhash
        _was_published = datum_kurz in entfall_published
        if _was_published and plan_exists(sftp, datum_kurz) and not _notiz_changed:
            print(f"[ENTFALL] {datum} bereits als Entfall veröffentlicht – nichts zu tun.")
            sftp.close(); ssh.close()
            return
        publish_entfall(sftp, datum, datum_kurz, wtag, _ef_notiz)
        _notiz_store[datum_kurz] = _ef_nhash
        if datum_kurz not in entfall_published:
            entfall_published.append(datum_kurz)
        if datum_kurz not in state.setdefault("generated_plans", []):
            state["generated_plans"].append(datum_kurz)
        state.get("plan_data", {}).pop(datum_kurz, None)  # alten Plan-Hash verwerfen
        save_state(sftp, state)
        if not _was_published:
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
        state.setdefault("entfall_notiz_hash", {}).pop(datum_kurz, None)
        remove_plan_files(sftp, datum_kurz)
        state.get("plan_data", {}).pop(datum_kurz, None)
        if datum_kurz in state.get("generated_plans", []):
            state["generated_plans"].remove(datum_kurz)
        save_state(sftp, state)

    # -- KI-Timing (aus Anmerkung) in trainer_timing mergen -> in allen Pfaden rot geblockt --
    try:
        for _kn, _kv in ki_timing_to_dict((fixed_for_date.get("ki") or {}).get("timing")).items():
            trainer_timing[_kn] = _kv
    except Exception:
        pass

    # -- Admin-Editor (manuell_bearbeitet aus admin.php) anwenden --------------
    force_regen = False
    admin_fixed_hash = ""
    if fixed_for_date.get("manuell_bearbeitet"):
        import hashlib as _hl
        admin_fixed_hash = _hl.md5(json.dumps(fixed_for_date, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()
        # Geraete/Rotation KONSISTENT zur eigentlichen Generierung bestimmen:
        # explizit gesetzte Geraete nutzen, sonst die KORREKTE Rotation fuer diesen Tag
        # (NICHT stumpf Boden/Barren) -> KI-/Admin-Plaene haben immer die richtigen Geraete.
        try:
            _ridx, _rg1s = get_next_gear(state, exclude_date=datum_kurz, fixed_entries=fixed_entries)
            _rg1, _rg2 = GERAETE_ROTATION[_ridx]
        except Exception:
            _rg1, _rg2, _rg1s = "Boden", "Barren", state.get("g1_starts_geraet2", False)
        _g1 = fixed_for_date.get("geraet_1") or _rg1
        _g2 = fixed_for_date.get("geraet_2") or _rg2
        _g1s = fixed_for_date.get("g1_starts_geraet2", _rg1s)
        _ki = fixed_for_date.get("ki") or {}
        _partial = fixed_for_date.get("fixed_trainer_partial") or {}
        if _partial:
            # Bugfix 19.08.2026 (Noah): Admin-Zellen sind eine harte Vorgabe, die KI/
            # Kommando-Anweisungen (ki.assign/merges) planen - falls vorhanden - den
            # Rest drumherum, statt komplett ignoriert zu werden (siehe build_admin_trainer_plan).
            _base_tp = build_admin_trainer_plan(absences, _g1, _g2, _g1s, _partial, ki=_ki)
        elif _ki.get("assign") or _ki.get("merges"):
            try:
                _base_tp, _sd, _ka = build_ki_einteilung(absences, _ki, _g1, _g2, _g1s)
            except Exception as _kierr:
                print(f"[KI] Einteilung fehlgeschlagen, Fallback build_admin: {_kierr}")
                _base_tp = build_admin_trainer_plan(absences, _g1, _g2, _g1s, _partial, ki=_ki)
        else:
            _base_tp = build_admin_trainer_plan(absences, _g1, _g2, _g1s, _partial, ki=_ki)
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
            _ng_idx, _ng_g1s = get_next_gear(state, exclude_date=datum_kurz, fixed_entries=fixed_entries)
            geraet_combo = GERAETE_ROTATION[_ng_idx]
            # plan_data mit Defaults initialisieren und stored_hash="" setzen
            # damit die Update-Logik unten sicher ausgeführt wird
            plan_data = {
                "absences_hash":    "",   # leer → abs_changed=True → Plan wird generiert
                "trainer_absences": list(absences.get("Trainer", [])),
                "stored_absences":  {},
                "geraet_1":         geraet_combo[0],
                "geraet_2":         geraet_combo[1],
                "g1_starts_geraet2": _ng_g1s,
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

        # Kurze Hinweise sammeln -> am Ende EINE konsolidierte Mail statt mehrerer
        hinweise = []

        if trainer_changed:
            added   = new_trainer_abs - stored_trainer_abs
            removed = stored_trainer_abs - new_trainer_abs
            wer = ', '.join(sorted(added | removed)) or '–'
            hinweise.append(f"die Trainereinteilung wurde geändert, da {wer} abgemeldet ist")

        # UPDATE: Plan wird IMMER automatisch (neu) erstellt - auch bei Trainerwechsel
        # oder Engpass. Noah muss dafuer nichts mehr manuell machen (07.07.2026).
        print("UPDATE: erstelle Plan automatisch neu...")

        geraet_1     = plan_data["geraet_1"]
        geraet_2     = plan_data["geraet_2"]
        g1_starts_g2 = plan_data["g1_starts_geraet2"]

        issues, anwesend_trainer, soft_warnings = detect_complex(absences, late_notes)
        # Bei gesperrtem Trainer-Plan: Trainer-Anzahl-Fehler ignorieren
        if lock_trainer:
            issues = [i for i in issues if "Trainer anwesend" not in i]
        if issues:
            # Auch bei Problemen: trotzdem automatisch fertigstellen, nur kurz vermerken
            hinweise.append("beim automatischen Erstellen gab es ein Problem, bitte kurz prüfen")
            print(f"[WARN] Issues trotzdem automatisch weitergebaut: {issues}")

        if lock_trainer and fixed_tp_raw:
            # Gesperrter Trainer-Plan: direkt aus fixed_entries (UPDATE path)
            trainer_plan = {
                k: [tuple(s) for s in v] if v is not None else None
                for k, v in fixed_tp_raw.items()
            }
            sondertiming = {}
            anmerkungen  = []
            trainer_plan = apply_timing_coverage(trainer_plan, trainer_timing)
            trainer_plan = apply_timing_blocks(trainer_plan, trainer_timing)
            print("[FIXED] Verwende gesperrten Trainer-Plan aus fixed_entries (UPDATE).")
        else:
            _prev_roles = _load_previous_trainer_roles(state, exclude_date=datum_kurz)
            trainer_plan, sondertiming, anmerkungen = build_trainer_plan(
                absences, geraet_1, geraet_2, g1_starts_g2, trainer_timing,
                prev_trainer_roles=_prev_roles,
            )
            trainer_plan = apply_timing_coverage(trainer_plan, trainer_timing)
            trainer_plan = apply_timing_blocks(trainer_plan, trainer_timing)

        # Verspätungen / frühes Gehen als Hinweis eintragen
        for note in late_notes:
            anmerkungen.append(note)

        # Trainer-Anmerkungen eintragen; Admin-"Hinweis" separat als Override
        _admin_notiz = ""
        for anm in anmerkungen_server:
            trainer_name = anm.get("trainer", "")
            notiz        = anm.get("notiz", "").strip()
            if not notiz:
                continue
            if trainer_name == "Hinweis":
                _admin_notiz = notiz
                continue
            anmerkungen.append(f"• {trainer_name}: {notiz}")

        # Auto-Anmerkungen fuer Admin-Vorbefuellung sichern (vor evtl. Override)
        upload_anmerkungen_auto(sftp, datum_kurz, anmerkungen)

        # Admin-Override: editierte Notiz ersetzt die Anmerkungen 1:1 (kein Doppeln)
        if _admin_notiz:
            anmerkungen = [ln.rstrip() for ln in _admin_notiz.split("\n") if ln.strip()]

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

        # Dedup: soft warnings (≤2 Turner in Gruppe, Engpass, ...)
        already_sent_warnings = set(plan_data.get("warnings_sent", []))
        new_soft_warnings     = [(k, t) for k, t in soft_warnings if k not in already_sent_warnings]

        # Dedup: Trainer-Anmerkungen (nur neue IDs sollen einen Hinweis ausloesen)
        already_sent_anm_ids  = set(plan_data.get("anm_notified_ids", []))
        new_anm_ids           = [a["id"] for a in anmerkungen_server if a.get("id") and a["id"] not in already_sent_anm_ids]

        # State: Hash + dedup-Listen aktualisieren (auch trainer_absences - wichtig,
        # da dieser Block jetzt auch bei Trainerwechsel laeuft)
        plan_data["absences_hash"]    = new_hash
        plan_data["trainer_absences"] = list(absences.get("Trainer", []))
        # Trainer-Rollen fuer Rotation im naechsten Plan speichern (Bug-Fix 19.08.2026)
        plan_data["trainer_roles"]    = _extract_trainer_roles(trainer_plan)
        if admin_fixed_hash:
            plan_data["fixed_hash"] = admin_fixed_hash
        plan_data["stored_absences"] = absences
        if new_late_notes:
            plan_data["late_notes_sent"] = list(already_sent_notes | set(new_late_notes))
        if new_soft_warnings:
            plan_data["warnings_sent"] = list(already_sent_warnings | {k for k, _ in new_soft_warnings})
        if new_anm_ids:
            plan_data["anm_notified_ids"] = list(already_sent_anm_ids | set(new_anm_ids))
        save_state(sftp, state)

        # Weitere kurze Hinweise sammeln - NUR was wirklich eine Aktion/Aenderung war
        # (Trainerwechsel, Engpass-Zusammenlegung, Trainer-Anmerkung, Verspaetung).
        # Reine Turner-bezogene Warnungen (z.B. "nur 2 Turner in Gruppe X") ohne
        # Handlungsbedarf loesen bewusst KEINE Mail aus - nur Log (Noah, 07.07.2026).
        if new_anm_ids:
            hinweise.append("es gibt neue Trainer-Anmerkungen")
        if new_late_notes:
            hinweise.append("es gibt neue Verspätungs-/Frühgeh-Hinweise")
        low_trainer_new = [(k, t) for k, t in new_soft_warnings if k.startswith("low_trainer_")]
        if low_trainer_new:
            hinweise.append("nur wenige Trainer da, Gruppen wurden zusammengelegt")
        other_new_warnings = [t for k, t in new_soft_warnings if not k.startswith("low_trainer_")]
        if other_new_warnings:
            print(f"[INFO] Reine Turner-Hinweise fuer {datum} (keine Mail, kein Handlungsbedarf): {other_new_warnings}")

        # EINE konsolidierte, kurze Mail - nur wenn es etwas zu berichten gibt.
        # Reine Turner-Abwesenheitsaenderungen ohne weitere Auffaelligkeit -> keine Mail.
        if hinweise:
            send_whatsapp(
                f"Hallo, der Trainingsplan für {datum} wurde automatisch aktualisiert: "
                + "; ".join(hinweise) + ". Bitte bei Gelegenheit kurz prüfen."
            )
            print(f"UPDATE FERTIG fuer {datum}. Mail gesendet: {hinweise}")
        else:
            print(f"UPDATE FERTIG fuer {datum}. Keine Mail (nur Turner-Abwesenheiten geändert).")

    else:
        # ── Kein Plan vorhanden: erstelle neuen ────────────────
        # NUR im Publikationsfenster (Mi oder Fr nach 22:00 CEST / 20:00 UTC) erstellen
        now_utc   = datetime.now(timezone.utc)
        days_away = (training_date - date.today()).days

        # Sicherheitsnetz: Steht das naechste Training unmittelbar bevor (<=1 Tag)
        # und existiert noch KEIN Plan, wird sofort erzeugt - auch ausserhalb des
        # 22:00-Fensters. So bleibt kein Training ohne Plan, falls das regulaere
        # Mi/Fr-Abendfenster verpasst wurde.
        imminent = days_away <= 1
        if not is_publication_window() and not force_regen and not imminent:
            print(
                f"Kein Plan vorhanden, aber außerhalb Publikationsfenster "
                f"({now_utc.strftime('%H:%M')} UTC, Wochentag {now_utc.weekday()}) → warte bis Mi/Fr 22:00 CEST."
            )
            sftp.close(); ssh.close()
            return
        if imminent and not is_publication_window() and not force_regen:
            print(f"[SICHERHEITSNETZ] Training in {days_away} Tag(en), kein Plan → erstelle sofort (Fenster uebergangen).")

        if days_away > 5 and not force_regen:
            print(
                f"Kein Plan vorhanden, Training in {days_away} Tagen "
                f"({now_utc.strftime('%H:%M')} UTC) → noch zu weit entfernt."
            )
            sftp.close(); ssh.close()
            return

        print(f"Kein Plan vorhanden, starte Generierung... (Training in {days_away} Tagen, {now_utc.strftime('%H:%M')} UTC)")

        # Geräte-Rotation: letzte (bis zu) 8 ECHTEN Trainings ansehen, ausgefallene
        # ignorieren und genau einen Schritt weiterrücken (robust gegen Drift).
        new_combo_idx, g1_starts_g2 = get_next_gear(state, exclude_date=datum_kurz, fixed_entries=fixed_entries)
        geraet_1, geraet_2 = GERAETE_ROTATION[new_combo_idx]

        # Bei gesperrtem Trainer-Plan: Geräte aus fixed_entries übernehmen (falls vorhanden)
        if lock_trainer and fixed_for_date.get("geraet_1"):
            geraet_1     = fixed_for_date["geraet_1"]
            geraet_2     = fixed_for_date["geraet_2"]
            g1_starts_g2 = fixed_for_date.get("g1_starts_geraet2", g1_starts_g2)
            print(f"[FIXED] Geräte aus fixed_entries: {geraet_1} + {geraet_2}, G1 starts G2: {g1_starts_g2}")

        # Kurze Hinweise sammeln -> am Ende EINE konsolidierte Mail statt mehrerer
        hinweise = []

        issues, anwesend_trainer, soft_warnings = detect_complex(absences, late_notes)
        # Bei gesperrtem Trainer-Plan: Trainer-Anzahl-Fehler ignorieren (NEW path)
        if lock_trainer:
            issues = [i for i in issues if "Trainer anwesend" not in i]

        if issues:
            # Auch bei Problemen: trotzdem automatisch fertigstellen, nur kurz vermerken
            hinweise.append("beim automatischen Erstellen gab es ein Problem, bitte kurz prüfen")
            print(f"[WARN] Issues trotzdem automatisch weitergebaut: {issues}")

        if lock_trainer and fixed_tp_raw:
            # Gesperrter Trainer-Plan: direkt aus fixed_entries (NEW path)
            trainer_plan = {
                k: [tuple(s) for s in v] if v is not None else None
                for k, v in fixed_tp_raw.items()
            }
            sondertiming = {}
            anmerkungen  = []
            trainer_plan = apply_timing_coverage(trainer_plan, trainer_timing)
            trainer_plan = apply_timing_blocks(trainer_plan, trainer_timing)
            print("[FIXED] Verwende gesperrten Trainer-Plan aus fixed_entries (NEW).")
        else:
            _prev_roles = _load_previous_trainer_roles(state, exclude_date=datum_kurz)
            trainer_plan, sondertiming, anmerkungen = build_trainer_plan(
                absences, geraet_1, geraet_2, g1_starts_g2, trainer_timing,
                prev_trainer_roles=_prev_roles,
            )
            trainer_plan = apply_timing_coverage(trainer_plan, trainer_timing)
            trainer_plan = apply_timing_blocks(trainer_plan, trainer_timing)

        # Verspätungen / frühes Gehen als Hinweis eintragen
        for note in late_notes:
            anmerkungen.append(note)

        # Trainer-Anmerkungen eintragen; Admin-"Hinweis" separat als Override
        _admin_notiz = ""
        for anm in anmerkungen_server:
            trainer_name = anm.get("trainer", "")
            notiz        = anm.get("notiz", "").strip()
            if not notiz:
                continue
            if trainer_name == "Hinweis":
                _admin_notiz = notiz
                continue
            anmerkungen.append(f"• {trainer_name}: {notiz}")

        # Auto-Anmerkungen fuer Admin-Vorbefuellung sichern (vor evtl. Override)
        upload_anmerkungen_auto(sftp, datum_kurz, anmerkungen)

        # Admin-Override: editierte Notiz ersetzt die Anmerkungen 1:1 (kein Doppeln)
        if _admin_notiz:
            anmerkungen = [ln.rstrip() for ln in _admin_notiz.split("\n") if ln.strip()]

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
            "trainer_roles":    _extract_trainer_roles(trainer_plan),  # fuer Rotation im naechsten Plan
            "stored_absences":  absences,
            "geraet_1":         geraet_1,
            "geraet_2":         geraet_2,
            "g1_starts_geraet2": g1_starts_g2,
            "late_notes_sent":  list(late_notes),            # bereits gesendet beim Erstellen
            "warnings_sent":    [k for k, _ in soft_warnings],  # bereits gesendet beim Erstellen
            "anm_notified_ids": [a["id"] for a in anmerkungen_server if a.get("id")],  # bereits gemeldet
            "fixed_hash":       admin_fixed_hash,
        }
        save_state(sftp, state)

        # Kurze Hinweise sammeln - fliessen in die EINE Erstellungs-Mail mit ein.
        # Reine Turner-bezogene Warnungen ohne Handlungsbedarf -> nur Log, keine Erwaehnung.
        if anmerkungen_server:
            hinweise.append("es gibt neue Trainer-Anmerkungen")
        if late_notes:
            hinweise.append("es gibt Verspätungs-/Frühgeh-Hinweise")
        low_trainer_warn = [(k, t) for k, t in soft_warnings if k.startswith("low_trainer_")]
        if low_trainer_warn:
            hinweise.append("nur wenige Trainer da, Gruppen wurden zusammengelegt")
        other_warnings = [t for k, t in soft_warnings if not k.startswith("low_trainer_")]
        if other_warnings:
            print(f"[INFO] Reine Turner-Hinweise fuer {datum} (keine Erwaehnung, kein Handlungsbedarf): {other_warnings}")

        # Haupt-Notification: Plan fertig (immer genau EINE Mail, kurz, mit Hinweisen falls vorhanden)
        text = f"Hallo, der Trainingsplan für {wtag}, {datum} ist erstellt."
        if hinweise:
            text += " " + "; ".join(hinweise).capitalize() + ". Bitte bei Gelegenheit kurz prüfen."
        send_whatsapp(text)
        print(f"FERTIG! Plan fuer {datum} hochgeladen. Hinweise: {hinweise or 'keine'}")

    sftp.close()
    ssh.close()

if __name__ == "__main__":
    main()
# (Trainingsentfall-Support 19.06.2026)
