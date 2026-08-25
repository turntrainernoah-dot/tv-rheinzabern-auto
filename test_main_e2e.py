#!/usr/bin/env python3
"""End-to-End-Smoketest fuer main(): mockt SFTP komplett, prueft dass ein
kompletter Lauf (neuer Plan, 3 Gruppen, 4 Trainer, keine Abwesenheiten) ohne
Exception durchlaeuft und ein plausibles Ergebnis (Excel/PDF/JSON) erzeugt."""
import sys, os, json, types
sys.path.insert(0, "/home/user/tv-rheinzabern-auto")
import auto_trainingsplan as a
from datetime import date, timedelta

class FakeFile:
    def __init__(self, store, key, mode):
        self.store, self.key, self.mode = store, key, mode
        self._buf = store.get(key, b"")
    def read(self):
        return self._buf
    def write(self, data):
        self.store[self.key] = data
    def close(self):
        pass

class FakeSFTP:
    def __init__(self, files):
        self.files = files  # {path: bytes}
        self.put_calls = []
    def open(self, path, mode="r"):
        if "r" in mode and path not in self.files:
            raise FileNotFoundError(path)
        return FakeFile(self.files, path, mode)
    def stat(self, path):
        if path not in self.files:
            raise FileNotFoundError(path)
        return True
    def remove(self, path):
        self.files.pop(path, None)
    def put(self, local_path, remote_path):
        self.put_calls.append((local_path, remote_path))
        with open(local_path, "rb") as f:
            self.files[remote_path] = f.read()
    def close(self):
        pass

class FakeSSH:
    def close(self):
        pass

config = {
    "gruppen": [{"name": g, "zeiten": a._default_zeiten_for(g)} for g in ("G1", "G2", "G3")],
    "personen": [
        {"id": "t01", "name_intern": "Trainer Eins", "anzeige": "Trainer Eins", "rolle": "trainer", "immer_springer": False},
        {"id": "t02", "name_intern": "Trainer Zwei", "anzeige": "Trainer Zwei", "rolle": "trainer", "immer_springer": False},
        {"id": "t03", "name_intern": "Trainer Drei", "anzeige": "Trainer Drei", "rolle": "trainer", "immer_springer": True},
        {"id": "t04", "name_intern": "Trainer Vier", "anzeige": "Trainer Vier", "rolle": "trainer", "immer_springer": False},
    ] + [
        {"id": f"p{i:02d}", "name_intern": f"Kind {i}", "anzeige": f"Kind {i}", "rolle": "turner", "gruppe": g}
        for i, g in enumerate([g for g in ("G1", "G2", "G3") for _ in range(5)], start=1)
    ],
}

files = {
    "config/config.json": json.dumps(config).encode("utf-8"),
    "abmeldungen/abmeldungen.json": b"[]",
    "anmerkungen/anmerkungen.json": b"[]",
}

fake_sftp = FakeSFTP(files)
a.get_sftp = lambda: (FakeSSH(), fake_sftp)
a.send_whatsapp = lambda text: print(f"[WA-MOCK] {text[:80]}")

# LibreOffice ist in dieser Sandbox nicht lauffaehig (bestaetigt: bricht schon
# bei einer trivialen frischen xlsx, unabhaengig von diesem Code) -- fuer den
# End-to-End-Test wird nur der PDF-Konvertierungsschritt simuliert, damit der
# Rest von main() (Upload/State-Persistenz) mitgeprueft werden kann.
def fake_build_excel(*args, **kwargs):
    kwargs2 = dict(zip(
        ["datum","wochentag","geraet_1","geraet_2","abwesend","trainer_plan","sondertiming","anmerkungen","grid_rows"],
        args)) | kwargs
    datum_kurz = kwargs2["datum"][0:2] + "." + kwargs2["datum"][3:5] + "." + kwargs2["datum"][8:10]
    out_dir = "/tmp/trainingsplan"; os.makedirs(out_dir, exist_ok=True)
    xlsx_path = os.path.join(out_dir, f"{datum_kurz}_Trainingsplan.xlsx")
    pdf_path = os.path.join(out_dir, f"{datum_kurz}_Trainingsplan.pdf")
    with open(xlsx_path, "wb") as f: f.write(b"FAKE-XLSX")
    with open(pdf_path, "wb") as f: f.write(b"FAKE-PDF")
    return xlsx_path, pdf_path
a.build_excel = fake_build_excel

# Naechsten Trainingstag erzwingen, damit main() garantiert einen neuen Plan baut
def fake_active_training_date():
    d = date.today() + timedelta(days=1)
    while d.weekday() not in (2, 4):
        d += timedelta(days=1)
    return d
a.active_training_date = fake_active_training_date
a.is_publication_window = lambda: True

try:
    a.main()
    print("\n[OK] main() durchgelaufen ohne Exception.")
except SystemExit as e:
    print(f"\n[FAIL] main() hat SystemExit ausgeloest: {e}")
    sys.exit(1)

datum_kurz = fake_active_training_date().strftime("%d.%m.%y")
pdf_key = f"trainingspläne/{datum_kurz}_Trainingsplan.pdf"
xlsx_key = f"trainingspläne/{datum_kurz}_Trainingsplan.xlsx"
json_key = "trainingspläne/trainingsplan_aktuell.json"

ok = True
for key in (pdf_key, xlsx_key, json_key):
    present = key in fake_sftp.files
    print(f"[{'OK  ' if present else 'FAIL'}] {key} wurde hochgeladen")
    ok = ok and present

aktuell = json.loads(fake_sftp.files[json_key])
print("aktuell_json Gruppen im Einteilungs-Text (Stichprobe):",
      {k: v[0] if v else None for k, v in list(aktuell["einteilung"].items())[:2]})

state = json.loads(fake_sftp.files["state_auto.json"])
has_roles = bool(state.get("plan_data", {}).get(datum_kurz, {}).get("trainer_roles"))
print(f"[{'OK  ' if has_roles else 'FAIL'}] trainer_roles im State gespeichert: {state.get('plan_data', {}).get(datum_kurz, {}).get('trainer_roles')}")
ok = ok and has_roles

sys.exit(0 if ok else 1)
