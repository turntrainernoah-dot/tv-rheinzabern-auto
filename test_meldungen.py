#!/usr/bin/env python3
"""Tests fuer meldungen.py (Problem-Meldungen in den Admin-Bereich, seit 20.10.2026).
Laeuft ohne Server: SFTP wird durch ein Fake-Objekt ersetzt."""
import io, json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import meldungen as m

fails = 0
def check(name, cond, detail=""):
    global fails
    print(("[OK  ] " if cond else "[FAIL] ") + name + ("" if cond else f"  -- {detail}"))
    if not cond:
        fails += 1


class FakeFile(io.BytesIO):
    def __init__(self, store, key, mode):
        super().__init__(store.get(key, b"") if "r" in mode else b"")
        self.store, self.key, self.mode = store, key, mode
    def close(self):
        if "w" in self.mode:
            self.store[self.key] = self.getvalue()
        super().close()
    def __exit__(self, *a):
        self.close()


class FakeSFTP:
    def __init__(self, files=None, kaputt=False):
        self.files = files if files is not None else {}
        self.kaputt = kaputt
    def open(self, path, mode="r"):
        if self.kaputt:
            raise OSError("Verbindung weg")
        if "r" in mode and path not in self.files:
            raise FileNotFoundError(path)
        return FakeFile(self.files, path, mode)


# 1) Neue Meldung landet oben, mit allen Feldern
s = FakeSFTP()
ok = m.melde_problem("Trainingsplan", "Text A", "Details A", sftp=s)
liste = json.loads(s.files[m.REMOTE_DATEI])
check("melde_problem gibt True zurueck", ok is True)
check("eine Meldung gespeichert", len(liste) == 1, liste)
e = liste[0]
check("Felder vorhanden", all(k in e for k in ("id", "datum", "zuletzt", "anzahl", "bereich", "text", "details", "erledigt")), e)
check("erledigt=False, anzahl=1", e["erledigt"] is False and e["anzahl"] == 1, e)

# 2) Gleiche offene Meldung wird hochgezaehlt, nicht verdoppelt
m.melde_problem("Trainingsplan", "Text A", "Details neu", sftp=s)
liste = json.loads(s.files[m.REMOTE_DATEI])
check("gleiche offene Meldung -> keine Dopplung", len(liste) == 1, liste)
check("anzahl hochgezaehlt + Details aktualisiert", liste[0]["anzahl"] == 2 and liste[0]["details"] == "Details neu", liste[0])

# 3) Ist die alte erledigt, gibt es eine neue
liste[0]["erledigt"] = True
s.files[m.REMOTE_DATEI] = json.dumps(liste).encode()
m.melde_problem("Trainingsplan", "Text A", "", sftp=s)
liste = json.loads(s.files[m.REMOTE_DATEI])
check("nach 'Erledigt' wird gleiche Meldung neu angelegt", len(liste) == 2 and liste[0]["erledigt"] is False, liste)

# 4) Neueste oben, Begrenzung auf 100
import contextlib
with contextlib.redirect_stdout(io.StringIO()):
    for i in range(150):
        m.melde_problem("Test", f"Meldung {i}", sftp=s)
liste = json.loads(s.files[m.REMOTE_DATEI])
check("auf 100 Eintraege begrenzt", len(liste) == m.MAX_EINTRAEGE, len(liste))
check("neueste Meldung steht oben", liste[0]["text"] == "Meldung 149", liste[0])

# 5) Kaputte Datei / kaputte Verbindung bricht nie ab
s2 = FakeSFTP({m.REMOTE_DATEI: b"{kein json"})
check("kaputte Datei -> wird neu angelegt", m.melde_problem("X", "Y", sftp=s2) is True and len(json.loads(s2.files[m.REMOTE_DATEI])) == 1)
try:
    r = m.melde_problem("X", "Y", sftp=FakeSFTP(kaputt=True))
    check("kaputte SFTP-Verbindung -> False statt Exception", r is False)
except Exception as ex:
    check("kaputte SFTP-Verbindung -> False statt Exception", False, repr(ex))
os.environ["SSH_HOST"] = "127.0.0.1"; os.environ["SSH_PORT"] = "1"
try:
    r = m.melde_problem("X", "Y")
    check("kein Server erreichbar -> False statt Exception", r is False)
except BaseException as ex:
    check("kein Server erreichbar -> False statt Exception", False, repr(ex))

# 6) Details werden gekuerzt
s3 = FakeSFTP()
with contextlib.redirect_stdout(io.StringIO()):
    m.melde_problem("X", "Y", "x" * 10000, sftp=s3)
check("Details auf MAX_DETAILS gekuerzt", len(json.loads(s3.files[m.REMOTE_DATEI])[0]["details"]) == m.MAX_DETAILS)

print()
if fails:
    print(f"{fails} FEHLGESCHLAGEN")
    sys.exit(1)
print("Alle Tests bestanden.")
