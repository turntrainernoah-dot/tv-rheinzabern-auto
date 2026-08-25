#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Einmalige Migration von config/config.json auf das neue Gruppenzeiten-Format
(Gruppenzeiten-Umbau, 25.08.2026): "gruppen" wird von einem String-Array zu
einem Array von {"name","zeiten"} umgestellt (mit den migrierten Default-
Werten aus auto_trainingsplan.py::_default_zeiten_for() -- siehe dort fuer
die Begruendung, warum G3 bewusst dieselbe versetzte Zeit wie G4 bekommt).
"personen"-Eintraege mit rolle="trainer" bekommen "immer_springer": false.

Aendert das Live-Verhalten NICHT, solange Noah im Admin-Bereich nichts
eintraegt (Ausnahme: der gedruckte Plan wird um eine redundante Zeitraster-
Zeile kuerzer -- siehe compute_time_grid()-Docstring in auto_trainingsplan.py,
harmlos).

Nutzung:
  python3 migrate_gruppenzeiten.py                    # echte SFTP-Migration
                                                       # (SSH_HOST/SSH_USER/
                                                       # SSH_PASSWORD/SSH_PORT
                                                       # als Env-Vars noetig)
  python3 migrate_gruppenzeiten.py --dry-run          # nur anzeigen, nichts schreiben
  python3 migrate_gruppenzeiten.py --local IN.json OUT.json   # lokale Dateien
                                                       # statt SFTP (fuer Tests,
                                                       # OUT wird nur bei
                                                       # fehlendem --dry-run geschrieben)

Ist "gruppen" schon im neuen Objekt-Format, macht das Skript nichts (sicher
mehrfach ausfuehrbar). Legt vor dem Schreiben zusaetzlich
config/config_vor_gruppenzeiten_migration.json auf dem Server an (zusaetzlich
zu Noahs bereits gezogenem vollstaendigem Server-Backup vom 24.08.2026).
"""
import argparse, json, os, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from auto_trainingsplan import _default_zeiten_for, get_sftp


def migrate(cfg: dict) -> dict:
    """Mutiert cfg in-place und gibt es zurueck."""
    gruppen_raw = cfg.get("gruppen") or []
    if gruppen_raw and isinstance(gruppen_raw[0], dict):
        print("[MIGRATE] gruppen ist bereits im neuen Format -- nichts zu tun bei den Gruppen.")
    else:
        new_gruppen = [{"name": g, "zeiten": _default_zeiten_for(g)} for g in gruppen_raw]
        cfg["gruppen"] = new_gruppen
        print(f"[MIGRATE] {len(new_gruppen)} Gruppen migriert: {[g['name'] for g in new_gruppen]}")

    n_trainer = 0
    for p in cfg.get("personen", []):
        if p.get("rolle") == "trainer" and "immer_springer" not in p:
            p["immer_springer"] = False
            n_trainer += 1
    print(f"[MIGRATE] {n_trainer} Trainer bekommen immer_springer=false.")
    return cfg


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--dry-run", action="store_true", help="nur anzeigen, nichts schreiben")
    ap.add_argument("--local", nargs=2, metavar=("IN", "OUT"), help="lokale Dateien statt SFTP (fuer Tests)")
    args = ap.parse_args()

    if args.local:
        in_path, out_path = args.local
        with open(in_path, encoding="utf-8") as f:
            cfg = json.load(f)
        migrated = migrate(cfg)
        if args.dry_run:
            print(json.dumps(migrated, indent=2, ensure_ascii=False))
            return
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(migrated, f, indent=2, ensure_ascii=False)
        print(f"[OK] Geschrieben: {out_path}")
        return

    print("Verbinde mit Server...")
    ssh, sftp = get_sftp()
    try:
        f = sftp.open("config/config.json", "r")
        cfg = json.loads(f.read().decode("utf-8"))
        f.close()
    except Exception as e:
        print(f"[FEHLER] config/config.json konnte nicht gelesen werden: {e!r}")
        sftp.close(); ssh.close()
        sys.exit(1)

    gruppen_raw = cfg.get("gruppen") or []
    if gruppen_raw and isinstance(gruppen_raw[0], dict):
        print("[MIGRATE] config.json ist bereits im neuen Format. Nichts zu tun.")
        sftp.close(); ssh.close()
        return

    original_json = json.dumps(cfg, indent=2, ensure_ascii=False)
    migrated = migrate(json.loads(original_json))  # migrate() auf einer tiefen Kopie

    if args.dry_run:
        print("\n--- DRY RUN: wuerde folgendes nach config/config.json schreiben ---")
        print(json.dumps(migrated, indent=2, ensure_ascii=False))
        sftp.close(); ssh.close()
        return

    backup_name = "config/config_vor_gruppenzeiten_migration.json"
    f = sftp.open(backup_name, "wb")
    f.write(original_json.encode("utf-8"))
    f.close()
    print(f"[OK] Zusaetzliches Server-Backup geschrieben: {backup_name}")

    f = sftp.open("config/config.json", "wb")
    f.write(json.dumps(migrated, indent=2, ensure_ascii=False).encode("utf-8"))
    f.close()
    print("[OK] config/config.json migriert und gespeichert.")

    sftp.close(); ssh.close()


if __name__ == "__main__":
    main()
