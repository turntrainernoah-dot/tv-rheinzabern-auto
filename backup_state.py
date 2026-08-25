# -*- coding: utf-8 -*-
# Taegliches State-Backup.
#
# Phase 2 (24.08.2026): Vorher lud dieses Skript die Zustands-/Konfig-Dateien
# vom Webspace und legte sie ins OEFFENTLICHE Repo (state_backups/<datum>/) --
# taeglich per Commit+Push (siehe auto_backup.yml). Das haeufte echte
# Kindernamen, Anwesenheiten, Trainer-Einteilungen und Abmeldungen im
# oeffentlichen Git-Verlauf an (siehe Vault: GitHub-Umzug, Phase-2-Inventur).
#
# Neu: die Dateien werden per SFTP (get) geholt und ANSCHLIESSEND per SFTP
# (put) direkt wieder auf den Server zurueckgeschrieben -- in den bereits per
# .htaccess geschuetzten Ordner config/_state_backups/<datum>/ (dieselbe
# Deny-from-all-Regel wie config/config.json, da .htaccess-Regeln in Apache
# rekursiv fuer Unterordner gelten, siehe Server-Verifikation in
# [[GitHub-Umzug]]). Es landet dadurch NICHTS mehr in git. Funktion
# unveraendert: taegliche Punkt-in-Zeit-Snapshots zur Drift-Diagnose bleiben
# erhalten, nur der Ablageort wechselt (Server statt oeffentliches Repo) --
# exakt das SFTP-Get-dann-Put-Muster, das auto_wc_video.py/backup_state.py
# an anderer Stelle schon nutzen.
import os, io, paramiko, datetime, posixpath

HOST=os.environ["SSH_HOST"]; USER=os.environ["SSH_USER"]; PW=os.environ["SSH_PASSWORD"]; PORT=int(os.environ.get("SSH_PORT","22"))
FILES=["state_auto.json","wc_state_auto.json","fixed_entries.json","config/config.json",
       "abmeldungen/abmeldungen.json","abmeldungen/trainingsentfall.json","abmeldungen/tokens.json",
       "anmerkungen/anmerkungen.json","trainingsnotizen/notizen.json"]

day = datetime.datetime.utcnow().strftime("%Y-%m-%d")
remote_dir = posixpath.join("config", "_state_backups", day)

t = paramiko.Transport((HOST, PORT)); t.connect(username=USER, password=PW)
sftp = paramiko.SFTPClient.from_transport(t)

# Zielordner anlegen (rekursiv, da mkdir bei bereits vorhandenem Ordner nur einen Fehler wirft)
parts = remote_dir.split("/")
cur = ""
for p in parts:
    cur = posixpath.join(cur, p) if cur else p
    try:
        sftp.mkdir(cur)
    except IOError:
        pass  # existiert schon

ok = 0
for rp in FILES:
    try:
        buf = io.BytesIO()
        sftp.getfo(rp, buf)
        remote_target = posixpath.join(remote_dir, rp.replace("/", "__"))
        buf.seek(0)
        sftp.putfo(buf, remote_target)
        print("OK", rp, "->", remote_target)
        ok += 1
    except Exception as e:
        print("SKIP", rp, e)

sftp.close(); t.close()
print(f"[FERTIG] {day}: {ok}/{len(FILES)} gesichert nach {remote_dir}/ (Server, nicht mehr in git)")
