# -*- coding: utf-8 -*-
# Laedt taeglich die Zustands-/Konfig-Dateien vom Webspace und legt sie ins Repo (state_backups/<datum>/).
import os, paramiko, datetime
HOST=os.environ["SSH_HOST"]; USER=os.environ["SSH_USER"]; PW=os.environ["SSH_PASSWORD"]; PORT=int(os.environ.get("SSH_PORT","22"))
FILES=["state_auto.json","wc_state_auto.json","fixed_entries.json","config/config.json",
       "abmeldungen/abmeldungen.json","abmeldungen/trainingsentfall.json","abmeldungen/tokens.json",
       "anmerkungen/anmerkungen.json","trainingsnotizen/notizen.json"]
day=datetime.datetime.utcnow().strftime("%Y-%m-%d")
out=os.path.join("state_backups", day); os.makedirs(out, exist_ok=True)
t=paramiko.Transport((HOST,PORT)); t.connect(username=USER,password=PW); sftp=paramiko.SFTPClient.from_transport(t)
ok=0
for rp in FILES:
    try:
        sftp.get(rp, os.path.join(out, rp.replace("/","__"))); print("OK", rp); ok+=1
    except Exception as e:
        print("SKIP", rp, e)
sftp.close(); t.close()
print(f"[FERTIG] {day}: {ok}/{len(FILES)} gesichert")
