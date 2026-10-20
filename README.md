# TV Rheinzabern – Auto Trainingsplan

Läuft via GitHub Actions (Trainingsplan, Wochenchallenge, Videos, Backup, Heartbeat). PC muss nicht an sein.
Details und Regeln: `CLAUDE.md`.

## Secrets (einmalig in GitHub eintragen: Settings → Secrets and variables → Actions)

Nur die Namen stehen hier; die Werte liegen im Passwort-Manager.

| Secret | Wert |
|--------|------|
| SSH_HOST | SFTP-Host des Webspace |
| SSH_USER | SFTP-Benutzer |
| SSH_PASSWORD | SFTP-Passwort |
| SSH_PORT | Port (22) |
| GMAIL_USER | Gmail-Adresse, von der Benachrichtigungen gesendet werden |
| GMAIL_APP_PASSWORD | Gmail App-Passwort |
| EMAIL_TO | Empfänger der Benachrichtigungen |
