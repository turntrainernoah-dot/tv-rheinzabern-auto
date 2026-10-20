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
| GMAIL_USER | Gmail-Adresse, an die die Wochenchallenge-Mails gehen |
| GMAIL_APP_PASSWORD | Gmail App-Passwort |

`GMAIL_USER`/`GMAIL_APP_PASSWORD` werden seit 20.10.2026 **nur noch zum Empfangen** der
Wochenchallenge-Mails (IMAP) gebraucht. Versendet wird nichts mehr: Probleme stehen im
Admin-Bereich der Website unter **„Nachrichten“** (Datei `config/meldungen.json` auf dem
Server, geschrieben von `meldungen.py`). `EMAIL_TO`, `WHATSAPP_PHONE` und
`CALLMEBOT_APIKEY` werden nicht mehr verwendet.
