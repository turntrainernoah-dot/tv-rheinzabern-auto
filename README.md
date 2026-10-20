# TV Rheinzabern – Auto Trainingsplan

Läuft via GitHub Actions alle 4 Stunden. PC muss nicht an sein.

## Secrets (einmalig in GitHub eintragen)

| Secret | Wert |
|--------|------|
| SSH_HOST | access-5017462830.webspace-host.com |
| SSH_USER | a2358459 |
| SSH_PASSWORD | (SFTP-Passwort) |
| SSH_PORT | 22 |
| GMAIL_USER | noahwoe0212@gmail.com |
| GMAIL_APP_PASSWORD | (Gmail App-Passwort) |

`GMAIL_USER`/`GMAIL_APP_PASSWORD` werden seit 20.10.2026 **nur noch zum Empfangen** der
Wochenchallenge-Mails (IMAP) gebraucht. Versendet wird nichts mehr: Probleme stehen im
Admin-Bereich der Website unter **„Nachrichten“** (Datei `config/meldungen.json` auf dem
Server, geschrieben von `meldungen.py`). `EMAIL_TO`, `WHATSAPP_PHONE` und
`CALLMEBOT_APIKEY` werden nicht mehr verwendet.
