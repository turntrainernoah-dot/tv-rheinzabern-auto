# tv-rheinzabern-auto

**Dieses Repo ist ÖFFENTLICH** (bewusst, wegen unbegrenzter GitHub-Actions-Minuten — siehe Struktur-Standard unten). Es betreibt die komplette Vereins-Automatik des TV Rheinzabern über GitHub Actions: Trainingsplan-Erstellung, Wochenchallenge-Auswertung, Video-Bau, Server-Backup und Heartbeat. Läuft automatisch alle paar Stunden, der PC muss dafür nicht an sein.

## ⚠️ Keine Klarnamen in diesem Repo

Weil das Repo öffentlich ist, dürfen hier **keine echten Namen von Kindern oder Trainern** stehen — weder in Code, Dateinamen, Kommentaren, Test-/Beispieldaten noch in Commit-Nachrichten. Die Zuordnung von Kürzel/ID zu Klarname liegt ausschließlich in `config/name_aliases.json` auf dem Server und wird zur Laufzeit geladen, nie eingecheckt. Wer hier versehentlich einen Namen einfügt, macht ihn öffentlich sichtbar — auch nach dem Löschen bleibt er in der Git-Historie.

## Ein Fehler hier legt den Vereinsbetrieb lahm

Die Workflows in diesem Repo sind keine Spielerei, sondern laufender Betrieb: Trainingspläne, Wochenchallenge-Status und Videos entstehen ausschließlich hier automatisiert. Ein kaputter Workflow oder ein fehlerhaftes Skript fällt nicht auf, bis der nächste automatische Lauf ausbleibt oder falsche Daten produziert — deshalb vor jeder Änderung besonders sorgfältig testen (z. B. `check_quick.py`) und nicht direkt auf `main` experimentieren, wenn Zweifel bestehen.

## Aufbau

- `auto_trainingsplan.py`, `auto_wochenchallenge.py`, `auto_wc_video.py`, `create_wochenchallenge_v2.py`, `wc_chat_parser.py`, `heartbeat.py`, `backup_state.py`, `build_step.py`, `check_quick.py` — die einzelnen Automatik-Skripte. `meldungen.py` — gemeinsame Problem-Meldung in den Admin-Bereich (siehe unten). `logschutz.py` — kürzt nach dem Roster-Laden Personennamen in stdout/stderr auf Initialen (Actions-Logs sind öffentlich); betrifft nur die Konsolenausgabe. Neue `print()`s mit Namen sind dadurch abgedeckt, Ausgaben **vor** `namen_setzen()` nicht.
- **Trainingsplan-Takt (seit 05.11.2026):** kein 30-Minuten-Cron mehr. Die Website startet `auto_trainingsplan.yml` per `repository_dispatch` (`trainingsplan`) beim Speichern; feste Zeiten nur noch als Sicherheitsnetz/Zeitfenster (siehe Kommentar im Workflow). `concurrency` verhindert Parallel-Läufe.
- `.github/workflows/`: `auto_trainingsplan.yml`, `auto_wochenchallenge.yml`, `auto_wc_video.yml`, `auto_backup.yml`, `heartbeat.yml` — je ein Workflow pro Automatik.
- `fonts/`, `state_backups/`: Hilfsdateien bzw. Zustandssicherungen der Automatik.
- `.github/workflows/keepalive.yml`: committet am 1. jedes Monats `keepalive.txt` (aktuelles Datum). Grund: GitHub deaktiviert Cron-Workflows in öffentlichen Repos nach 60 Tagen ohne Commit. **Nicht entfernen.**
- `Alt/`: abgelegte, nicht mehr aktive Workflows (z. B. `abi_auto_deploy.yml`, der frühere 5-Minuten-Deploy der privaten Abi-Lernapp). Nichts löschen; zurück nach `.github/workflows/` verschieben würde den Workflow wieder aktivieren.
- Zugangsdaten (SSH, Gmail) liegen ausschließlich als GitHub Secrets, nie im Code — siehe `README.md`.

## Benachrichtigungen (Stand 20.10.2026): "Nachrichten" im Admin-Bereich statt Mail

Seit 20.10.2026 wird **keine E-Mail und keine WhatsApp mehr versendet** (Übergabe an
Trainer, die nur Website + GitHub haben). Stattdessen schreiben die Skripte **nur bei
Fehlern/Problemen** über `meldungen.py::melde_problem(bereich, text, details, sftp=None)`
einen Eintrag per SFTP in `config/meldungen.json` auf dem Server. `admin.php` (Repo
`tv-trainerportal`) zeigt sie im Bereich **„Nachrichten“** an (roter Zähler im Menü,
Knöpfe „Erledigt“/„Alle erledigt“, sonst „Alles läuft“).

- Erfolgs-/Info-Meldungen ("Plan erstellt", "Video hochgeladen", "WC fertig", "Alles läuft")
  gibt es **nicht mehr** – keine Meldung = alles in Ordnung.
- `melde_problem()` wirft nie eine Exception (ein Melde-Fehler bricht nie den Lauf ab),
  hält die Datei auf max. 100 Einträge und zählt eine gleiche offene Meldung nur hoch
  (`anzahl`), statt sie bei jedem 30-Minuten-Lauf neu anzulegen.
- `text` = kurzer Satz für Trainer ohne Technikwissen (was ist los, was tun);
  `details` = technischer Fehler/Traceback zum Reinkopieren in Claude.
- Keine Klarnamen in Meldungstexten im Code (öffentliches Repo). Laufzeitdaten in
  `details` landen nur auf dem Server (`config/` ist gesperrt und wird nie deployt).
- Das **Empfangen** der Wochenchallenge-Mails per IMAP (`GMAIL_USER`/`GMAIL_APP_PASSWORD`)
  ist unverändert. Nicht mehr genutzte Secrets: `EMAIL_TO`, `WHATSAPP_PHONE`, `CALLMEBOT_APIKEY`.
- `heartbeat.py` braucht dafür jetzt die `SSH_*`-Secrets (im Workflow eingetragen).

---

## Struktur-Standard (Kurzfassung, Stand 04.11.2026)

Gilt für alle Repos und den PC-Ordner `C:\Claude` und steht wortgleich in jeder
CLAUDE.md, damit auch eine Session, die nur ein Repo sieht, ihn kennt. **Nicht
wegoptimieren.** Volltext mit Begründungen und Historie: `tv-memory` →
`00 Kontext/Struktur-Standard.md` (nur bei Bedarf lesen).

**Wohin gehört was**
- Code einer Website → ihr eigenes Repo. Wissen zu einer Website → `CLAUDE.md`/`docs/` dieses Repos.
- Übergreifendes Vereinswissen → `tv-memory`; Privates (Familie, Studium, Schreibstil) →
  `noah-memory`, nie in ein Vereins-Repo. Beide sind selbst Noahs Obsidian-Vaults:
  Notizen direkt dort bearbeiten, committen, pushen (sonst sehen PC und iPad verschiedene Stände).
- Datenbanken, Uploads, echte Nutzerdaten → nur auf den Server, in **kein** Repo (auch kein privates).
- Zugangsdaten, Schlüssel, Tokens → Passwort-Manager bzw. GitHub Actions Secret – nie in Code, Notizen, Commits.
- PC: Wegwerf-Dateien von Anfang an in `C:\Claude\_work\`, Werkzeuge in `Aktiv\`,
  Abgelegtes in `Alt\` – nie lose ins Stammverzeichnis.

**Repos** (alle privat außer `tv-rheinzabern-auto`): `tv-trainerportal` (Hauptseite und
Trainerportal, `/`), `tv-spiele` (`/spiele/`), `tv-wogele2026` (`/wogele2026/`),
`tv-belege` (Abrechnungsportal auf Azure), `abi-lernen` (`/abi/`), `familie-website` und
`dashboard-privat` (privat, keine Verbindung zu Vereins-Repos), `tv-memory`, `noah-memory`,
`tv-backups`, `tv-rheinzabern-auto` (Vereins-Automatik).

**Niemals**
- Echte Daten oder Schlüssel committen – was einmal in der Historie liegt, bleibt dort.
- Ein Repo öffentlich stellen, um Kosten oder Kontingent zu umgehen. `tv-rheinzabern-auto`
  ist absichtlich öffentlich (30-Minuten-Cron, unbegrenzte Actions-Minuten) und darf **nie
  privat** werden, sonst steht die Vereins-Automatik (Ausfall 19.08.2026). Dort keine
  Klarnamen von Kindern oder Trainern (Zuordnung liegt auf dem Server).
- Per Deploy `config.php`, `data/`, `uploads/` oder von der Automatik gepflegte Dateien
  überschreiben (Deploys nutzen eine Erlaubnisliste). Braucht Code eine neue Konstante in
  `config.php`, muss sie **vor** dem Deploy von Hand auf den Server.
- Backups über GitHub Actions: Backups laufen nur manuell vom PC („Server-Backup
  ziehen.bat“), keine Daten bei GitHub, auch nicht verschlüsselt. `git revert` hilft nur
  für Code, nie für Daten.

**Arbeitsweise**
- Jede Erkenntnis an genau einer Stelle, sonst verweisen (`[[Wikilinks]]`).
- Nach jeder Aufgabe selbst prüfen, was sich geändert hat, und es an der richtigen Stelle nachtragen.
- Nichts löschen, sondern nach `Alt\` verschieben; vorher prüfen, ob feste Pfade darauf zeigen.

**Deploy** (Code: Repo → Deploy → Server; Daten bleiben auf dem Server). Wie dieses Repo
deployt wird, steht oben in dieser Datei.
- `tv-spiele`, `tv-wogele2026`, `tv-trainerportal`: Workflow „Website deployen“
  (`deploy_website.yml`) in `tv-rheinzabern-auto`, Eingaben `website`, `dry_run`, `trotzdem`.
  Erst Trockenlauf, dann echt. **Auf das Ergebnis warten** (ein Start sagt nichts), im Log
  `Fertig: N Datei(en)` bzw. `ABBRUCH` prüfen, danach
  `https://tv-rheinzabern.e-websolutions.de/<pfad>/` abrufen (muss 200 liefern). Abbruch mit
  „ungemergte Arbeit“: fertiger Code liegt auf einem anderen Branch – erst mergen. Ohne `gh`
  (Cloud-Session): GitHub-MCP `actions_run_trigger` mit `workflow_id: deploy_website.yml`
  (Dateiname, nicht Anzeigename). Ohne Zugriff: Noah den Knopf drücken lassen, nie „erledigt“ melden.
- `familie-website`, `abi-lernen`: eigener Workflow im Repo. `tv-belege`: Push auf `main` → Azure.
