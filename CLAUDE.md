# tv-rheinzabern-auto

**Dieses Repo ist ÖFFENTLICH** (bewusst, wegen unbegrenzter GitHub-Actions-Minuten — siehe Struktur-Standard unten, Abschnitt 2). Es betreibt die komplette Vereins-Automatik des TV Rheinzabern über GitHub Actions: Trainingsplan-Erstellung, Wochenchallenge-Auswertung, Video-Bau, Server-Backup und Heartbeat. Läuft automatisch alle paar Stunden, der PC muss dafür nicht an sein.

## ⚠️ Keine Klarnamen in diesem Repo

Weil das Repo öffentlich ist, dürfen hier **keine echten Namen von Kindern oder Trainern** stehen — weder in Code, Dateinamen, Kommentaren, Test-/Beispieldaten noch in Commit-Nachrichten. Die Zuordnung von Kürzel/ID zu Klarname liegt ausschließlich in `config/name_aliases.json` auf dem Server und wird zur Laufzeit geladen, nie eingecheckt. Wer hier versehentlich einen Namen einfügt, macht ihn öffentlich sichtbar — auch nach dem Löschen bleibt er in der Git-Historie.

## Ein Fehler hier legt den Vereinsbetrieb lahm

Die Workflows in diesem Repo sind keine Spielerei, sondern laufender Betrieb: Trainingspläne, Wochenchallenge-Status und Videos entstehen ausschließlich hier automatisiert. Ein kaputter Workflow oder ein fehlerhaftes Skript fällt nicht auf, bis der nächste automatische Lauf ausbleibt oder falsche Daten produziert — deshalb vor jeder Änderung besonders sorgfältig testen (z. B. `check_quick.py`) und nicht direkt auf `main` experimentieren, wenn Zweifel bestehen.

## Aufbau

- `auto_trainingsplan.py`, `auto_wochenchallenge.py`, `auto_wc_video.py`, `create_wochenchallenge_v2.py`, `wc_chat_parser.py`, `heartbeat.py`, `backup_state.py`, `build_step.py`, `check_quick.py` — die einzelnen Automatik-Skripte.
- `.github/workflows/`: `auto_trainingsplan.yml`, `auto_wochenchallenge.yml`, `auto_wc_video.yml`, `auto_backup.yml`, `heartbeat.yml` — je ein Workflow pro Automatik.
- `fonts/`, `state_backups/`: Hilfsdateien bzw. Zustandssicherungen der Automatik.
- Zugangsdaten (SSH, Gmail) liegen ausschließlich als GitHub Secrets, nie im Code — siehe `README.md`.


---

## Struktur-Standard (Stand 25.08.2026)

Dieser Abschnitt gilt für **alle** Repos und für den lokalen Ordner `C:\Claude`. Er wird wortgleich in jede CLAUDE.md eingefügt, damit jede Claude-Session ihn kennt — auch eine Cloud-Session, die nur ein einziges Repo sieht.

**Diese Regeln nicht wegoptimieren.** Sie stehen bewusst redundant in jedem Repo, weil eine Session, die nur ein Repo sieht, sonst nichts von ihnen wüsste. Am 11.07.2026 wurde `C:\Claude` schon einmal von über 4000 losen Dateien aufgeräumt; sechs Wochen später lagen wieder 83 lose Dateien dort. Ursache war nie fehlende Ordnung, sondern dass die Regel nicht an der Quelle griff.

**Seit 25.08.2026 sind `tv-memory` und `noah-memory` selbst Noahs Obsidian-Vaults** — es gibt keine separate lokale Vault-Kopie mehr (das frühere `C:\Claude\Obsidian Vault` ist archiviert, siehe `C:\Claude\Archiv\`). Wer am PC oder iPad Notizen bearbeitet, tut das direkt in einem dieser beiden Repos und committet/pusht die Änderung — sonst sehen PC und iPad unterschiedliche Stände.

### 1. Wo gehört was hin

| Was | Wohin | Niemals |
|---|---|---|
| Quellcode einer Website | in ihr eigenes Repo | in ein fremdes Repo |
| Datenbanken, Uploads, echte Nutzerdaten | nur auf den Server | in **kein** Repo, auch kein privates |
| Zugangsdaten, Schlüssel, Tokens | Passwort-Manager bzw. GitHub Actions Secret | nicht in Code, nicht in Notizen, nicht in Commits |
| Wissen zu **einer** Website | `CLAUDE.md` / `docs/` dieses Repos | nicht ins Memory-Repo |
| Übergreifendes Vereinswissen, Obsidian-Notizen dazu | `tv-memory` (ist selbst der Vault) | nicht in ein Website-Repo, nicht in einem lokalen Vault-Ordner |
| Privates (Familie, Studium, Schreibstil), Obsidian-Notizen dazu | `noah-memory` (ist selbst der Vault) | niemals in ein Vereins-Repo, nicht in einem lokalen Vault-Ordner |
| Wegwerf-Dateien einer Session | `C:\Claude\_work\` | nie lose im Stammverzeichnis |
| Dauerhaft genutzte Werkzeuge | `C:\Claude\Aktiv\` | nicht in `_work\` |
| Abgelegtes, das aufbewahrt wird | `C:\Claude\Archiv\` | nicht löschen |

(Die letzten drei Zeilen sind identisch mit der Wegwerf-Dateien-Regel oben — hier nur der Vollständigkeit halber mit aufgeführt, weil dieselbe Tabelle auch in jedem Repo steht.)

### 2. Die neun Repos

| Repo | Sichtbarkeit | Inhalt |
|---|---|---|
| `tv-trainerportal` | privat | Vereins-Hauptseite, Trainer-Portal, Admin |
| `tv-spiele` | privat | Turnierplan und Rangliste |
| `tv-belege` | privat | Belege- und Stundennachweis-Portal |
| `tv-wogele2026` | privat | Wogele-Seite |
| `familie-website` | privat | **Privat.** Keine Verbindung zu Vereins-Repos |
| `tv-memory` | privat, weitergebbar | Vereinswissen ohne Zugangsdaten — **ist selbst Noahs Vereins-Obsidian-Vault** |
| `noah-memory` | privat | Persönliches, Studium, Schreibstil — **ist selbst Noahs privater Obsidian-Vault** |
| `tv-backups` | privat | Server-Backup-Automatik |
| `tv-rheinzabern-auto` | **öffentlich** | Automatisierung. Öffentlich wegen unbegrenzter Actions-Minuten |

**Zu `tv-rheinzabern-auto`:** Das Repo ist absichtlich öffentlich, weil die Vereins-Automatik sonst am Minuten-Kontingent scheitert (Ausfall vom 19.08.2026). Deshalb dürfen dort **keine Klarnamen von Kindern oder Trainern** stehen — die Zuordnung liegt in `config/name_aliases.json` auf dem Server und wird zur Laufzeit geladen. Wer dort Namen einfügt, macht sie öffentlich.

### 3. Was niemals passieren darf

- **Keine echten Daten in ein Repo.** In der Belege-Datenbank stehen IBANs, in der Familienwebsite private Familiendaten. Auch ein privates Repo ist eine Kopie an einem fremden Ort.
- **Keinen Schlüssel committen.** Was einmal in der Historie liegt, bleibt dort, auch nach dem Löschen der Datei.
- **Kein Repo öffentlich stellen**, um ein Kontingent- oder Kostenproblem zu umgehen. Die einzige Ausnahme ist `tv-rheinzabern-auto`, und die ist oben begründet.
- **`tv-rheinzabern-auto` darf niemals auf privat gestellt werden.** Das Repo läuft rund um die Uhr mit einem 30-Minuten-Cron. Als öffentliches Repo kostet das keine Actions-Minuten. Würde es privat gestellt, wäre das Monatskontingent von 2000 Minuten in etwa zwei Wochen aufgebraucht und die komplette Vereins-Automatik stünde still — genau das ist am 19.08.2026 passiert. Wer die Sichtbarkeit dieses Repos ändern will, muss vorher den Cron-Takt drastisch senken, sonst legt er den Verein lahm.
- **Kein Deploy überschreibt `config.php`, `data/` oder `uploads/`.** Die Workflows nutzen dafür eine Erlaubnisliste statt einer Ausschlussliste.
- **Neue Konstanten in `config.php` brauchen zuerst eine manuelle Ergänzung auf dem Server.** `config.php` wird nie deployt (s.o.) — der Deploy kann eine neue Konstante deshalb nicht automatisch nachziehen. Setzt eine Code-Änderung eine neue Konstante voraus, muss die Server-`config.php` sie **vor** dem Deploy von Hand bekommen, sonst bricht die Seite direkt danach. Beispiel: `PUBLIC_DIR` fehlte am 25.08.2026 beim ersten Live-Deploy von `tv-belege` aus genau diesem Grund (behoben, dokumentiert in `tv-belege/docs/DEPLOY.md`).
- **Beim Trainerportal zusätzlich:** Dateien, die die Automatik pflegt (Trainingspläne, Wochenchallenge-Status, `config/`), dürfen von keinem Deploy angefasst werden.

### 4. Arbeitsweise

- **Jede Erkenntnis an genau einer Stelle.** Nicht mehrfach ablegen — stattdessen mit `[[Wikilinks]]` verweisen. Ausnahme ist dieser Struktur-Standard selbst, siehe Begründung oben.
- **Wegwerf-Dateien von Anfang an in `_work\` anlegen**, nicht erst im Stammverzeichnis erzeugen und später aufräumen. Nach Abschluss einer Aufgabe dort aufräumen.
- **Nach jeder abgeschlossenen Aufgabe** eigenständig prüfen, was sich dadurch geändert hat, und es an der inhaltlich richtigen Stelle ergänzen — ohne dass jemand danach fragen muss.
- **Vor dem Verschieben von Dateien** prüfen, ob fest eingetragene Pfade darauf zeigen. Beim Aufräumen im Juli 2026 zeigten danach `.bat`-Dateien ins Leere.
- **Nichts löschen.** Was wegsoll, wird nach `Archiv\` verschoben.

### 5. Deploy und Datenfluss

```
Code:    iPad / PC  →  Repo  →  Deploy  →  Server
Daten:                          bleiben ausschliesslich auf dem Server
Backup:  Server  →  PC   (manuell, "Server-Backup ziehen.bat" auf dem Desktop)
```

**Deploy** laeuft normalerweise ueber GitHub Actions bei einem Push auf `main`. Solange
Actions in privaten Repos gesperrt sind, geht es mit `Aktiv\deploy_belege.py` auch lokal
per SFTP. Beide Wege nutzen dieselbe Erlaubnisliste und fassen `config.php`, `data/` und
`uploads/` niemals an.

**Backups laufen ausschliesslich manuell.** Ein automatisches Backup ueber GitHub Actions
war vorbereitet und wurde am 25.08.2026 auf Noahs ausdruecklichen Wunsch wieder entfernt:
Die Belege-Datenbank enthaelt IBANs, die Familienwebsite private Daten — diese sollen
grundsaetzlich nicht bei GitHub liegen, auch nicht verschluesselt. Nicht ohne Ruecksprache
wieder einbauen.

Zurueckrollen per `git revert` gilt **nur fuer Code**. Fuer Daten gibt es keinen
Repo-Rueckweg — dafuer zaehlt allein das manuell gezogene Server-Backup.

### 6. Stand und offene Punkte

- GitHub Actions sind bis zum **1. September 2026** blockiert (Monatskontingent aufgebraucht, kein Zahlungsproblem). Bis dahin kein Deploy über Actions — Deploys laufen in der Zwischenzeit lokal per Skript (SFTP, Python + `paramiko`).
- **Verbrauchsrechnung** (grob, damit künftige Sessions einschätzen können, was noch ins Kontingent passt): `tv-rheinzabern-auto` verbraucht 0 Minuten, weil öffentlich. Backups laufen seit 25.08.2026 nicht mehr über Actions (siehe unten), verbrauchen also ebenfalls 0 Minuten. Die fünf Website-Deploys laufen nur bei einem Push, grob 1–2 Minuten pro Deploy. Zusammen deutlich unter 100 von 2000 Minuten — unkritisch, sobald der Block am 1. September endet. Der Kern des Problems war nie die Zahl der privaten Repos, sondern ausschließlich ein zeitgesteuerter Dauerlauf (Cron) in einem privaten Repo.
- **Backups laufen seit 25.08.2026 ausschließlich manuell vom PC**, per Desktop-Datei "Server-Backup ziehen.bat" — bewusst nicht mehr automatisiert über GitHub Actions. Entscheidung von Noah: es sollen keine Daten bei GitHub liegen, auch nicht verschlüsselt. Die alte Actions-Backup-Automatik wurde entfernt, nicht nur pausiert. Details/Anleitung: `BACKUP-HINWEIS.md` im `tv-backups`-Repo.
- **`tv-belege` läuft seit 25.08.2026 live auf der neuen Struktur** — erster echter Live-Deploy erfolgreich: `public/` wird ausgeliefert, `lib/`/`data/`/`uploads/` sind per `.htaccess` gesperrt (403 verifiziert), eine `.htaccess` im Wurzelverzeichnis leitet intern nach `public/` um. Nach dem Deploy verifiziert: Startseite/Login/Registrierung je HTTP 200, Datenbank SHA-256-identisch zum Stand davor, alle 16 Belegdateien unverändert. Der Deploy lief **nicht** über GitHub Actions (s.o., bis 1. September gesperrt), sondern lokal per SFTP-Skript — das bleibt bis dahin der Weg für jeden weiteren Deploy. Zwei dabei gefundene Fehler (Server-`config.php` kannte die neue Konstante `PUBLIC_DIR` nicht, siehe Regel in Abschnitt 3; eine Rewrite-Regel lief in eine Endlosschleife) sind behoben. Ausführlich dokumentiert in `docs/DEPLOY.md` im `tv-belege`-Repo.
