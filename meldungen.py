# -*- coding: utf-8 -*-
"""
meldungen.py - Problem-Meldungen fuer den Admin-Bereich (seit 20.10.2026)
=========================================================================
Ersetzt den frueheren Mail-/WhatsApp-Versand der Automatik. Die Skripte
schreiben NUR noch bei Fehlern/Problemen einen Eintrag per SFTP in
config/meldungen.json auf dem Server. admin.php zeigt diese Datei im Bereich
"Nachrichten" an (Knoepfe "Erledigt" / "Alle erledigt").

Regeln:
  - melde_problem() wirft NIE eine Exception - ein Fehler beim Melden darf den
    eigentlichen Lauf nie abbrechen.
  - Die Datei wird auf die letzten MAX_EINTRAEGE Eintraege begrenzt.
  - Gleiche offene Meldung (gleicher Bereich + gleicher Text) wird nicht
    doppelt angelegt, sondern nur hochgezaehlt (anzahl/zuletzt) - sonst wuerde
    z.B. der 30-Minuten-Trainingsplan-Lauf bei einem Dauerfehler die Liste
    fluten.
  - Keine Klarnamen im Code (oeffentliches Repo). Was zur Laufzeit in den
    Fehlerdetails steht, landet nur auf dem Server (config/ ist per .htaccess
    gesperrt und wird nie deployt).

Format eines Eintrags (Liste, neueste zuerst):
  {"id": "...", "datum": "2026-10-20T08:15:00+02:00", "zuletzt": "...",
   "anzahl": 1, "bereich": "Trainingsplan", "text": "...", "details": "...",
   "erledigt": false}
"""
import json
import os
import uuid
from datetime import datetime, timezone

REMOTE_DATEI = "config/meldungen.json"
MAX_EINTRAEGE = 100
MAX_DETAILS = 3000

def _jetzt():
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("Europe/Berlin")).isoformat(timespec="seconds")
    except Exception:
        return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _sftp_oeffnen():
    import paramiko
    t = paramiko.Transport((os.environ.get("SSH_HOST", ""),
                            int(os.environ.get("SSH_PORT", "22") or 22)))
    t.connect(username=os.environ.get("SSH_USER", ""),
              password=os.environ.get("SSH_PASSWORD", ""))
    return paramiko.SFTPClient.from_transport(t), t


def eintrag_einfuegen(liste, bereich, text, details=""):
    """Reine Listen-Logik (ohne SFTP, damit testbar). Gibt die neue Liste zurueck."""
    if not isinstance(liste, list):
        liste = []
    jetzt = _jetzt()
    details = (details or "").strip()[-MAX_DETAILS:]
    for e in liste:
        if (isinstance(e, dict) and not e.get("erledigt")
                and e.get("bereich") == bereich and e.get("text") == text):
            e["anzahl"] = int(e.get("anzahl", 1) or 1) + 1
            e["zuletzt"] = jetzt
            if details:
                e["details"] = details
            liste.remove(e)
            liste.insert(0, e)
            return liste[:MAX_EINTRAEGE]
    liste.insert(0, {
        "id": uuid.uuid4().hex[:12],
        "datum": jetzt,
        "zuletzt": jetzt,
        "anzahl": 1,
        "bereich": bereich,
        "text": text,
        "details": details,
        "erledigt": False,
    })
    return liste[:MAX_EINTRAEGE]


def melde_problem(bereich, text, details="", sftp=None):
    """Schreibt eine Problem-Meldung nach config/meldungen.json.
    bereich: z.B. "Trainingsplan", "Wochenchallenge", "Wochenchallenge-Video",
             "System-Check".
    text:    kurzer, fuer Trainer verstaendlicher Satz (was ist los, was tun).
    details: technische Fehlerdetails (fuer Claude zum Reinkopieren).
    sftp:    vorhandene paramiko-SFTP-Verbindung (optional) - sonst wird eine
             eigene aufgebaut.
    Wirft nie eine Exception, gibt True/False zurueck."""
    print(f"[MELDUNG] {bereich}: {text}")
    if details:
        print(f"[MELDUNG] Details: {str(details)[:500]}")
    eigene = None
    try:
        if sftp is None:
            sftp, eigene = _sftp_oeffnen()
        try:
            with sftp.open(REMOTE_DATEI, "r") as f:
                liste = json.loads(f.read().decode("utf-8", errors="replace") or "[]")
        except Exception:
            liste = []
        liste = eintrag_einfuegen(liste, bereich, text, str(details or ""))
        daten = json.dumps(liste, indent=1, ensure_ascii=False).encode("utf-8")
        with sftp.open(REMOTE_DATEI, "w") as f:
            f.write(daten)
        print("[MELDUNG] in config/meldungen.json gespeichert.")
        return True
    except Exception as e:
        print(f"[MELDUNG] konnte nicht gespeichert werden: {e!r}")
        return False
    finally:
        if eigene is not None:
            try:
                sftp.close()
            except Exception:
                pass
            try:
                eigene.close()
            except Exception:
                pass
