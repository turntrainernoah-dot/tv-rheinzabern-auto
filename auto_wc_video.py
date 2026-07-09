#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
auto_wc_video.py - Cloud-Generierung der Wochenchallenge-Videos (GitHub Actions).
Laeuft OHNE PC:
  1. laedt Quell-Clips + verlauf.json vom Server (/wc_video_source)
  2. baut leichtes + schweres Video mit build_step.py (identisch zum PC)
  3. laedt beide Videos nach /wochen-challenge-videos hoch
  4. schreibt verlauf.json zurueck (LRU-Abwechslung bleibt erhalten)
  5. schickt eine kurze E-Mail ("Video ist hochgeladen.")
Geplant: Donnerstag 22:00 (fuer die kommende Woche Sa-Di).
"""
import os, sys, glob, ssl, smtplib, traceback, shutil
from datetime import date, timedelta
from email.mime.text import MIMEText
import paramiko

HOST = os.environ["SSH_HOST"]
USER = os.environ["SSH_USER"]
PASS = os.environ["SSH_PASSWORD"]
PORT = int(os.environ.get("SSH_PORT", "22"))
GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_APP  = os.environ.get("GMAIL_APP_PASSWORD", "")
# Bestaetigung geht an Noahs persoenliche Adresse (+ optional EMAIL_TO-Secret)
RECIPIENTS = list(dict.fromkeys(
    r for r in ["noahwoe0212@gmail.com", os.environ.get("EMAIL_TO", "")] if r))

BASE    = os.path.dirname(os.path.abspath(__file__))
WORK    = os.path.join(BASE, "wc_video_work")
UPLOAD  = os.path.join(WORK, "Upload")
OUTPUT  = os.path.join(WORK, "Output")
FERTIG  = os.path.join(WORK, "Fertig")
VERLAUF = os.path.join(WORK, "verlauf.json")
STATE   = os.path.join(WORK, "wc_state.json")

REMOTE_SRC  = "/wc_video_source"           # Quell-Clips + verlauf.json
REMOTE_DEST = "/wochen-challenge-videos"   # Ziel der fertigen Videos

WC_EVENT = os.environ.get("WC_EVENT", "workflow_dispatch")  # 'schedule' oder 'workflow_dispatch'


def target_sam_str():
    """Naechster Samstag (wie build_step) -> Dateiname-Datum der Zielwoche."""
    today = date.today()
    days = (5 - today.weekday()) % 7
    if days == 0:
        days = 7
    sam = today + timedelta(days=days)
    return sam.strftime("%d.%m.%y")


def already_done_this_week():
    """True, wenn die Videos der Zielwoche schon auf dem Server liegen.
    Verhindert bei mehreren Cron-Zeiten einen zweiten (abweichenden) Build.
    Manuelles workflow_dispatch baut IMMER (kein Skip)."""
    if WC_EVENT != "schedule":
        return False
    sam = target_sam_str()
    want = {f"Wochenchallenge ab {sam} leicht.mp4", f"Wochenchallenge ab {sam} schwer.mp4"}
    sftp, t = sftp_open()
    try:
        have = set(sftp.listdir(REMOTE_DEST))
    except Exception:
        have = set()
    finally:
        sftp.close(); t.close()
    return want.issubset(have)


def sftp_open():
    t = paramiko.Transport((HOST, PORT))
    t.connect(username=USER, password=PASS)
    return paramiko.SFTPClient.from_transport(t), t


def download_sources():
    os.makedirs(UPLOAD, exist_ok=True)
    sftp, t = sftp_open()
    try:
        files = sftp.listdir(REMOTE_SRC)
        clips = [f for f in files if f.lower().endswith((".mp4", ".mov"))]
        for f in clips:
            sftp.get(f"{REMOTE_SRC}/{f}", os.path.join(UPLOAD, f))
        if "verlauf.json" in files:
            sftp.get(f"{REMOTE_SRC}/verlauf.json", VERLAUF)
    finally:
        sftp.close(); t.close()
    print(f"[download] {len(clips)} Clips | verlauf={'vorhanden' if os.path.exists(VERLAUF) else 'neu'}")
    return len(clips)


def build_videos():
    os.makedirs(OUTPUT, exist_ok=True)
    os.makedirs(FERTIG, exist_ok=True)
    # build_step.py aus dem Repo-Root importieren, Pfade auf WORK umbiegen
    sys.path.insert(0, BASE)
    import build_step as bs
    bs.UPLOAD, bs.OUTPUT, bs.FERTIG = UPLOAD, OUTPUT, FERTIG
    bs.VERLAUF, bs.STATE = VERLAUF, STATE
    guard = 0
    while True:
        done = bs.step()
        guard += 1
        if done:
            break
        if guard > 1000:
            raise RuntimeError("Build nicht fertig nach 1000 Schritten")
    vids = sorted(glob.glob(os.path.join(FERTIG, "**", "*.mp4"), recursive=True),
                  key=os.path.getmtime)
    latest = vids[-2:]
    if len(latest) < 2:
        raise RuntimeError(f"Nur {len(latest)} fertige Videos gefunden")
    print("[build] fertig:", [os.path.basename(v) for v in latest])
    return latest


def upload_results(videos):
    sftp, t = sftp_open()
    try:
        for v in videos:
            sftp.put(v, f"{REMOTE_DEST}/{os.path.basename(v)}")
            print("[upload]", os.path.basename(v))
        if os.path.exists(VERLAUF):
            sftp.put(VERLAUF, f"{REMOTE_SRC}/verlauf.json")
            print("[upload] verlauf.json zurueckgeschrieben")
    finally:
        sftp.close(); t.close()


def send_mail(subject, body):
    if not (GMAIL_USER and GMAIL_APP):
        print("[mail] keine Gmail-Zugangsdaten -> uebersprungen")
        return
    msg = MIMEText(body, _charset="utf-8")
    msg["Subject"] = subject
    msg["From"] = GMAIL_USER
    msg["To"] = ", ".join(RECIPIENTS)
    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ctx) as s:
        s.login(GMAIL_USER, GMAIL_APP)
        s.sendmail(GMAIL_USER, RECIPIENTS, msg.as_string())
    print("[mail] gesendet an", ", ".join(RECIPIENTS))


def main():
    if already_done_this_week():
        print(f"[skip] Videos fuer Woche ab {target_sam_str()} existieren bereits - kein erneuter Build (event={WC_EVENT}).")
        return
    if os.path.isdir(WORK):
        shutil.rmtree(WORK, ignore_errors=True)
    n = download_sources()
    if n == 0:
        raise RuntimeError("Keine Quell-Clips in " + REMOTE_SRC)
    videos = build_videos()
    upload_results(videos)
    names = ", ".join(os.path.basename(v) for v in videos)
    send_mail("Wochenchallenge Videos", "Video ist hochgeladen.\n\n" + names)
    print("DONE")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        tb = traceback.format_exc()
        print("ERROR:\n" + tb)
        try:
            send_mail("FEHLER Wochenchallenge Videos", "Fehler beim Erstellen:\n\n" + tb[-1500:])
        except Exception:
            pass
        sys.exit(1)
