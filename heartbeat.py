# -*- coding: utf-8 -*-
# Taeglicher System-Check: prueft letzten erfolgreichen Trainingsplan-Lauf, meldet per WhatsApp.
import os, json, urllib.request, urllib.parse, datetime
REPO=os.environ.get("GITHUB_REPOSITORY",""); TOKEN=os.environ.get("GH_TOKEN_AUTO","")
PHONE=os.environ.get("WHATSAPP_PHONE",""); KEY=os.environ.get("CALLMEBOT_APIKEY","")
def gh(path):
    req=urllib.request.Request(f"https://api.github.com/repos/{REPO}/{path}",
        headers={"Authorization":f"Bearer {TOKEN}","Accept":"application/vnd.github+json","User-Agent":"hb"})
    with urllib.request.urlopen(req,timeout=30) as r: return json.load(r)
def send_email(text):
    import os, smtplib, ssl as _ssl
    from email.mime.text import MIMEText
    user = os.environ.get("GMAIL_USER", "")
    pw   = os.environ.get("GMAIL_APP_PASSWORD", "")
    to   = os.environ.get("EMAIL_TO", "") or user
    if not user or not pw:
        print("[MAIL] keine Gmail-Zugangsdaten - uebersprungen.")
        return
    try:
        lines = [l for l in text.strip().splitlines() if l.strip()]
        subj = "TV Rheinzabern: " + (lines[0][:80] if lines else "Info")
        msg = MIMEText(text, _charset="utf-8")
        msg["Subject"] = subj
        msg["From"] = user
        msg["To"] = to
        ctx = _ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ctx) as s:
            s.login(user, pw)
            s.sendmail(user, [to], msg.as_string())
        print("[MAIL] gesendet an", to)
    except Exception as e:
        print("[MAIL] Fehler:", e)


def wa(text):
    send_email(text)
    if not PHONE or not KEY: print("[WA-TEST]", text); return
    url="https://api.callmebot.com/whatsapp.php?phone=%s&text=%s&apikey=%s"%(PHONE,urllib.parse.quote(text),KEY)
    try:
        with urllib.request.urlopen(url,timeout=20) as r: print("WA HTTP", r.status)
    except Exception as e: print("WA Fehler", e)
now=datetime.datetime.now(datetime.timezone.utc)
try:
    runs=gh("actions/workflows/auto_trainingsplan.yml/runs?per_page=20").get("workflow_runs",[])
except Exception as e:
    wa("Hi Noah, Cloude hier 🚨\n\nSystem-Check fehlgeschlagen (GitHub API): %s"%str(e)[:150]); raise SystemExit(0)
succ=[r for r in runs if r.get("conclusion")=="success"]
if succ:
    lt=datetime.datetime.fromisoformat(succ[0]["updated_at"].replace("Z","+00:00"))
    hrs=(now-lt).total_seconds()/3600
    if hrs<=3:
        wa("Hi Noah, Cloude hier ✅\n\nSystem-Check: Alles laeuft. Letzter erfolgreicher Trainingsplan-Lauf vor %.1f Std."%hrs)
    else:
        wa("Hi Noah, Cloude hier ⚠️\n\nSystem-Check: Seit %.1f Std KEIN erfolgreicher Trainingsplan-Lauf. Bitte GitHub Actions + Token (Ablauf 12.07!) pruefen."%hrs)
else:
    wa("Hi Noah, Cloude hier 🚨\n\nSystem-Check: In den letzten 20 Laeufen KEIN Erfolg. Bitte pruefen.")
print("[FERTIG]")
