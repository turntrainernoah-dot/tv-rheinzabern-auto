# -*- coding: utf-8 -*-
# Taeglicher System-Check: prueft letzten erfolgreichen Trainingsplan-Lauf, meldet per WhatsApp.
import os, json, urllib.request, urllib.parse, datetime
REPO=os.environ.get("GITHUB_REPOSITORY",""); TOKEN=os.environ.get("GH_TOKEN_AUTO","")
PHONE=os.environ.get("WHATSAPP_PHONE",""); KEY=os.environ.get("CALLMEBOT_APIKEY","")
def gh(path):
    req=urllib.request.Request(f"https://api.github.com/repos/{REPO}/{path}",
        headers={"Authorization":f"Bearer {TOKEN}","Accept":"application/vnd.github+json","User-Agent":"hb"})
    with urllib.request.urlopen(req,timeout=30) as r: return json.load(r)
def wa(text):
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
