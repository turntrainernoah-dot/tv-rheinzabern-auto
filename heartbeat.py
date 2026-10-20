# -*- coding: utf-8 -*-
# Taeglicher System-Check: prueft den letzten erfolgreichen Trainingsplan-Lauf.
# Seit 20.10.2026 keine E-Mail/WhatsApp mehr und keine "Alles laeuft"-Nachricht:
# NUR wenn etwas nicht stimmt, landet eine Meldung im Admin-Bereich
# ("Nachrichten", config/meldungen.json - siehe meldungen.py).
import os, json, urllib.request, datetime
from meldungen import melde_problem

REPO=os.environ.get("GITHUB_REPOSITORY",""); TOKEN=os.environ.get("GH_TOKEN_AUTO","")
BEREICH="System-Check"
WAS_TUN=("In GitHub im Repo tv-rheinzabern-auto unter \"Actions\" -> \"Automatischer "
         "Trainingsplan\" nachsehen, ob die Laeufe rot sind. Bitte Claude fragen und "
         "diese Meldung hineinkopieren.")
def gh(path):
    req=urllib.request.Request(f"https://api.github.com/repos/{REPO}/{path}",
        headers={"Authorization":f"Bearer {TOKEN}","Accept":"application/vnd.github+json","User-Agent":"hb"})
    with urllib.request.urlopen(req,timeout=30) as r: return json.load(r)
now=datetime.datetime.now(datetime.timezone.utc)
try:
    runs=gh("actions/workflows/auto_trainingsplan.yml/runs?per_page=20").get("workflow_runs",[])
except Exception as e:
    melde_problem(BEREICH,
        "Der taegliche System-Check konnte nicht pruefen, ob die Automatik laeuft "
        "(GitHub war nicht erreichbar). Kommt das nur einmal vor, ist es harmlos. "
        "Kommt es mehrere Tage hintereinander: Claude fragen und diese Meldung hineinkopieren.",
        "GitHub API: %s"%str(e)[:500])
    raise SystemExit(0)
succ=[r for r in runs if r.get("conclusion")=="success"]
if succ:
    lt=datetime.datetime.fromisoformat(succ[0]["updated_at"].replace("Z","+00:00"))
    hrs=(now-lt).total_seconds()/3600
    if hrs<=3:
        print("System-Check: Alles laeuft. Letzter erfolgreicher Trainingsplan-Lauf vor %.1f Std."%hrs)
    else:
        melde_problem(BEREICH,
            "Die Trainingsplan-Automatik ist seit %.0f Stunden nicht mehr erfolgreich "
            "gelaufen. Neue Abmeldungen kommen evtl. nicht im Plan an. %s"%(hrs, WAS_TUN),
            "Letzter erfolgreicher Lauf: %s (vor %.1f Std)"%(succ[0].get("updated_at"), hrs))
else:
    melde_problem(BEREICH,
        "Die Trainingsplan-Automatik ist in den letzten 20 Laeufen kein einziges Mal "
        "erfolgreich gelaufen. Neue Abmeldungen kommen nicht im Plan an. %s"%WAS_TUN,
        "Laeufe: " + ", ".join("%s=%s"%(r.get("updated_at"), r.get("conclusion")) for r in runs[:5]))
print("[FERTIG]")
