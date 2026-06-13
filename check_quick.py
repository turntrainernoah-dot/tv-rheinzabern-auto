#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_quick.py - Mikro-Check fuer GitHub Actions
=================================================
Lauft in < 1 Minute (nur paramiko, kein openpyxl, kein LibreOffice).
Setzt GITHUB_OUTPUT: needs_update=true/false

Logik:
  Kein Plan vorhanden + Veroeffentlichungsfenster (Mi/Fr 20:00+ UTC) → true
  Plan vorhanden + Abwesenheiten/Anmerkungen geaendert              → true
  Sonst                                                              → false
"""

import json, hashlib, os
from datetime import date, timedelta, datetime, timezone

import paramiko

SSH_HOST     = os.environ.get("SSH_HOST",     "access-5017462830.webspace-host.com")
SSH_USER     = os.environ.get("SSH_USER",     "a2358459")
SSH_PASSWORD = os.environ.get("SSH_PASSWORD", "")
SSH_PORT     = int(os.environ.get("SSH_PORT", "22"))

# ════════════════════════════════════════════════════════════
def set_output(key, value):
    path = os.environ.get("GITHUB_OUTPUT", "")
    if path:
        with open(path, "a") as f:
            f.write(f"{key}={value}\n")
    print(f"  OUTPUT → {key}={value}")

def next_training_date():
    """Naechster Trainingstag (Mi=2 oder Fr=4) ab morgen."""
    d = date.today() + timedelta(days=1)
    while d.weekday() not in (2, 4):
        d += timedelta(days=1)
    return d

def is_publication_window():
    """Mi oder Fr nach 20:00 UTC (= 22:00 CEST Sommerzeit)."""
    now = datetime.now(timezone.utc)
    return now.weekday() in (2, 4) and now.hour >= 20

def is_within_creation_window(next_date):
    """Erlaubt Plan-Erstellung wenn Training innerhalb von 5 Tagen liegt
    ODER im Publikationsfenster (Mi/Fr 22:00 CEST)."""
    days_away = (next_date - date.today()).days
    return days_away <= 5 or is_publication_window()

def md5(data: bytes) -> str:
    return hashlib.md5(data).hexdigest()

# ════════════════════════════════════════════════════════════
def main():
    print("=== Trainingsplan Schnell-Check ===")
    next_date  = next_training_date()
    datum_kurz = next_date.strftime("%d.%m.%y")
    print(f"Naechstes Training: {datum_kurz}")

    # SFTP verbinden
    try:
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(SSH_HOST, port=SSH_PORT, username=SSH_USER,
                       password=SSH_PASSWORD, timeout=15)
        sftp = client.open_sftp()
        print("SFTP verbunden.")
    except Exception as e:
        print(f"[ERROR] SFTP-Verbindung fehlgeschlagen: {e}")
        set_output("needs_update", "false")
        return

    # State laden
    try:
        f     = sftp.open("state_auto.json", "r")
        state = json.loads(f.read().decode("utf-8"))
        f.close()
    except Exception:
        state = {}

    # Primär: State-Check. Sekundär: SFTP-Stat (robust gegen manuelle Uploads/Löschungen)
    in_state = datum_kurz in state.get("generated_plans", [])
    plan_exists = in_state
    if in_state:
        try:
            sftp.stat(f"trainingspläne/{datum_kurz}_Trainingsplan.pdf")
        except FileNotFoundError:
            print(f"Plan in State, aber PDF nicht auf Server – behandle als fehlend")
            plan_exists = False
        except Exception:
            pass  # Stat-Fehler: lieber annehmen Plan existiert
    else:
        # Auch manuell hochgeladene Pläne erkennen (nicht im State)
        try:
            sftp.stat(f"trainingspläne/{datum_kurz}_Trainingsplan.pdf")
            plan_exists = True
            print(f"Plan auf Server gefunden (manuell hochgeladen, nicht im State)")
        except Exception:
            pass
    print(f"Plan vorhanden: {plan_exists} (state={in_state})")

    # ── Kein Plan vorhanden ────────────────────────────────
    if not plan_exists:
        sftp.close(); client.close()
        if is_within_creation_window(next_date):
            now = datetime.now(timezone.utc)
            days_away = (next_date - date.today()).days
            print(f"Kein Plan + {days_away} Tage bis Training ({now.strftime('%H:%M')} UTC) → ERSTELLEN")
            set_output("needs_update", "true")
        else:
            now = datetime.now(timezone.utc)
            days_away = (next_date - date.today()).days
            print(f"Kein Plan, Training in {days_away} Tagen ({now.strftime('%H:%M')} UTC) → WARTEN")
            set_output("needs_update", "false")
        return

    # ── Plan vorhanden: Abwesenheits-Hash vergleichen ─────
    plan_data   = state.get("plan_data", {}).get(datum_kurz, {})
    stored_hash = plan_data.get("absences_hash", "")

    if not stored_hash:
        # Plan ohne gespeicherten Hash (manuell erstellt) → Hash-Tracking initialisieren
        print(f"Kein gespeicherter Hash fuer {datum_kurz} → Hash-Tracking starten")
        sftp.close(); client.close()
        set_output("needs_update", "true")
        return

    # Abmeldungs-Hash laden
    current_abs_hash = ""
    try:
        f = sftp.open("abmeldungen/abmeldungen.json", "r")
        raw = f.read()
        f.close()
        current_abs_hash = md5(raw)
    except Exception as e:
        print(f"[WARN] Abmeldungen nicht lesbar: {e}")

    # Ungelesene Anmerkungen pruefen
    has_new_anmerkungen = False
    try:
        f    = sftp.open("anmerkungen/anmerkungen.json", "r")
        anms = json.loads(f.read().decode("utf-8"))
        f.close()
        unread = [a for a in anms if not a.get("gelesen", False)]
        if unread:
            print(f"Ungelesene Anmerkungen: {len(unread)} → triggert Update")
            has_new_anmerkungen = True
    except Exception:
        pass

    sftp.close()
    client.close()

    abs_changed = (current_abs_hash != stored_hash and current_abs_hash != "")

    if abs_changed:
        print(f"Abmeldungsaenderung! Alt: {stored_hash[:8]}… → Neu: {current_abs_hash[:8]}…")
    elif not has_new_anmerkungen:
        print(f"Keine Aenderung (Hash: {stored_hash[:8]}…)")

    if abs_changed or has_new_anmerkungen:
        set_output("needs_update", "true")
    else:
        set_output("needs_update", "false")

# ════════════════════════════════════════════════════════════
if __name__ == "__main__":
    main()
