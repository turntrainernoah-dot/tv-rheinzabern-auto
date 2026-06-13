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

    plan_exists = datum_kurz in state.get("generated_plans", [])
    print(f"Plan vorhanden: {plan_exists}")

    # ── Kein Plan vorhanden ────────────────────────────────
    if not plan_exists:
        sftp.close(); client.close()
        if is_publication_window():
            now = datetime.now(timezone.utc)
            print(f"Kein Plan + Veroeffentlichungsfenster ({now.strftime('%H:%M')} UTC) → ERSTELLEN")
            set_output("needs_update", "true")
        else:
            now = datetime.now(timezone.utc)
            print(f"Kein Plan, aber ausserhalb Fenster ({now.strftime('%H:%M')} UTC) → WARTEN")
            set_output("needs_update", "false")
        return

    # ── Plan vorhanden: Abwesenheits-Hash vergleichen ─────
    plan_data   = state.get("plan_data", {}).get(datum_kurz, {})
    stored_hash = plan_data.get("absences_hash", "")

    if not stored_hash:
        # Alter Plan ohne gespeicherten Hash (vor dem Update) – nicht anfassen
        print(f"Kein gespeicherter Hash fuer {datum_kurz} (alter Plan) → unveraendert lassen")
        # Aber noch Anmerkungen pruefen (koennen trotzdem relevant sein)
        # → bei alten Plaenen lieber nicht anfassen, Anmerkungen beim naechsten neuen Plan
        sftp.close(); client.close()
        set_output("needs_update", "false")
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
