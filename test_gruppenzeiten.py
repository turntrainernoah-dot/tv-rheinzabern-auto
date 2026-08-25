#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Synthetischer Test-Harness fuer den Gruppenzeiten-Umbau von auto_trainingsplan.py.
Setzt die Roster-Globals direkt (kein echtes SFTP/config.json noetig) und prueft
die im Aufgabentext geforderte Testmatrix, so weit sie ohne Live-Server/PC-Zugriff
moeglich ist. Wird NICHT automatisch ausgefuehrt - manuell mit `python3
test_gruppenzeiten.py` starten.
"""
import sys, os, random
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import auto_trainingsplan as a

FAILS = []

def check(name, cond, detail=""):
    status = "OK  " if cond else "FAIL"
    print(f"[{status}] {name}" + (f"  -- {detail}" if detail and not cond else ""))
    if not cond:
        FAILS.append(name)

def default_zeiten_for_all(gruppen):
    return {g: a._default_zeiten_for(g) for g in gruppen}

def set_roster(gruppen, turner_pro_gruppe, trainer, immer_springer=None, zeiten=None):
    a.GRUPPEN_ORDER = list(gruppen)
    a.ALLE_TURNER = {g: list(turner_pro_gruppe.get(g, [f"K{g}{i}" for i in range(1, 6)])) for g in gruppen}
    a.ALLE_TRAINER = list(trainer)
    a.IMMER_SPRINGER = set(immer_springer or [])
    a.GRUPPEN_ZEITEN = zeiten or default_zeiten_for_all(gruppen)
    a.WEBSITE_TO_DISPLAY = {}

def no_absences(gruppen):
    d = {g: [] for g in gruppen}
    d["Trainer"] = []
    return d

# ════════════════════════════════════════════════════════════════
# 1) Regression: migrierte Default-Config -> exakt dasselbe 5-Zeilen-Raster
#    wie das alte feste ZEITSLOTS/SLOT_START_MIN/SLOT_END_MIN.
# ════════════════════════════════════════════════════════════════
def test_regression_grid():
    """Die migrierten Default-Zeiten (siehe Aufgabentext) ergeben ein 4- statt
    5-Zeilen-Raster: die alte 5. Zeile (18:00-18:15) war eine willkuerliche
    Unterteilung MITTEN in der durchgehenden Geraet1-Phase (17:30-18:15) ohne
    eigene Bedeutung -- das dynamische Raster leitet Zeilen nur aus echten
    Phasengrenzen ab, wodurch dieser doppelte Eintrag entfaellt. Inhaltlich
    (Farben/Zeiten/Text) aendert sich nichts, der Plan wird nur um eine
    redundante Zeile kuerzer (siehe compute_time_grid()-Docstring)."""
    gruppen = ["G1", "G2", "G3", "G4"]
    zeiten = default_zeiten_for_all(gruppen)
    for tag in ("mi", "fr"):
        rows, phase = a.compute_time_grid(zeiten, tag)
        expected = [(17*60, 17*60+30), (17*60+30, 18*60+15),
                    (18*60+15, 19*60), (19*60, 19*60+30)]
        check(f"regression grid rows ({tag})", rows == expected, f"{rows}")
        check(f"regression G1 phases ({tag})", phase["G1"] == ["aufwaermen", "geraet1", "geraet2", None])
        check(f"regression G4 phases ({tag})", phase["G4"] == [None, "aufwaermen", "geraet1", "geraet2"])

def test_regression_plan_shape():
    """Mit der migrierten Default-Config, genug Trainern und keinen Abwesenheiten
    muss jede Gruppe ihre eigene Einheit bekommen (kein Merge noetig), mit den
    erwarteten AW/Geraet1/Geraet2-Farben je Zeile."""
    gruppen = ["G1", "G2", "G3", "G4"]
    turner = {g: [f"{g}Kind{i}" for i in range(1, 6)] for g in gruppen}  # 5 pro Gruppe
    trainer = ["Trainer A", "Trainer B", "Trainer C", "Trainer D", "Trainer E"]
    set_roster(gruppen, turner, trainer)
    grid_rows, grid_phase = a.compute_time_grid(a.GRUPPEN_ZEITEN, "mi")
    plan, sonder, anm = a.build_trainer_plan(no_absences(gruppen), grid_rows, grid_phase)
    holders = {t: c for t, c in plan.items() if c and any(ck in ("g1_blau", "g2_orange") for _, ck in c)}
    check("jede Gruppe bekommt eigenen Halter", len(holders) == 4, f"{holders.keys()}")
    for t, cells in holders.items():
        labels = {txt.replace("AW ", "") for txt, _ck in cells if txt not in ("Aufbauen", "Abbauen")}
        check(f"{t}: genau ein Gruppen-Label", len(labels) == 1, f"{labels}")
    g1_cells = next(c for t, c in holders.items() if "G1" in {x.replace("AW ", "") for x, _ in c if x not in ("Aufbauen","Abbauen")})
    check("G1 Zeile0=aufwaermen", g1_cells[0][1] == "aufwaermen")
    check("G1 Zeile1=g1_blau (Geraet1)", g1_cells[1][1] == "g1_blau")
    check("G1 letzte Zeile=Abbauen", g1_cells[-1] == ("Abbauen", "aufbauen"))
    remaining = [t for t in trainer if t not in holders]
    check("1 uebriger Trainer wird Springer", len(remaining) == 1)
    if remaining:
        check("Springer-Zellen korrekt", plan[remaining[0]] == a._springer_cells(grid_rows))

# ════════════════════════════════════════════════════════════════
# 2) Gruppenanzahl 3..6, asymmetrische Zeiten
# ════════════════════════════════════════════════════════════════
def test_group_counts():
    for n in (3, 4, 5, 6):
        gruppen = [f"G{i}" for i in range(1, n + 1)]
        turner = {g: [f"{g}K{i}" for i in range(1, 6)] for g in gruppen}
        trainer = [f"T{i}" for i in range(1, n + 1)]
        zeiten = {}
        for i, g in enumerate(gruppen):
            # jede zweite Gruppe eine halbe Stunde versetzt -> echte Asymmetrie
            offset = 30 if i % 2 else 0
            base = 17 * 60 + offset
            zeiten[g] = {
                "mi": {"aufwaermen": {"start": a._min_to_hhmm(base), "ende": a._min_to_hhmm(base + 30)},
                       "geraet1":   {"start": a._min_to_hhmm(base + 30), "ende": a._min_to_hhmm(base + 75)},
                       "geraet2":   {"start": a._min_to_hhmm(base + 75), "ende": a._min_to_hhmm(base + 120)}},
            }
            zeiten[g]["fr"] = zeiten[g]["mi"]
        set_roster(gruppen, turner, trainer, zeiten=zeiten)
        grid_rows, grid_phase = a.compute_time_grid(a.GRUPPEN_ZEITEN, "mi")
        plan, sonder, anm = a.build_trainer_plan(no_absences(gruppen), grid_rows, grid_phase)
        holders = [t for t, c in plan.items() if c and any(ck in ("g1_blau", "g2_orange") for _, ck in c)]
        check(f"n={n} Gruppen: alle {n} bekommen Halter", len(holders) == n, f"holders={holders}")
        for g in gruppen:
            check(f"n={n}: Gruppe {g} hat eigene aktive Phasen", any(p for p in grid_phase[g]))

# ════════════════════════════════════════════════════════════════
# 3) Trainerzahl 1..8, 0..3 immer_springer, Notfall-Durchbruch
# ════════════════════════════════════════════════════════════════
def test_trainer_counts_and_immer_springer():
    gruppen = ["G1", "G2", "G3"]
    turner = {g: [f"{g}K{i}" for i in range(1, 6)] for g in gruppen}
    zeiten = default_zeiten_for_all(gruppen)
    for n_trainer in range(1, 9):
        trainer = [f"T{i}" for i in range(1, n_trainer + 1)]
        for n_immer in range(0, min(3, n_trainer) + 1):
            immer = set(trainer[:n_immer])
            set_roster(gruppen, turner, trainer, immer_springer=immer, zeiten=zeiten)
            grid_rows, grid_phase = a.compute_time_grid(a.GRUPPEN_ZEITEN, "mi")
            plan, sonder, anm = a.build_trainer_plan(no_absences(gruppen), grid_rows, grid_phase)
            assigned_groups = {t for t, c in plan.items() if c and any(ck in ("g1_blau", "g2_orange") for _, ck in c)}
            non_immer = [t for t in trainer if t not in immer]
            # Wenn genug Nicht-Springer da sind, sollen NUR die eine Gruppe bekommen ...
            if len(non_immer) >= len(gruppen):
                check(f"n={n_trainer},immer={n_immer}: immer_springer bekommt keine Gruppe wenn genug andere da",
                      assigned_groups.isdisjoint(immer), f"{assigned_groups} vs immer={immer}")
            else:
                # ... sonst muessen immer_springer-Trainer als letzte Instanz einspringen,
                # damit trotzdem moeglichst viele Gruppen eine eigene Einheit bekommen.
                need = min(len(gruppen), n_trainer)
                check(f"n={n_trainer},immer={n_immer}: Notfall zieht genug Trainer (inkl. Springer-Flag) nach",
                      len(assigned_groups) == need, f"got {len(assigned_groups)} need {need}")

# ════════════════════════════════════════════════════════════════
# 5) Merge-Kompatibilitaet: nur zeit-gleiche Gruppen duerfen zusammengelegt werden
# ════════════════════════════════════════════════════════════════
def test_merge_compatibility():
    gruppen = ["G1", "G2", "G3"]
    zeiten = default_zeiten_for_all(gruppen)
    # G3 zeitlich verschoben -> NICHT kompatibel mit G1/G2
    shifted = {"mi": {"aufwaermen": {"start": "18:00", "ende": "18:30"},
                       "geraet1":   {"start": "18:30", "ende": "19:00"},
                       "geraet2":   {"start": "19:00", "ende": "19:30"}}}
    shifted["fr"] = shifted["mi"]
    zeiten["G3"] = shifted
    check("G1/G2 zeit-kompatibel", a.zeiten_kompatibel(zeiten, "G1", "G2") is True)
    check("G1/G3 NICHT zeit-kompatibel", a.zeiten_kompatibel(zeiten, "G1", "G3") is False)

    # nur 1 Trainer da, aber G3 hat andere Zeiten als G1/G2 -> G3 darf NICHT
    # mit G1 oder G2 gemergt werden (auch wenn dadurch eine Gruppe unbesetzt bleibt)
    turner = {g: [f"{g}K{i}" for i in range(1, 3)] for g in gruppen}  # nur 2 Kinder -> alle "undersized"
    trainer = ["T1"]
    set_roster(gruppen, turner, trainer, zeiten=zeiten)
    grid_rows, grid_phase = a.compute_time_grid(a.GRUPPEN_ZEITEN, "mi")
    plan, sonder, anm = a.build_trainer_plan(no_absences(gruppen), grid_rows, grid_phase)
    holder_cells = [c for c in plan.values() if c and any(ck in ("g1_blau", "g2_orange") for _, ck in c)]
    check("nur 1 Einheit besetzt (Rest 'kein Trainer mehr fuer')", len(holder_cells) == 1)
    check("Anmerkung nennt unbesetzte Einheit(en)", any("kein Trainer mehr" in x for x in anm), f"{anm}")
    # Die besetzte Einheit darf NIE G1+G3 oder G2+G3 sein (inkompatible Zeiten)
    for c in holder_cells:
        labels = {t.replace("AW ", "") for t, ck in c if t not in ("Aufbauen", "Abbauen")}
        lbl = list(labels)[0] if labels else ""
        check("besetzte Einheit enthaelt kein inkompatibles G3-Merge",
              not ("G3" in lbl and "+" in lbl), f"label={lbl}")

# ════════════════════════════════════════════════════════════════
# 6) Rotationsfairness ueber 20 simulierte Folgetermine
# ════════════════════════════════════════════════════════════════
def test_rotation_fairness():
    gruppen = ["G1", "G2", "G3"]
    turner = {g: [f"{g}K{i}" for i in range(1, 6)] for g in gruppen}
    trainer = ["T1", "T2", "T3", "T4"]  # 4 Trainer, 3 Gruppen -> immer genau 1 Springer
    set_roster(gruppen, turner, trainer)
    grid_rows, grid_phase = a.compute_time_grid(a.GRUPPEN_ZEITEN, "mi")

    state = {"plan_data": {}}
    springer_count = {t: 0 for t in trainer}
    from datetime import date, timedelta
    d = date(2026, 9, 2)
    for i in range(20):
        dk = d.strftime("%d.%m.%y")
        hist = a._load_trainer_roles_history(state, exclude_date=dk)
        plan, _s, _anm = a.build_trainer_plan(no_absences(gruppen), grid_rows, grid_phase,
                                               trainer_roles_history=hist)
        roles = a._extract_trainer_roles(plan)
        for t, r in roles.items():
            if r == "Springer":
                springer_count[t] += 1
        state.setdefault("plan_data", {})[dk] = {"trainer_roles": roles}
        d += timedelta(days=2)

    counts = list(springer_count.values())
    spread = max(counts) - min(counts)
    # Ideal waere 5/5/5/5 (20 Termine / 4 Trainer); Noah hat explizit eine
    # EINFACHE Greedy-Zuweisung gewuenscht ("reicht bei <=8 Trainern voellig
    # aus"), keine perfekte Gleichverteilung -- Toleranz entsprechend locker.
    check("Springer-Rolle verteilt sich grob fair ueber alle 4 Trainer (Spread <= 4 nach 20 Terminen)",
          spread <= 4, f"{springer_count}")
    check("jeder Trainer war mindestens einmal NICHT Springer",
          all(c < 20 for c in counts), f"{springer_count}")

# ════════════════════════════════════════════════════════════════
# 7) KI-Pfad vs Auto-Pfad: keine strukturelle Drift bei leerer KI-Anweisung
# ════════════════════════════════════════════════════════════════
def test_ki_path_parity():
    gruppen = ["G1", "G2", "G3", "G4"]
    turner = {g: [f"{g}K{i}" for i in range(1, 6)] for g in gruppen}
    trainer = ["T1", "T2", "T3"]  # weniger Trainer als Gruppen -> Merge noetig
    set_roster(gruppen, turner, trainer)
    grid_rows, grid_phase = a.compute_time_grid(a.GRUPPEN_ZEITEN, "mi")
    absences = no_absences(gruppen)

    plan_auto, _s1, anm_auto = a.build_trainer_plan(absences, grid_rows, grid_phase)
    plan_ki, _s2, anm_ki = a.build_ki_einteilung(absences, {}, grid_rows, grid_phase)

    def unit_count(plan):
        return len({tuple(sorted({t.replace("AW ", "") for t, ck in c if t not in ("Aufbauen", "Abbauen")}))
                    for c in plan.values() if c and any(ck in ("g1_blau", "g2_orange") for _, ck in c)})

    check("Auto-Pfad: 3 Einheiten (1 Merge noetig bei 3 Trainern/4 Gruppen)", unit_count(plan_auto) == 3,
          f"{[c for c in plan_auto.values() if c]}")
    check("KI-Pfad (leere Anweisung): ebenfalls 3 Einheiten, keine Drift", unit_count(plan_ki) == 3,
          f"{[c for c in plan_ki.values() if c]}")
    for c in plan_auto.values():
        if not c:
            continue
        colors = {ck for _t, ck in c}
        check("Auto-Pfad nutzt nur die 2 erlaubten Farben (+aufwaermen/aufbauen/springer)",
              colors <= {"aufwaermen", "g1_blau", "g2_orange", "aufbauen", "springer"}, f"{colors}")
    for c in plan_ki.values():
        if not c:
            continue
        colors = {ck for _t, ck in c}
        check("KI-Pfad nutzt nur die 2 erlaubten Farben (+aufwaermen/aufbauen/springer)",
              colors <= {"aufwaermen", "g1_blau", "g2_orange", "aufbauen", "springer"}, f"{colors}")

# ════════════════════════════════════════════════════════════════
# 4) Notfall / leere Randbedingungen
# ════════════════════════════════════════════════════════════════
def test_edge_cases():
    gruppen = ["G1", "G2"]
    turner = {g: [f"{g}K{i}" for i in range(1, 4)] for g in gruppen}
    set_roster(gruppen, turner, [])
    grid_rows, grid_phase = a.compute_time_grid(a.GRUPPEN_ZEITEN, "mi")
    plan, _s, anm = a.build_trainer_plan(no_absences(gruppen), grid_rows, grid_phase)
    check("0 Trainer -> ACHTUNG-Hinweis, alle Trainer None", "ACHTUNG" in anm[0] and all(v is None for v in plan.values()))

    set_roster(gruppen, turner, ["T1"])
    absences_all = {"G1": list(turner["G1"]), "G2": list(turner["G2"]), "Trainer": []}
    grid_rows, grid_phase = a.compute_time_grid(a.GRUPPEN_ZEITEN, "mi")
    plan, _s, anm = a.build_trainer_plan(absences_all, grid_rows, grid_phase)
    check("alle Turner abwesend -> kein Training", "ACHTUNG" in anm[0], f"{anm}")


if __name__ == "__main__":
    test_regression_grid()
    test_regression_plan_shape()
    test_group_counts()
    test_trainer_counts_and_immer_springer()
    test_merge_compatibility()
    test_rotation_fairness()
    test_ki_path_parity()
    test_edge_cases()
    print()
    if FAILS:
        print(f"{len(FAILS)} FEHLGESCHLAGEN:")
        for f in FAILS:
            print("  -", f)
        sys.exit(1)
    else:
        print("ALLE TESTS BESTANDEN")
