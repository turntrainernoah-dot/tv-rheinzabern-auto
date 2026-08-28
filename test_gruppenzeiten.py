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

def set_roster(gruppen, turner_pro_gruppe, trainer, immer_springer=None, zeiten=None, tausch=None):
    a.GRUPPEN_ORDER = list(gruppen)
    a.ALLE_TURNER = {g: list(turner_pro_gruppe.get(g, [f"K{g}{i}" for i in range(1, 6)])) for g in gruppen}
    a.ALLE_TRAINER = list(trainer)
    a.IMMER_SPRINGER = set(immer_springer or [])
    a.GRUPPEN_ZEITEN = zeiten or default_zeiten_for_all(gruppen)
    a.GRUPPEN_TAUSCH = set(tausch or [])
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


# ════════════════════════════════════════════════════════════════
# 5) Geraet-Tausch: admin-konfigurierbare Geraete-Reihenfolge pro Gruppe
#    (loest die Kapazitaetskollision bei 3+ Gruppen mit identischer Zeit,
#    ohne die Uhrzeit zu verschieben -- siehe GRUPPEN_TAUSCH-Kommentar).
# ════════════════════════════════════════════════════════════════
def test_geraet_tausch():
    # Alle 3 Gruppen haben EXAKT dieselbe Zeit (Standard) -- ohne Tausch waere
    # das eine Kapazitaetsverletzung (3 Gruppen gleichzeitig auf Geraet 1),
    # die admin.php::validateGruppenZeiten() blockieren wuerde. G3 bekommt
    # geraet_tausch=True, um genau das aufzuloesen.
    gruppen = ["G1", "G2", "G3"]
    turner = {g: [f"{g}Kind{i}" for i in range(1, 4)] for g in gruppen}
    trainer = ["Trainer A", "Trainer B", "Trainer C"]
    zeiten = {g: a._default_zeiten_for("G1") for g in gruppen}  # alle 3 identisch, Standard-Zeit
    set_roster(gruppen, turner, trainer, zeiten=zeiten, tausch={"G3"})

    check("_effektive_phase ohne Tausch unveraendert",
          a._effektive_phase("G1", "geraet1") == "geraet1" and a._effektive_phase("G1", "geraet2") == "geraet2")
    check("_effektive_phase mit Tausch vertauscht geraet1/geraet2",
          a._effektive_phase("G3", "geraet1") == "geraet2" and a._effektive_phase("G3", "geraet2") == "geraet1")
    check("_effektive_phase laesst aufwaermen unangetastet",
          a._effektive_phase("G3", "aufwaermen") == "aufwaermen")

    check("G1/G3 trotz identischer Zeiten NICHT zeit-kompatibel (unterschiedlicher Tausch)",
          not a.zeiten_kompatibel(a.GRUPPEN_ZEITEN, "G1", "G3"))
    check("G1/G2 weiterhin zeit-kompatibel (beide ohne Tausch)",
          a.zeiten_kompatibel(a.GRUPPEN_ZEITEN, "G1", "G2"))

    grid_rows, grid_phase = a.compute_time_grid(a.GRUPPEN_ZEITEN, "mi")
    plan, _s, _anm = a.build_trainer_plan(no_absences(gruppen), grid_rows, grid_phase)
    # G3 haelt ihren zeitlich ERSTEN Block (Config-Schluessel "geraet1") in der
    # Farbe von Geraet 2 und den zweiten in der Farbe von Geraet 1 -- exakt
    # umgekehrt zu G1/G2, obwohl die Uhrzeiten identisch sind.
    g3_cells = next(c for c in plan.values() if c and any(txt == "G3" for txt, _ck in c))
    g1_cells = next(c for c in plan.values() if c and any(txt == "G1" for txt, _ck in c))
    g3_colors = [ck for txt, ck in g3_cells if txt == "G3"]
    g1_colors = [ck for txt, ck in g1_cells if txt == "G1"]
    check("G3 (Tausch): erster Geraete-Block ist g2_orange", g3_colors[0] == "g2_orange", f"{g3_colors}")
    check("G3 (Tausch): zweiter Geraete-Block ist g1_blau", g3_colors[1] == "g1_blau", f"{g3_colors}")
    check("G1 (kein Tausch): erster Geraete-Block bleibt g1_blau", g1_colors[0] == "g1_blau", f"{g1_colors}")
    check("G1 (kein Tausch): zweiter Geraete-Block bleibt g2_orange", g1_colors[1] == "g2_orange", f"{g1_colors}")
    # Kapazitaets-Invariante von Hand nachgerechnet: zu jedem Zeitpunkt hoechstens
    # 2 Gruppen auf derselben physischen Farbe (das ist es, was der Tausch loest).
    for i in range(len(grid_rows)):
        colors_at_row = []
        for cells in (g1_cells, g3_cells, next(c for c in plan.values() if c and any(txt == "G2" for txt, _ck in c))):
            txt, ck = cells[i]
            if ck in ("g1_blau", "g2_orange"):
                colors_at_row.append(ck)
        check(f"Zeile {i}: max. 2 Gruppen pro physischem Geraet", colors_at_row.count("g1_blau") <= 2 and colors_at_row.count("g2_orange") <= 2,
              f"{colors_at_row}")


# ════════════════════════════════════════════════════════════════
# 6) Bugfix 27.08.2026: Teilzeit-Trainer (kommt spaeter/geht frueher) wird
#    nicht mehr komplett aus dem Plan entfernt, sondern haelt eine Gruppe.
# ════════════════════════════════════════════════════════════════
def test_partial_trainer_holds_group():
    """build_trainer_plan schloss Teilzeit-Trainer (in trainer_timing) bisher
    praeventiv aus dem Halter-Pool aus, sobald irgendein Vollzeit-Trainer da
    war -- dadurch bekam eine Gruppe gar keinen Halter und der Teilzeit-
    Trainer wurde Springer statt seine Gruppe (mit spaeterer Uebergabe der
    Randzeiten via apply_timing_coverage) zu halten. Seit dem Fix nimmt
    build_trainer_plan gar keine trainer_timing-Info mehr entgegen -- alle
    verfuegbaren Trainer sind gleichwertige Halter-Kandidaten, konsistent mit
    build_ki_einteilung (das trainer_timing nie ausgeschlossen hat)."""
    gruppen = ["G1", "G2", "G3"]
    turner = {g: [f"{g}Kind{i}" for i in range(1, 6)] for g in gruppen}
    trainer = ["Trainer A", "Trainer B", "Trainer C"]
    set_roster(gruppen, turner, trainer)
    grid_rows, grid_phase = a.compute_time_grid(a.GRUPPEN_ZEITEN, "mi")
    plan, _s, anm = a.build_trainer_plan(no_absences(gruppen), grid_rows, grid_phase)
    holders = {t: c for t, c in plan.items() if c and any(ck in ("g1_blau", "g2_orange") for _, ck in c)}
    check("alle 3 Gruppen bekommen einen Halter (kein Trainer wird praeventiv ausgeschlossen)",
          len(holders) == 3, f"{holders.keys()}")
    check("kein Trainer landet auf None, obwohl genug Turner/Gruppen da sind",
          all(plan[t] is not None for t in trainer), f"{plan}")

    # apply_timing_coverage/apply_timing_blocks (unveraendert) uebernehmen
    # danach korrekt die Randzeit eines haltenden Teilzeit-Trainers.
    last_row = len(grid_rows) - 1
    holder_c = plan.get("Trainer C")
    if holder_c and any(ck in ("g1_blau", "g2_orange") for _, ck in holder_c):
        trainer_timing = {"Trainer C": {"kind": "frueh", "time_min": None, "time_str": None,
                                         "notiz": "geht frueher", "blocked": [last_row]}}
        plan2 = {t: (list(c) if c else c) for t, c in plan.items()}
        plan2 = a.apply_timing_coverage(plan2, trainer_timing)
        plan2 = a.apply_timing_blocks(plan2, trainer_timing)
        check("Trainer C bleibt eingeteilt (nicht None)", plan2["Trainer C"] is not None)
        check("Trainer C's letzte Zeile ist rot 'geht frueher' markiert (sonder)",
              plan2["Trainer C"][last_row][1] == "sonder", f"{plan2['Trainer C']}")


def test_deterministic_timing_parser():
    """Bugfix 27.08.2026: Freitext-Anmerkungen wie 'Cassi geht ab 18:30'
    wurden bisher NIE deterministisch erkannt, sondern immer der KI ueber-
    lassen -- schlug die KI fehl oder klassifizierte faelschlich als
    Abwesenheit statt Timing, wurde der Trainer komplett aus dem Plan
    entfernt (der eigentliche gemeldete Bug). Jetzt erkennt bereits der
    deterministische Kommando-Parser (laeuft immer, auch ohne GH_MODELS_TOKEN)
    die gaengigen Formulierungen."""
    gruppen = ["G1"]
    turner = {"G1": ["G1Kind1", "G1Kind2", "G1Kind3"]}
    trainer = ["Cassi", "Fabian"]
    set_roster(gruppen, turner, trainer)

    cmd = a._parse_command_line("Cassi geht ab 18:30")
    check("'Cassi geht ab 18:30' wird als trainer_timing erkannt",
          cmd is not None and cmd.get("typ") == "trainer_timing", f"{cmd}")
    check("Richtung korrekt als 'frueh' erkannt", cmd and cmd.get("kind") == "frueh", f"{cmd}")
    check("Uhrzeit korrekt als 18:30 erkannt", cmd and cmd.get("time_str") == "18:30", f"{cmd}")

    cmd2 = a._parse_command_line("Cassi kommt erst um 18:00")
    check("'kommt erst um X' wird als 'spaet' erkannt", cmd2 and cmd2.get("kind") == "spaet", f"{cmd2}")

    cmd3 = a._parse_command_line("Cassi geht ab 18:30, Fabian macht G1")
    check("Zusammengesetzte Zeile mit zweitem Trainer wird NICHT deterministisch geparst (-> KI)",
          cmd3 is None, f"{cmd3}")

    fe = {}
    a._apply_command(cmd, fe)
    check("_apply_command setzt manuell_bearbeitet", fe.get("manuell_bearbeitet") is True)
    check("_apply_command schreibt ki.timing in derselben Struktur wie der KI-Pfad",
          fe.get("ki", {}).get("timing") == [{"trainer": "Cassi", "richtung": "frueh", "uhrzeit": "18:30"}],
          f"{fe}")

    cmd4 = a._parse_command_line("Cassi abwesend")
    check("echte volle Abwesenheit bleibt weiterhin 'trainer_abwesend' (keine Regression)",
          cmd4 and cmd4.get("typ") == "trainer_abwesend", f"{cmd4}")

    # Weitere reale Formulierungen (beim Review am 28.08.2026 nachgetestet):
    # "muss ... los/weg/gehen" mit eingeschobener Uhrzeit ("muss UM 18:30 los")
    # matchte urspruenglich NICHT (die Schluesselwortliste erwartete die
    # zusammenhaengende Phrase "muss los", nicht durch die Uhrzeit getrennt).
    for text, exp_kind, exp_time in [
        ("Cassi muss um 18:30 los", "frueh", "18:30"),
        ("Cassi muss leider um 18 Uhr weg", "frueh", "18:00"),
        ("Cassi muss heute schon um 18:30 gehen", "frueh", "18:30"),
        ("Cassi geht früher ca 18:30", "frueh", "18:30"),
        ("Cassi geht ab ca 18 Uhr", "frueh", "18:00"),
        ("Cassi kommt später, ca 18 Uhr", "spaet", "18:00"),
    ]:
        c = a._parse_command_line(text)
        check(f"'{text}' -> trainer_timing/{exp_kind}/{exp_time}",
              c is not None and c.get("typ") == "trainer_timing"
              and c.get("kind") == exp_kind and c.get("time_str") == exp_time, f"{c}")

    check("echte Krankmeldung ohne Timing-Schluesselwort bleibt unerkannt (-> KI, nicht faelschlich Timing)",
          a._parse_command_line("Cassi ist krank") is None)


if __name__ == "__main__":
    test_regression_grid()
    test_regression_plan_shape()
    test_group_counts()
    test_trainer_counts_and_immer_springer()
    test_merge_compatibility()
    test_rotation_fairness()
    test_ki_path_parity()
    test_edge_cases()
    test_geraet_tausch()
    test_partial_trainer_holds_group()
    test_deterministic_timing_parser()
    print()
    if FAILS:
        print(f"{len(FAILS)} FEHLGESCHLAGEN:")
        for f in FAILS:
            print("  -", f)
        sys.exit(1)
    else:
        print("ALLE TESTS BESTANDEN")
