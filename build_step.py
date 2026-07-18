#!/usr/bin/env python3
"""
build_step.py v2 — Staged video builder. (Stand 04.06.26)
Run repeatedly; each call renders ONE segment and saves progress.
Outputs exit 0 + "BUILD_COMPLETE" when both videos are fully rendered.

Changes vs v1:
  - Leicht: SW 1-5 Priorität | Schwer: SW 5-9 Priorität | Max 3 Overlap
  - Video unter Balken (setpts+pad+crop: y=40), unten 40px abgeschnitten
  - Letzten 5s jedes Fensters: Beep-Countdown-Ton (880 Hz, 0.15s/Beep)
  - Finale Videos in Fertig\ Subfolder

Usage:
  python3 build_step.py          -> next step
  python3 build_step.py reset    -> delete state, start fresh
"""
import subprocess, os, sys, shutil, tempfile, math, json, re
from datetime import date, timedelta

BASE    = os.path.dirname(os.path.abspath(__file__))
UPLOAD  = os.path.join(BASE, "Upload")
OUTPUT  = os.path.join(BASE, "Output")
FERTIG  = os.path.join(BASE, "Fertig")
VERLAUF = os.path.join(BASE, "verlauf.json")
STATE   = os.path.join(BASE, "wc_state.json")
FONT    = "/usr/share/fonts/truetype/google-fonts/Poppins-Bold.ttf"
DJ_FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"  # Pfeil-Symbole ▶◀●
os.makedirs(OUTPUT, exist_ok=True)
os.makedirs(FERTIG, exist_ok=True)

TYP_NAMEN = {'k': 'Kraft', 'd': 'Dehnen', 's': 'Sonstiges'}
DIR_LABEL = {'r': 'RECHTS  \u25B6', 'l': '\u25C0  LINKS', 'm': '\u25CF  MITTE'}
W, H, SR   = 1024, 576, 44100
ERKL_WIN   = 30
EXEC_SECS  = 60
TARGET_SEC = 600
MAX_SEC    = 670
OVERHEAD   = 8
MAX_OVERLAP = 3

# Nummern, die NIE zusammen in einer Wochenchallenge vorkommen duerfen (Noah 10.07.26)
INKOMPATIBEL = {frozenset((11, 29))}
def _nrint(x):
    try: return int(str(x))
    except Exception: return None
def _drop_inkompatibel(alle, lu):
    """Bei einem verbotenen Paar bleibt nur die LAENGER nicht verwendete Nummer (LRU); die andere faellt raus."""
    alle = list(alle)
    for pair in INKOMPATIBEL:
        present = [u for u in alle if _nrint(u.get("nr")) in pair]
        if len(present) >= 2:
            present.sort(key=lambda u: lu.get(u.get("nr"), "2000-01-01"), reverse=True)
            drop = present[0]
            alle = [u for u in alle if u is not drop]
            print(f"[inkompatibel] Nr {drop.get('nr')} entfernt (Paar {tuple(pair)})")
    return alle

# ⭐ Video wird 40px nach unten verschoben, unten 40px abgeschnitten
SCALE = (f"scale={W}:{H}:force_original_aspect_ratio=decrease,"
         f"pad={W}:{H}:(ow-iw)/2:(oh-ih)/2:black,"
         f"pad={W}:{H+40}:0:40:black,"
         f"crop={W}:{H}:0:0")

BEEP_DUR   = 0.15   # Sekunden pro Beep
BEEP_FREQ  = 880    # Hz
BEEP_VOL   = 0.08


GERAETE_FORMEN = {
    'flasche': ('Flasche', 'Flaschen'),
    'flaschen': ('Flasche', 'Flaschen'),
    'stuhl': ('Stuhl', 'Stühle'),
    'stühle': ('Stuhl', 'Stühle'),
    'stuehle': ('Stuhl', 'Stühle'),
    'stab': ('Stab', 'Stäbe'),
    'stäbe': ('Stab', 'Stäbe'),
    'treppe': ('Treppe', 'Treppen'),
    'treppen': ('Treppe', 'Treppen'),
    'ball': ('Ball', 'Bälle'),
    'bälle': ('Ball', 'Bälle'),
    'kissen': ('Kissen', 'Kissen'),
}


def _parse_material(s):
    if not s:
        return None
    s = s.strip()
    if s in ('', '_'):
        return None
    m = re.match(r'^(\d+)\s*\.?\s*(.+)$', s)
    if m:
        return int(m.group(1)), m.group(2).strip()
    return 1, s


def dedup_materialien(mats):
    """Gleiche Gegenstände zusammenfassen -> höchste Menge, saubere Schreibweise."""
    groups = {}
    for s in mats:
        pm = _parse_material(s)
        if not pm:
            continue
        qty, noun = pm
        forms = GERAETE_FORMEN.get(noun.lower())
        if forms:
            sing, plur = forms
            gkey = sing.lower()
        else:
            sing = plur = noun
            gkey = noun.lower()
        if gkey not in groups or qty > groups[gkey][0]:
            groups[gkey] = (qty, sing, plur)
    out = [f"{qty} {plur if qty != 1 else sing}"
           for qty, sing, plur in groups.values()]
    return sorted(out)


def beep_expr_multi(boundaries, win_dur):
    """Beep in den letzten 5s VOR jedem Zeitpunkt in 'boundaries'. Kommas escaped."""
    gates = []
    for b in boundaries:
        bs = b - 5
        gates.append(fr'(gte(t\,{bs})*lt(t\,{b}))')
    gate = '(' + '+'.join(gates) + ')'
    return (f'{BEEP_VOL}*sin({BEEP_FREQ}*2*PI*t)'
            fr'*lt(mod(t\,1)\,{BEEP_DUR})'
            f'*{gate}')


def _middle_frame(src, out_png):
    """Mittleres Einzelbild des Clips als PNG (out_png liegt in lokalem /tmp)."""
    d = get_dur(src)
    ts = max(0.0, d / 2.0)
    run(['ffmpeg', '-y', '-ss', f'{ts:.3f}', '-i', src,
         '-frames:v', '1', out_png])


def beep_expr(win_dur):
    """FFmpeg aevalsrc expression: beep in last 5s. Commas escaped for lavfi URL."""
    bs = win_dur - 5
    return (f'{BEEP_VOL}*sin({BEEP_FREQ}*2*PI*t)'
            fr'*lt(mod(t-{bs}\,1)\,{BEEP_DUR})'
            fr'*gte(t\,{bs})*lt(t\,{win_dur})')


def assign_videos(sw):
    sw = int(sw)
    if sw <= 2:   return ['leicht']
    elif sw <= 7: return ['leicht', 'schwer']
    else:          return ['schwer']

def run(cmd):
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        print("FFMPEG ERR:", r.stderr[-600:])
        raise RuntimeError(f"ffmpeg failed: {cmd[:3]}")
    return r

def get_dur(path):
    r = subprocess.run(
        ['ffprobe', '-v', 'error', '-show_entries', 'format=duration',
         '-of', 'default=noprint_wrappers=1:nokey=1', path],
        capture_output=True, text=True)
    try:   return float(r.stdout.strip())
    except: return 0.0

def vargs():
    return ['-r', '30', '-c:v', 'libx264', '-preset', 'ultrafast',
            '-pix_fmt', 'yuv420p', '-c:a', 'aac', '-ar', str(SR), '-ac', '1']

def write_tmp(tmp_path, dest):
    """Copy from /tmp to CIFS (CIFS blocks direct ffmpeg writes)."""
    shutil.copy2(tmp_path, dest)
    os.remove(tmp_path)


# ─── renderers ────────────────────────────────────────────────────────────────

def render_intro(level_label, materialien, zeitraum, out):
    mat_list = ', '.join(dedup_materialien(materialien))
    mat = f"Material\\: {mat_list}" if mat_list else 'Kein Material nötig!'
    dur = 4.0
    vf  = (f"drawtext=fontfile={FONT}:fontsize=52:fontcolor=white"
           f":x=(w-tw)/2:y=(h-th)/2-60:text='Wochenchallenge {level_label}'"
           f":shadowcolor=black@0.6:shadowx=2:shadowy=2,"
           f"drawtext=fontfile={FONT}:fontsize=26:fontcolor=white@0.85"
           f":x=(w-tw)/2:y=(h-th)/2+10:text='{mat}',"
           f"drawtext=fontfile={FONT}:fontsize=26:fontcolor=0xFFEB3B"
           f":x=(w-tw)/2:y=(h-th)/2+52:text='{zeitraum}'")
    tmp = f"/tmp/wc_intro_{os.getpid()}.mp4"
    run(['ffmpeg', '-y',
         '-f', 'lavfi', '-i', f"color=0x1A237E:s={W}x{H}:d={dur}",
         '-f', 'lavfi', '-i', f'aevalsrc=0:c=mono:s={SR}:d={dur}',
         '-filter_complex', f'[0:v]{vf}[v]',
         '-map', '[v]', '-map', '1:a',
         *vargs(), '-shortest', tmp])
    write_tmp(tmp, out)


def render_erkl(src, out):
    """30s explanation window with beep in last 5s."""
    dur = get_dur(src)
    tmpd = tempfile.mkdtemp(prefix='wc_erkl_')
    try:
        erkl_out = f"{tmpd}/erkl.mp4"

        # Video filter chain (bar + counter, video shifted below bar)
        vf_chain = (
            f"setpts=PTS-STARTPTS,{SCALE},"
            f"drawbox=x=0:y=0:w=iw:h=40:color=0x1A237E:t=fill,"
            f"drawtext=fontfile={FONT}:fontsize=28:fontcolor=white"
            f":x=(w-tw)/2:y=6:text='Erklärung',"
            f"drawbox=x=w-140:y=h-110:w=135:h=100:color=black@0.65:t=fill,"
            f"drawtext=fontfile={FONT}:fontsize=72:fontcolor=0xFFEB3B"
            f":x=w-132:y=h-108"
            f":text='%{{eif\\:ceil({ERKL_WIN}-t)\\:d}}'"
        )

        clip_dur = min(dur, float(ERKL_WIN))

        if dur >= ERKL_WIN:
            # Full 30s from video — beep at t=25-30, mix with original audio
            bexpr = beep_expr(ERKL_WIN)
            run(['ffmpeg', '-y', '-i', src,
                 '-f', 'lavfi', '-i',
                 f'aevalsrc={bexpr}:c=mono:s={SR}:d={ERKL_WIN}',
                 '-t', str(ERKL_WIN),
                 '-filter_complex',
                 f'[0:v]{vf_chain}[v];'
                 f'[0:a]aformat=sample_rates={SR}:channel_layouts=mono,loudnorm=I=-18:LRA=7:TP=-2,volume=4.0[oa];'
                 f'[oa][1:a]amix=inputs=2:duration=first:weights=1 1[a]',
                 '-map', '[v]', '-map', '[a]',
                 *vargs(), erkl_out])
            write_tmp(erkl_out, out)
        else:
            # Short video: erkl_base + gleich screen
            bs_erkl = ERKL_WIN - 5  # 25 — beep start in assembled time
            if dur > bs_erkl:
                # Beep appears in erkl clip too
                bexpr_e = beep_expr_clip(bs_erkl, clip_dur)
                run(['ffmpeg', '-y', '-i', src,
                     '-f', 'lavfi', '-i',
                     f'aevalsrc={bexpr_e}:c=mono:s={SR}:d={clip_dur}',
                     '-filter_complex',
                     f'[0:v]{vf_chain}[v];'
                     f'[0:a]aformat=sample_rates={SR}:channel_layouts=mono,loudnorm=I=-18:LRA=7:TP=-2,volume=4.0[oa];'
                     f'[oa][1:a]amix=inputs=2:duration=first:weights=1 1[a]',
                     '-map', '[v]', '-map', '[a]',
                     *vargs(), erkl_out])
            else:
                # No beep in erkl clip (beep is entirely in gleich screen)
                run(['ffmpeg', '-y', '-i', src,
                     '-filter_complex',
                     f'[0:v]{vf_chain}[v];[0:a]loudnorm=I=-18:LRA=7:TP=-2,volume=4.0[a]',
                     '-map', '[v]', '-map', '[a]',
                     *vargs(), erkl_out])

            # "Gleich geht's los" screen
            rest = math.ceil(ERKL_WIN - dur)
            cd_start = math.ceil(ERKL_WIN - dur)
            gleich = f"{tmpd}/gleich.mp4"

            # Beep in gleich: assembled t=25-30 → local t=(25-dur) to (30-dur)
            gleich_bs = max(0.0, bs_erkl - dur)
            bexpr_g = beep_expr_clip(gleich_bs, float(rest))

            vfg = (
                f"drawtext=fontfile={FONT}:fontsize=44:fontcolor=white"
                f":x=(w-tw)/2:y=(h-th)/2-80:text='Gleich geht es los...',"
                f"drawtext=fontfile={FONT}:fontsize=130:fontcolor=0xFFEB3B"
                f":x=(w-tw)/2:y=(h-th)/2"
                f":text='%{{eif\\:ceil({cd_start}-t)\\:d}}'"
                f":shadowcolor=black@0.7:shadowx=3:shadowy=3"
            )
            run(['ffmpeg', '-y',
                 '-f', 'lavfi', '-i', f"color=0x1565C0:s={W}x{H}:d={rest}",
                 '-f', 'lavfi', '-i', f'aevalsrc={bexpr_g}:c=mono:s={SR}:d={rest}',
                 '-filter_complex', f'[0:v]{vfg}[v]',
                 '-map', '[v]', '-map', '1:a',
                 *vargs(), '-shortest', gleich])

            cf = f"{tmpd}/c.txt"
            with open(cf, 'w') as f:
                f.write(f"file '{erkl_out}'\nfile '{gleich}'\n")
            tmp_concat = f'/tmp/wc_erkl_concat_{os.getpid()}.mp4'
            run(['ffmpeg', '-y', '-f', 'concat', '-safe', '0', '-i', cf, '-c', 'copy', tmp_concat])
            write_tmp(tmp_concat, out)
    finally:
        shutil.rmtree(tmpd, ignore_errors=True)


def beep_expr_clip(beep_start, clip_dur):
    """Beep expression for a clip. Commas escaped for lavfi URL."""
    return (f'{BEEP_VOL}*sin({BEEP_FREQ}*2*PI*t)'
            fr'*lt(mod(t-{beep_start}\,1)\,{BEEP_DUR})'
            fr'*gte(t\,{beep_start})*lt(t\,{clip_dur})')


def _render_durch_normal(src, out):
    """60s Loop: stumm ausser Beep in letzten 5s. (unveraenderte Original-Logik)"""
    vf = (
        f"setpts=PTS-STARTPTS,{SCALE},"
        f"drawbox=x=0:y=0:w=iw:h=40:color=0xC62828:t=fill,"
        f"drawtext=fontfile={FONT}:fontsize=28:fontcolor=white"
        f":x=(w-tw)/2:y=6:text='Jetzt mitmachen',"
        f"drawbox=x=w-140:y=h-110:w=135:h=100:color=black@0.65:t=fill,"
        f"drawtext=fontfile={FONT}:fontsize=72:fontcolor=0xFFEB3B"
        f":x=w-132:y=h-108"
        f":text='%{{eif\\:ceil({EXEC_SECS}-t)\\:d}}'"
    )
    bexpr = beep_expr(EXEC_SECS)
    tmp = f"/tmp/wc_durch_{os.getpid()}.mp4"
    run(['ffmpeg', '-y',
         '-stream_loop', '20', '-i', src, '-t', str(EXEC_SECS),
         '-f', 'lavfi', '-i', f'aevalsrc={bexpr}:c=mono:s={SR}',
         '-vf', vf,
         '-map', '0:v', '-map', '1:a',
         *vargs(), '-t', str(EXEC_SECS), tmp])
    write_tmp(tmp, out)


def render_durch(seg, out):
    """
    Durchfuehrung (60s). Unterstuetzt:
      - Normal: Video 60s Loop (unveraendert)
      - Standbild (s): eingefrorenes mittleres Bild statt Loop
      - Richtungen (r/l/m): 2 -> je 30s (r,l), 3 -> je 20s (r,l,m); Timer durchgehend
        60->0, Beep vor JEDEM Wechsel, Richtungs-Label (RECHTS/LINKS/MITTE) eingeblendet.
    """
    standbild  = seg.get('standbild', False)
    richtungen = seg.get('richtungen')   # [[dir, path], ...] R,L,M  oder None

    if not standbild and not richtungen:
        return _render_durch_normal(seg['src'], out)

    if richtungen:
        sources = [p for (_d, p) in richtungen]
        n = len(sources)
        seg_dur = EXEC_SECS // n            # 2->30, 3->20
    else:
        sources = [seg['src']]
        n = 1
        seg_dur = EXEC_SECS

    boundaries = [seg_dur * (k + 1) for k in range(n)]

    tmpd = tempfile.mkdtemp(prefix='wc_durch_')
    try:
        prepped = []
        for i, p in enumerate(sources):
            if standbild:
                png = os.path.join(tmpd, f"frame{i}.png")
                _middle_frame(p, png)
                prepped.append((['-loop', '1', '-t', str(seg_dur)], png))
            else:
                prepped.append((['-stream_loop', '120', '-t', str(seg_dur)], p))

        cmd = ['ffmpeg', '-y']
        for pre, path in prepped:
            cmd += pre + ['-i', path]
        bexpr = beep_expr_multi(boundaries, EXEC_SECS)
        cmd += ['-f', 'lavfi', '-i',
                f'aevalsrc={bexpr}:c=mono:s={SR}:d={EXEC_SECS}']

        fc = []
        labels = ''
        for i in range(len(prepped)):
            fc.append(f"[{i}:v]setpts=PTS-STARTPTS,{SCALE},fps=30,"
                      f"trim=duration={seg_dur},setpts=PTS-STARTPTS[s{i}]")
            labels += f"[s{i}]"
        if len(prepped) > 1:
            fc.append(f"{labels}concat=n={len(prepped)}:v=1:a=0[cat]")
            base = "[cat]"
        else:
            base = "[s0]"

        chain = [
            f"drawbox=x=0:y=0:w=iw:h=40:color=0xC62828:t=fill",
            f"drawtext=fontfile={FONT}:fontsize=28:fontcolor=white"
            f":x=(w-tw)/2:y=6:text='Jetzt mitmachen'",
            f"drawbox=x=w-140:y=h-110:w=135:h=100:color=black@0.65:t=fill",
            f"drawtext=fontfile={FONT}:fontsize=72:fontcolor=0xFFEB3B"
            f":x=w-132:y=h-108"
            f":text='%{{eif\\:ceil({EXEC_SECS}-t)\\:d}}'",
        ]
        # Richtungs-Label pro Segment (nur bei Richtungen)
        if richtungen:
            for idx, (d, _p) in enumerate(richtungen):
                s0 = seg_dur * idx
                e0 = seg_dur * (idx + 1)
                txt = DIR_LABEL.get(d, d.upper())
                chain.append(
                    f"drawtext=fontfile={DJ_FONT}:fontsize=34:fontcolor=white"
                    f":x=(w-tw)/2:y=50:box=1:boxcolor=black@0.55:boxborderw=10"
                    f":text='{txt}':enable='between(t\\,{s0}\\,{e0})'"
                )
        fc.append(f"{base}{','.join(chain)}[v]")

        audio_idx = len(prepped)
        tmp = f"/tmp/wc_durch_{os.getpid()}.mp4"
        cmd += ['-filter_complex', ';'.join(fc),
                '-map', '[v]', '-map', f'{audio_idx}:a',
                *vargs(), '-t', str(EXEC_SECS), tmp]
        run(cmd)
        write_tmp(tmp, out)
    finally:
        shutil.rmtree(tmpd, ignore_errors=True)


def render_outro(out):
    dur = 4.0
    vf  = (
        f"drawtext=fontfile={FONT}:fontsize=56:fontcolor=white"
        f":x=(w-tw)/2:y=(h-th)/2-30"
        f":text='Fertig\\! Gut gemacht\\!'"
        f":shadowcolor=black@0.6:shadowx=2:shadowy=2,"
        f"drawtext=fontfile={FONT}:fontsize=26:fontcolor=0xFFEB3B"
        f":x=(w-tw)/2:y=(h-th)/2+35"
        f":text='Wochenchallenge geschafft\\!'"
    )
    tmp = f"/tmp/wc_outro_{os.getpid()}.mp4"
    run(['ffmpeg', '-y',
         '-f', 'lavfi', '-i', f"color=0x1A237E:s={W}x{H}:d={dur}",
         '-f', 'lavfi', '-i', f'aevalsrc=0:c=mono:s={SR}:d={dur}',
         '-filter_complex', f'[0:v]{vf}[v]',
         '-map', '[v]', '-map', '1:a',
         *vargs(), '-shortest', tmp])
    write_tmp(tmp, out)


# ─── scan + select ────────────────────────────────────────────────────────────

def _parse_name(base):
    """Zerlegt Dateinamen -> Metadaten. Reihenfolge: Nr,Bereich,SW,Geraet,[s],[r/l/m].
       Leerfeld=_ ; s=Standbild (nach Geraet) ODER Bereich 'Sonstiges' (Pos.2 + SW-Zahl danach)."""
    parts = [p.strip() for p in base.split(',')]
    while parts and parts[-1] == '':
        parts.pop()
    if not parts or not parts[0].isdigit():
        return None
    info = {'nr': parts[0], 'erkl': False, 'typ': None, 'sw': None,
            'material': None, 'standbild': False, 'richtung': None}
    toks = parts[1:]
    if not toks:
        info['erkl'] = True
        return info
    if toks[-1].lower() in ('r', 'l', 'm'):
        info['richtung'] = toks[-1].lower()
        toks = toks[:-1]
    if len(toks) >= 2 and toks[0] in ('k', 'd', 's') and toks[1].isdigit():
        info['typ'] = toks[0]
        info['sw'] = int(toks[1])
        rest = toks[2:]
        if rest and rest[-1] == 's':
            info['standbild'] = True
            rest = rest[:-1]
        if rest:
            g = rest[0].strip()
            info['material'] = None if g in ('', '_') else g
    else:
        if toks and toks[-1] == 's':
            info['standbild'] = True
            toks = toks[:-1]
    return info


def scan_upload():
    RICHT_ORDER = {'r': 0, 'l': 1, 'm': 2}
    raw = {}
    for f in sorted(os.listdir(UPLOAD)):
        if not (f.lower().endswith('.mp4') or f.lower().endswith('.mov')):
            continue
        info = _parse_name(f[:-4])
        if not info:
            continue
        nr = info['nr']
        u = raw.setdefault(nr, {'nr': nr, 'typ': None, 'schwierigkeit': None,
                                'material': None, 'erkl': None, 'durch': None,
                                'standbild': False, '_richt': {}})
        path = os.path.join(UPLOAD, f)
        if info['erkl']:
            u['erkl'] = path
            continue
        if info['typ'] in ('k', 'd', 's') and info['sw'] and 1 <= info['sw'] <= 9:
            u['typ'] = info['typ']
            u['schwierigkeit'] = info['sw']
            if info['material'] is not None:
                u['material'] = info['material']
        if info['standbild']:
            u['standbild'] = True
        if info['richtung']:
            u['_richt'][info['richtung']] = path
        else:
            u['durch'] = path
    result = []
    for u in raw.values():
        if u['_richt']:
            u['richtungen'] = [[d, u['_richt'][d]]
                               for d in sorted(u['_richt'], key=lambda x: RICHT_ORDER.get(x, 9))]
        else:
            u['richtungen'] = None
        del u['_richt']
        if u['erkl'] and u['typ'] and (u['durch'] or u['richtungen']):
            result.append(u)
    return result


def _fill_typmix(pool, n):
    """Waehlt n Uebungen mit Zielverteilung ~50% Kraft / 30% Dehnen / 20% Sonstiges
    (soweit der Pool es hergibt) und Typ-Abwechslung."""
    pool = list(pool)
    result = []
    last_typ = None
    tgt = {"k": round(n * 0.5), "d": round(n * 0.3)}
    tgt["s"] = max(0, n - tgt["k"] - tgt["d"])
    def cnt(t):
        return sum(1 for u in result if u.get("typ") == t)
    def avail(t):
        return any(u.get("typ") == t for u in pool)
    while len(result) < n and pool:
        under = [t for t in ("k", "d", "s") if cnt(t) < tgt.get(t, 0) and avail(t)]
        cand = [u for u in pool if u.get("typ") in under] if under else list(pool)
        diff = [u for u in cand if u.get("typ") != last_typ]
        pick = diff if diff else cand
        ub = pick[0]
        pool.remove(ub)
        result.append(ub)
        last_typ = ub.get("typ")
    return result


def select_exercises_pair(alle, verlauf):
    """
    Select exercises for Leicht AND Schwer together.
    Leicht: SW 1-5 priority (ascending). Schwer: SW 5-9 priority (descending).
    Max MAX_OVERLAP exercises shared between the two videos.
    """
    lu = verlauf.get('letzte_verwendung', {})
    alle = _drop_inkompatibel(alle, lu)

    # How many exercises per video
    per_ex = ERKL_WIN + EXEC_SECS
    n, total = 0, 0
    while total + per_ex <= MAX_SEC - OVERHEAD:
        total += per_ex
        n += 1
        if total >= TARGET_SEC - OVERHEAD:
            break

    leicht_pool = [u for u in alle if 'leicht' in assign_videos(u['schwierigkeit'])]
    schwer_pool  = [u for u in alle if 'schwer' in assign_videos(u['schwierigkeit'])]

    # Sort: Leicht = low SW first, then LRU
    leicht_pool.sort(key=lambda u: (lu.get(u['nr'], '2000-01-01'), u['schwierigkeit']))
    # Sort: Schwer = high SW first, then LRU
    schwer_pool.sort(key=lambda u: (lu.get(u['nr'], '2000-01-01'), -u['schwierigkeit']))

    leicht_sel = _fill_typmix(leicht_pool, n)
    schwer_sel  = _fill_typmix(schwer_pool, n)

    # Enforce max overlap
    l_nrs = {u['nr'] for u in leicht_sel}
    s_nrs = {u['nr'] for u in schwer_sel}
    overlap = l_nrs & s_nrs

    if len(overlap) > MAX_OVERLAP:
        # Replace lowest-SW overlapping exercises in Schwer with Schwer-exclusive ones
        schwer_excl = [u for u in schwer_pool if u['nr'] not in l_nrs]
        # Overlapping in Schwer, sorted by SW ascending (least schwer first → replace those)
        overlap_in_s = sorted(
            [u for u in schwer_sel if u['nr'] in overlap],
            key=lambda u: u['schwierigkeit']
        )
        new_schwer = list(schwer_sel)
        used_nrs = {u['nr'] for u in new_schwer}
        for to_rm in overlap_in_s:
            if len(overlap) <= MAX_OVERLAP:
                break
            repl = next((u for u in schwer_excl if u['nr'] not in used_nrs), None)
            if repl:
                idx = next(i for i, u in enumerate(new_schwer) if u['nr'] == to_rm['nr'])
                new_schwer[idx] = repl
                used_nrs.discard(to_rm['nr'])
                used_nrs.add(repl['nr'])
                overlap.discard(to_rm['nr'])
        schwer_sel = new_schwer

    print(f"  Leicht ({n}): {[u['nr'] for u in leicht_sel]}")
    print(f"  Schwer ({n}): {[u['nr'] for u in schwer_sel]}")
    overlap_final = {u['nr'] for u in leicht_sel} & {u['nr'] for u in schwer_sel}
    print(f"  Overlap: {sorted(overlap_final)} ({len(overlap_final)})")
    return leicht_sel, schwer_sel


# ─── state helpers ────────────────────────────────────────────────────────────

def load_state():
    with open(STATE, encoding='utf-8') as f:
        return json.load(f)

def save_state(s):
    # Write to /tmp first, then copy to CIFS
    tmp = f"/tmp/wc_state_{os.getpid()}.json"
    with open(tmp, 'w') as f:
        json.dump(s, f, ensure_ascii=False, indent=2)
    shutil.copy2(tmp, STATE)
    os.remove(tmp)

def load_verlauf():
    try:
        with open(VERLAUF, encoding='utf-8') as f:
            return json.load(f)
    except:
        return {'letzte_verwendung': {}, 'videos': []}

def save_verlauf(verlauf, used_nrs, results, sam_str, zeitraum):
    heute = date.today().isoformat()
    for nr in used_nrs:
        verlauf['letzte_verwendung'][nr] = heute
    verlauf['videos'].append({
        'datum': heute, 'samstag': sam_str, 'zeitraum': zeitraum,
        'videos': [os.path.basename(r) for r in results],
        'uebungen': sorted(used_nrs),
    })
    tmp = f"/tmp/wc_verlauf_{os.getpid()}.json"
    with open(tmp, 'w') as f:
        json.dump(verlauf, f, ensure_ascii=False, indent=2)
    shutil.copy2(tmp, VERLAUF)
    os.remove(tmp)


# ─── init ─────────────────────────────────────────────────────────────────────

def init_state():
    alle    = scan_upload()
    verlauf = load_verlauf()
    _days = (5 - date.today().weekday()) % 7
    if _days == 0: _days = 7  # Heute ist Samstag -> naechste Woche
    sam     = date.today() + timedelta(days=_days)
    di      = sam + timedelta(days=3)
    sam_str = sam.strftime('%d.%m.%y')
    zeitraum = f"{sam.day}.{sam.month}. - {di.day}.{di.month}."

    leicht_ueb, schwer_ueb = select_exercises_pair(alle, verlauf)

    # Wochen-Unterordner in Fertig\ (z.B. Fertig\ab 13.06.26\)
    week_dir = os.path.join(FERTIG, f"ab {sam_str}")
    os.makedirs(week_dir, exist_ok=True)

    state = {'sam_str': sam_str, 'zeitraum': zeitraum, 'videos': {}}

    for vtyp, label, uebungen in [('leicht', 'Leicht', leicht_ueb),
                                    ('schwer',  'Schwer',  schwer_ueb)]:
        mat  = [u.get('material') for u in uebungen]
        segs = []
        seg_base = os.path.join(OUTPUT, f"_seg_{vtyp}_")
        segs.append({'name': 'intro', 'done': False,
                     'out': seg_base + 'intro.mp4'})
        for i, ub in enumerate(uebungen):
            segs.append({'name': f'erkl_{i}', 'done': False,
                         'src': ub['erkl'], 'out': seg_base + f'erkl{i}.mp4'})
            segs.append({'name': f'durch_{i}', 'done': False,
                         'src': ub.get('durch'),
                         'standbild': ub.get('standbild', False),
                         'richtungen': ub.get('richtungen'),
                         'out': seg_base + f'durch{i}.mp4'})
        segs.append({'name': 'outro', 'done': False,
                     'out': seg_base + 'outro.mp4'})

        state['videos'][vtyp] = {
            'label': label,
            'materialien': mat,
            'uebungen_nrs': [u['nr'] for u in uebungen],
            'segments': segs,
            'concat_done': False,
            'output': os.path.join(week_dir, f"Wochenchallenge ab {sam_str} {vtyp}.mp4"),
        }

    save_state(state)
    total_segs = sum(len(v['segments']) for v in state['videos'].values())
    print(f"PLAN: {sam_str} | {total_segs} Segmente total")


# ─── main step logic ──────────────────────────────────────────────────────────

def step():
    if not os.path.exists(STATE):
        init_state()
        return False

    state = load_state()
    if 'videos' not in state:  # empty/reset state -> re-init
        init_state()
        return False

    for vtyp in ['leicht', 'schwer']:
        v = state['videos'][vtyp]

        # Render next pending segment
        for seg in v['segments']:
            if seg['done']:
                continue
            name = seg['name']
            out  = seg['out']
            print(f"  [{vtyp}] {name}")

            if name == 'intro':
                render_intro(v['label'], v['materialien'], state['zeitraum'], out)
            elif name.startswith('erkl_'):
                render_erkl(seg['src'], out)
            elif name.startswith('durch_'):
                render_durch(seg, out)
            elif name == 'outro':
                render_outro(out)

            seg['done'] = True
            save_state(state)
            return False

        # All segments done — concat
        if not v['concat_done']:
            segs_done = [seg['out'] for seg in v['segments']]
            cf = os.path.join(OUTPUT, f"_concat_{vtyp}.txt")
            with open(cf, 'w') as f:
                for p in segs_done:
                    f.write(f"file '{p}'\n")
            tmp_out = f"/tmp/wc_final_{vtyp}_{os.getpid()}.mp4"
            run(['ffmpeg', '-y', '-f', 'concat', '-safe', '0',
                 '-i', cf, '-c', 'copy', tmp_out])
            write_tmp(tmp_out, v['output'])
            dur = get_dur(v['output'])
            print(f"  CONCAT {vtyp}: {os.path.basename(v['output'])} ({dur/60:.1f} min)")
            v['concat_done'] = True
            save_state(state)
            # Cleanup concat file (ignore CIFS delete errors)
            try: os.remove(cf)
            except: pass
            return False

    # Both done — finalize
    verlauf  = load_verlauf()
    used_nrs = set()
    results  = []
    for vtyp in ['leicht', 'schwer']:
        v = state['videos'][vtyp]
        results.append(v['output'])
        used_nrs.update(v['uebungen_nrs'])
    save_verlauf(verlauf, used_nrs, results, state['sam_str'], state['zeitraum'])

    # Upload bat
    claude_root = os.path.dirname(BASE)
    bat = os.path.join(claude_root, "upload_now.bat")
    lines = ["@echo off", "chcp 65001 >nul 2>&1",
             "echo ============================================",
             "echo  Wochenchallenge Videos hochladen",
             "echo ============================================", "echo."]
    for i, out in enumerate(results, 1):
        win = out.replace(claude_root, "C:\\Claude").replace("/", "\\")
        fname = os.path.basename(out)
        lines += [
            f"echo [{i}/{len(results)}] Lade hoch: {fname}",
            f'python "C:\\Claude\\upload.py" "{win}" wochen-challenge-videos'
            f' >> "C:\\Claude\\upload_result.txt" 2>&1',
            'if errorlevel 1 (',
            '    echo FEHLER! Details: C:\\Claude\\upload_result.txt',
            '    pause', '    exit /b 1', ')', "echo.",
        ]
    lines += ["echo ============================================",
              "echo  Alle Videos erfolgreich hochgeladen!",
              "echo ============================================", "timeout /t 3"]
    tmp_bat = f"/tmp/wc_upload_{os.getpid()}.bat"
    with open(tmp_bat, "w", newline="\r\n", encoding="utf-8") as f:
        f.write("\r\n".join(lines) + "\r\n")
    shutil.copy2(tmp_bat, bat)
    os.remove(tmp_bat)

    # Try to remove state (may fail on CIFS, that's OK)
    try: os.remove(STATE)
    except: pass

    print("BUILD_COMPLETE")
    for r in results:
        d = get_dur(r)
        print(f"  {os.path.basename(r)}  ({d/60:.1f} min)")
    return True


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'reset':
        for f in [STATE]:
            try: os.remove(f)
            except: pass
        for fn in os.listdir(OUTPUT):
                try: os.remove(os.path.join(OUTPUT, fn))
                except: pass
        print("State reset.")
        sys.exit(0)

    done = step()
    sys.exit(0 if done else 1)
