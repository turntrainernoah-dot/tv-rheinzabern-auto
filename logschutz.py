"""Kompakte Log-Ausgabe fuer die Actions-Laeufe.

Die Workflow-Logs dieses Repos sind oeffentlich einsehbar. Nach dem Laden der
Personenliste vom Server kuerzt dieses Modul Personennamen in allem, was auf
stdout/stderr landet, auf Initialen ("Vorname N." -> "V.N."). Dateien,
Server-Uploads, Meldungen (melde_problem) und Rueckgabewerte bleiben
unveraendert - es betrifft ausschliesslich die Konsolenausgabe.

Benutzung (nach dem Roster-Laden):
    import logschutz
    logschutz.namen_setzen(liste_von_namen)
"""

import re
import sys

_BUCHST = "A-Za-zÄÖÜäöüßÀ-ÿ"
_muster = None
_ersatz = {}


def _initialen(name):
    teile = re.findall(f"[{_BUCHST}]+", name)
    return "".join(t[0].upper() + "." for t in teile) or "?"


def kuerzen(text):
    """Ersetzt bekannte Namen in `text` durch Initialen (sonst unveraendert)."""
    if _muster is None or not isinstance(text, str) or not text:
        return text
    try:
        return _muster.sub(lambda m: _ersatz.get(m.group(0).lower(), "?"), text)
    except Exception:
        return text


class _Filter:
    """Duenner Wrapper um einen Text-Stream; nur write()/writelines() filtern."""

    def __init__(self, stream):
        self._stream = stream

    def write(self, s):
        return self._stream.write(kuerzen(s))

    def writelines(self, lines):
        for line in lines:
            self.write(line)

    def __getattr__(self, name):
        return getattr(self._stream, name)


def namen_setzen(namen):
    """Registriert die zu kuerzenden Namen und aktiviert den Filter.
    Mehrfacher Aufruf ergaenzt die Liste. Wirft nie eine Exception."""
    global _muster
    try:
        for n in namen or ():
            if not isinstance(n, str):
                continue
            n = re.sub(r"\s*\(G\d+\)\s*$", "", n).strip()
            if len(n) < 3:
                continue
            _ersatz.setdefault(n.lower(), _initialen(n))
            for teil in re.findall(f"[{_BUCHST}]{{3,}}", n):
                _ersatz.setdefault(teil.lower(), teil[0].upper() + ".")
        if not _ersatz:
            return
        schluessel = sorted(_ersatz, key=len, reverse=True)
        _muster = re.compile(
            f"(?<![{_BUCHST}0-9_])(?:" + "|".join(re.escape(k) for k in schluessel)
            + f")(?![{_BUCHST}0-9_])",
            re.IGNORECASE,
        )
        if not isinstance(sys.stdout, _Filter):
            sys.stdout = _Filter(sys.stdout)
        if not isinstance(sys.stderr, _Filter):
            sys.stderr = _Filter(sys.stderr)
    except Exception:
        pass
