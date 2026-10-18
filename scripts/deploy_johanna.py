#!/usr/bin/env python3
"""
Vertrauenswuerdiges Deploy-Skript fuer die Johanna-Website (/johanna).

WICHTIG (Sicherheit): Dieses Skript liegt bewusst HIER (tv-rheinzabern-auto) und NICHT im
Repo johanna-website. Johanna bzw. Claude Code duerfen dort alles aendern; trotzdem kommt
dadurch niemals fremder Code mit den Server-Zugangsdaten in Beruehrung. Aus dem Repo
johanna-website werden nur Dateien aus public/ als DATEN gelesen, nie ausgefuehrt.

Aufruf:  python scripts/deploy_johanna.py <pfad-zum-ausgecheckten-johanna-repo>
Umgebung: SSH_HOST, SSH_USER, SSH_PASSWORD, SSH_PORT (22), DRY_RUN (true/false)
Ziel ist fest /johanna.
"""

import os
import posixpath
import sys

TARGET = "/johanna"
MAX_FILE_BYTES = 10 * 1024 * 1024
MAX_FILES = 500
BLOCKED_NAMES = {".htaccess", ".htpasswd", ".env", "config.php", "secrets.php", ".deployed_sha"}
BLOCKED_SUFFIXES = (
    ".php", ".phtml", ".phar", ".php3", ".php4", ".php5", ".php7", ".pht", ".cgi", ".pl",
    ".py", ".sh", ".sqlite", ".db", ".pyc",
)


def collect(site_root: str):
    root = os.path.realpath(os.path.join(site_root, "public"))
    if not os.path.isdir(root):
        print("FEHLER: public/ fehlt.")
        sys.exit(1)
    files = []
    for dirpath, dirnames, filenames in os.walk(root, followlinks=False):
        dirnames[:] = sorted(d for d in dirnames if not d.startswith("."))
        for fname in sorted(filenames):
            local = os.path.join(dirpath, fname)
            rel = os.path.relpath(local, root).replace(os.sep, "/")
            name = fname.lower()
            if os.path.islink(local):
                print(f"WARNUNG: {rel} ist ein Link, uebersprungen.")
                continue
            if name in BLOCKED_NAMES or name.endswith(BLOCKED_SUFFIXES) or fname.startswith("."):
                print(f"WARNUNG: {rel} ist gesperrt, uebersprungen.")
                continue
            if os.path.getsize(local) > MAX_FILE_BYTES:
                print(f"WARNUNG: {rel} ist zu gross (>10 MB), uebersprungen.")
                continue
            if not os.path.realpath(local).startswith(root + os.sep):
                continue
            files.append((local, rel))
    if len(files) > MAX_FILES:
        print(f"FEHLER: mehr als {MAX_FILES} Dateien - Abbruch.")
        sys.exit(1)
    return files


def mkdir_p(sftp, remote_dir: str):
    current = ""
    for part in remote_dir.split("/"):
        if not part:
            continue
        current = f"{current}/{part}"
        try:
            sftp.stat(current)
        except FileNotFoundError:
            sftp.mkdir(current)


def main():
    if len(sys.argv) != 2:
        print("Aufruf: deploy_johanna.py <repo-pfad>")
        sys.exit(2)
    dry = os.environ.get("DRY_RUN", "false").strip().lower() in ("1", "true", "yes")
    files = collect(sys.argv[1])
    if not any(rel == "index.html" for _, rel in files):
        print("FEHLER: public/index.html fehlt - Abbruch.")
        sys.exit(1)
    print(f"=== {'TROCKENLAUF' if dry else 'DEPLOY'}: {len(files)} Datei(en) -> {TARGET} ===")
    for _, rel in files:
        print(f"  {rel}")
    if dry:
        return

    import paramiko

    t = paramiko.Transport((os.environ["SSH_HOST"], int(os.environ.get("SSH_PORT") or 22)))
    try:
        t.connect(username=os.environ["SSH_USER"], password=os.environ["SSH_PASSWORD"])
        sftp = paramiko.SFTPClient.from_transport(t)
        mkdir_p(sftp, TARGET)
        for local, rel in files:
            remote = posixpath.join(TARGET, rel)
            mkdir_p(sftp, posixpath.dirname(remote))
            sftp.put(local, remote)
            print(f"  uebertragen: {rel}")
        sftp.close()
    finally:
        t.close()
    print(f"=== Fertig: {len(files)} Datei(en) ===")


if __name__ == "__main__":
    main()
