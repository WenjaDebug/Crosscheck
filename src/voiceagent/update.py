"""Selbst aktualisieren: Vor jedem Start holt Gegenprobe die neueste Version von GitHub.

Läuft aus VoiceAgent.command und VoiceAgent.bat, bevor der Server startet:
1. Ist der Ordner ein Git-Klon auf dem Zweig main ohne eigene Änderungen, holt es die
   neuen Änderungen (git pull --ff-only). Offline oder ohne Anmeldung startet einfach die
   bisherige Version. Auf einem anderen Zweig (zum Testen eines PRs) bleibt alles, wie es ist.
2. Haben sich die Pakete geändert (pyproject.toml), installiert es sie nach.

Wer den Ordner als ZIP von GitHub geladen hat (ohne .git), bekommt keinen Code aus dem Netz. Höchstens einmal am Tag
fragt Gegenprobe GitHub ohne Konto, wann die neueste Version entstanden ist, und sagt, wenn es eine neuere gibt. Dabei
geht nichts aus dem Ordner oder aus Mappen mit, nur die Anfrage selbst. Woher die ZIP stammt, steht in stand.txt:
Beim Herunterladen schreibt GitHub (git archive, export-subst in .gitattributes) Commit und Zeit hinein.

Ausschalten mit VOICEAGENT_NO_UPDATE=1.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
MARKER = ".voiceagent-pakete"
# Was beim letzten Start passiert ist, für die Zustandsseite.
LAST_RUN = ".voiceagent-update.json"
# Die einzige Adresse, unter der Gegenprobe für eine ZIP-Installation nach einer neueren Version fragt.
GITHUB = "https://github.com/WenjaDebug/Crosscheck"
# Wann zuletzt bei GitHub nachgefragt wurde und was dabei herauskam.
CHECKED = ".voiceagent-nachgeschaut.json"
# So oft fragt eine ZIP-Installation höchstens nach (Sekunden).
CHECK_EVERY = 20 * 3600


def _git(root: Path, *args: str, timeout: float = 60) -> subprocess.CompletedProcess:
    # Nie nach Benutzername oder Passwort fragen: Dann würde der Start hängen.
    env = {**os.environ, "GIT_TERMINAL_PROMPT": "0", "GIT_ASKPASS": "", "SSH_ASKPASS": ""}
    return subprocess.run(
        ["git", *args], cwd=root, env=env, capture_output=True, text=True, timeout=timeout,
        encoding="utf-8", errors="replace",
    )


def pull(root: Path = ROOT) -> str:
    """Die neueste Version holen. Gibt zurück, was passiert ist, in einem Satz."""
    if os.environ.get("VOICEAGENT_NO_UPDATE", "").strip() not in ("", "0"):
        return "Automatisches Update ist ausgeschaltet (VOICEAGENT_NO_UPDATE)."
    if not (root / ".git").exists() or shutil.which("git") is None:
        return ""
    try:
        branch = _git(root, "rev-parse", "--abbrev-ref", "HEAD").stdout.strip()
        if branch != "main":
            return f"Du bist auf dem Zweig „{branch}“, deshalb kein automatisches Update. Zurück mit: git checkout main"
        if _git(root, "status", "--porcelain", "--untracked-files=no").stdout.strip():
            return "Es gibt eigene Änderungen im Ordner, deshalb kein automatisches Update."
        before = _git(root, "rev-parse", "HEAD").stdout.strip()
        result = _git(root, "pull", "--ff-only", "--quiet")
        if result.returncode != 0:
            return "Ich konnte nicht nach Updates schauen (offline oder nicht angemeldet). Gegenprobe startet mit der bisherigen Version."
        after = _git(root, "rev-parse", "HEAD").stdout.strip()
    except (OSError, subprocess.SubprocessError):
        return "Ich konnte nicht nach Updates schauen. Gegenprobe startet mit der bisherigen Version."
    if after == before:
        return "Gegenprobe ist aktuell."
    count = _git(root, "rev-list", "--count", f"{before}..{after}").stdout.strip()
    return f"Neue Version geladen ({count} Änderungen)." if count.isdigit() else "Neue Version geladen."


def stand(path: Path) -> tuple[str, int] | None:
    """(Commit, Zeit als Unix-Sekunden) einer ZIP von GitHub, sonst None. Im Git-Klon steht in stand.txt noch der
    Platzhalter $Format:…$, den erst GitHub beim Packen der ZIP ausfüllt."""
    try:
        parts = path.read_text(encoding="utf-8").split()
    except OSError:
        return None
    if len(parts) != 2 or not re.fullmatch(r"[0-9a-f]{7,40}", parts[0]) or not parts[1].isdigit():
        return None
    return parts[0], int(parts[1])


def _api(url: str) -> str:
    """Die Adresse bei GitHub, die ohne Konto sagt, wann der neueste Stand auf main entstanden ist."""
    owner_repo = url.rstrip("/").split("github.com/", 1)[1]
    return f"https://api.github.com/repos/{owner_repo}/commits?sha=main&per_page=1"


def _contexts():
    """Erst die Zertifikate des Systems; Python von python.org am Mac kennt die oft nicht (bis man „Install
    Certificates“ startet), dann die von pip mitgebrachten. Ohne Prüfung des Zertifikats geht nie etwas hinaus."""
    import ssl

    yield ssl.create_default_context()
    try:
        # pip steckt in jeder .venv und bringt seine Zertifikate mit.
        from pip._vendor import certifi

        yield ssl.create_default_context(cafile=certifi.where())
    except (ImportError, OSError):
        return


def latest(url: str = GITHUB, timeout: float = 4) -> int | None:
    """Wann der neueste Stand auf main entstanden ist (Unix-Sekunden), oder None (offline, gesperrt, Grenze erreicht)."""
    import ssl
    import urllib.error

    request = urllib.request.Request(_api(url), headers={"Accept": "application/vnd.github+json", "User-Agent": "Gegenprobe"})
    for context in _contexts():
        try:
            with urllib.request.urlopen(request, timeout=timeout, context=context) as response:
                found = json.loads(response.read(200_000).decode("utf-8"))
            when = found[0]["commit"]["committer"]["date"]
            return int(datetime.fromisoformat(when.replace("Z", "+00:00")).timestamp())
        except urllib.error.URLError as err:
            if isinstance(err.reason, ssl.SSLCertVerificationError):
                continue
            return None
        except (OSError, ValueError, KeyError, IndexError, TypeError):
            return None
    return None


def hint(root: Path = ROOT, prefix: Path | None = None, fetch=latest, now=time.time) -> str:
    """Für eine ZIP-Installation: ein Satz, wenn es auf GitHub eine neuere Version gibt, sonst "". Lädt nie Code."""
    if (root / ".git").exists():
        return ""
    mine = stand(root / "src" / "voiceagent" / "stand.txt")
    if mine is None:
        return ""
    cache = Path(prefix or sys.prefix) / CHECKED
    try:
        seen = json.loads(cache.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        seen = {}
    if not isinstance(seen, dict) or not isinstance(seen.get("time"), (int, float)) or now() - seen["time"] > CHECK_EVERY:
        # Auch eine Anfrage ohne Antwort zählt: Offline wartet der Start sonst jedes Mal.
        seen = {"time": now(), "newest": fetch()}
        try:
            cache.write_text(json.dumps(seen), encoding="utf-8")
        except OSError:
            pass
    newest = seen.get("newest")
    if not isinstance(newest, int) or newest <= mine[1]:
        return ""
    day = datetime.fromtimestamp(newest, timezone.utc).strftime("%d.%m.%Y")
    return (f"Es gibt eine neuere Version von Gegenprobe (vom {day}). Du kannst sie wie beim ersten Mal als ZIP laden: "
            f"{GITHUB}")


# Pakete der früheren Sprach- und Bildschirmteile (Whisper, Tesseract), die Gegenprobe nicht mehr braucht. Jeder Start
# nimmt sie aus der eigenen Umgebung .venv, falls sie noch da sind (zusammen mehrere hundert MB). Das hängt bewusst nicht
# an der Marke von sync_packages: Beim ersten Start nach dem Update läuft noch die alte Fassung dieser Datei.
RETIRED = ("faster-whisper", "ctranslate2", "onnxruntime", "av", "tokenizers", "huggingface-hub", "pytesseract",
           "pillow", "numpy", "sympy", "mpmath", "flatbuffers", "coloredlogs", "humanfriendly")


def _norm(name: str) -> str:
    return re.sub(r"[-_.]+", "-", name).lower()


def _environment() -> tuple[set[str], set[str]]:
    """Welche Pakete in dieser Umgebung stecken und welche davon andere Pakete (außer den alten) noch brauchen."""
    from importlib import metadata

    installed: set[str] = set()
    needed: set[str] = set()
    retired = {_norm(n) for n in RETIRED}
    for dist in metadata.distributions():
        name = _norm(dist.metadata.get("Name") or "")
        if not name:
            continue
        installed.add(name)
        if name in retired:
            continue
        for requirement in dist.requires or ():
            if "extra ==" in requirement:
                continue
            match = re.match(r"\s*([A-Za-z0-9][A-Za-z0-9._-]*)", requirement)
            if match:
                needed.add(_norm(match.group(1)))
    return installed, needed


def drop_retired(run=subprocess.run, environment=_environment) -> list[str]:
    """Alte Pakete entfernen, die nichts mehr braucht. Klappt es nicht, versucht es der nächste Start wieder."""
    installed, needed = environment()
    retired = [name for name in RETIRED if _norm(name) in installed and _norm(name) not in needed]
    if not retired:
        return []
    print("Nicht mehr gebrauchte Pakete werden entfernt …", flush=True)
    result = run([sys.executable, "-m", "pip", "uninstall", "-y", "--quiet", "--disable-pip-version-check", *retired])
    return retired if getattr(result, "returncode", 1) == 0 else []


def sync_packages(root: Path = ROOT, prefix: Path | None = None, run=subprocess.run) -> bool:
    """Pakete nachinstallieren, wenn sich pyproject.toml seit dem letzten Mal geändert hat."""
    project = root / "pyproject.toml"
    if not project.exists():
        return False
    digest = hashlib.sha256(project.read_bytes()).hexdigest()
    marker = Path(prefix or sys.prefix) / MARKER
    try:
        if marker.read_text(encoding="utf-8").strip() == digest:
            return False
    except OSError:
        pass
    print("Pakete werden aktualisiert, das dauert einen Moment …", flush=True)
    result = run([sys.executable, "-m", "pip", "install", "-e", str(root), "--quiet", "--disable-pip-version-check"])
    if getattr(result, "returncode", 1) != 0:
        print("Die Pakete konnten nicht aktualisiert werden. Gegenprobe versucht es beim nächsten Start noch einmal.")
        return False
    marker.write_text(digest, encoding="utf-8")
    return True


def version(root: Path = ROOT) -> dict:
    """Welche Version läuft: Stand von Git (Kurzname und Datum), sonst die Paketversion."""
    from . import __version__

    info = {"version": __version__, "commit": None, "date": None, "branch": None}
    if not (root / ".git").exists():
        found = stand(root / "src" / "voiceagent" / "stand.txt")
        if found:
            day = datetime.fromtimestamp(found[1], timezone.utc).strftime("%Y-%m-%d")
            info.update(commit=found[0][:7], date=day)
        return info
    if shutil.which("git"):
        try:
            line = _git(root, "log", "-1", "--format=%h %cs", timeout=5).stdout.split()
            branch = _git(root, "rev-parse", "--abbrev-ref", "HEAD", timeout=5).stdout.strip()
        except (OSError, subprocess.SubprocessError):
            return info
        if len(line) == 2:
            info.update(commit=line[0], date=line[1], branch=branch or None)
    return info


def record(message: str, packages: bool, prefix: Path | None = None) -> None:
    try:
        (Path(prefix or sys.prefix) / LAST_RUN).write_text(
            json.dumps({"time": time.time(), "message": message, "packages": packages}, ensure_ascii=False), encoding="utf-8",
        )
    except OSError:
        pass


def last_run(prefix: Path | None = None) -> dict:
    """Was beim letzten Start passiert ist: {"time", "message", "packages"} oder {}."""
    try:
        found = json.loads((Path(prefix or sys.prefix) / LAST_RUN).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return found if isinstance(found, dict) else {}


def main() -> int:
    message = pull()
    if not message:
        # Aus einer ZIP: kein Code aus dem Netz, nur ein Hinweis.
        try:
            message = hint()
        except OSError:
            message = ""
    if message:
        print(message, flush=True)
    packages = False
    try:
        packages = sync_packages()
    except OSError as err:
        print(f"Die Pakete konnten nicht geprüft werden: {err}")
    # Nur in der eigenen Umgebung .venv aufräumen, nie in einer fremden Python-Umgebung (etwa beim Entwickeln).
    if Path(sys.prefix).resolve() == (ROOT / ".venv").resolve():
        try:
            drop_retired()
        except (OSError, subprocess.SubprocessError):
            pass
    record(message, packages)
    if message.startswith("Neue Version") or packages:
        from . import dienst

        # Der Prüfer im Hintergrund läuft sonst mit der alten Version weiter.
        if dienst.restart():
            print("Der Prüfer im Hintergrund läuft jetzt mit der neuen Version.")
    # Ein Update darf den Start nie verhindern.
    return 0


if __name__ == "__main__":
    sys.exit(main())
