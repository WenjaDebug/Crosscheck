"""Zustand und Fehlerbericht: Was läuft, was fehlt, und ein Bericht, den man weitergeben kann.

Die Zustandsseite prüft jeden Baustein (Version, Excel, Datenordner) und sagt in einem Satz, was zu tun ist,
wenn etwas fehlt. Der Fehlerbericht ist eine Textdatei mit diesem Zustand und den letzten Warnungen und
Fehlern. Er enthält nie Zellinhalte: Namen der Ordner unter dem eigenen Benutzer werden zu „~“, E-Mail,
IBAN und Ähnliches geschwärzt.
"""

from __future__ import annotations

import logging
import platform
import re
import sys
import time
from collections import deque
from datetime import datetime
from importlib import metadata
from pathlib import Path

from .privacy import Privacy

PACKAGES = ("fastapi", "uvicorn", "pypdf")
# Was nach einem Schlüssel aussieht, kommt nie in den Bericht, auch nicht in einer Fehlermeldung.
_SECRETS = re.compile(
    r"(AIza[0-9A-Za-z_\-]{20,}|gsk_[0-9A-Za-z]{20,}|sk-[0-9A-Za-z_\-]{16,}|nvapi-[0-9A-Za-z_\-]{16,}"
    r"|(?i:bearer)\s+[0-9A-Za-z._\-]{12,}|(?i:(?:api[_-]?key|key|token)=)[0-9A-Za-z._\-]{8,})"
)


class RecentLog(logging.Handler):
    """Die letzten Warnungen und Fehler, nur im Arbeitsspeicher."""

    def __init__(self, size: int = 200) -> None:
        super().__init__(logging.WARNING)
        self.records: deque[tuple[float, str, str, str]] = deque(maxlen=size)
        self.setFormatter(logging.Formatter("%(message)s"))
        self._last: logging.LogRecord | None = None

    def emit(self, record: logging.LogRecord) -> None:
        if record is self._last:
            return  # derselbe Eintrag über „uvicorn.error“ und über den Haupt-Logger
        self._last = record
        try:
            self.records.append((record.created, record.levelname, record.name, self.format(record)))
        except Exception:  # noqa: BLE001 - ein Protokoll darf nie selbst etwas kaputt machen
            pass


LOG = RecentLog()


def install() -> None:
    """Warnungen von Gegenprobe und vom Server mitschreiben. Mehrfaches Aufrufen schadet nicht.

    Läuft erst beim Start des Servers, weil uvicorn vorher seine eigenen Logger neu einrichtet."""
    root = logging.getLogger()
    if not root.handlers:
        # Was bisher im Terminal stand, soll dort auch weiter stehen.
        terminal = logging.StreamHandler()
        terminal.setLevel(logging.WARNING)
        terminal.setFormatter(logging.Formatter("%(levelname)s %(name)s: %(message)s"))
        root.addHandler(terminal)
    for logger in (root, logging.getLogger("uvicorn.error")):
        if LOG not in logger.handlers:
            logger.addHandler(LOG)


def scrub(text: str, secrets: list[str] = ()) -> str:
    """Schlüssel, eigene Ordnernamen und persönliche Daten aus einem Text entfernen."""
    for secret in secrets:
        if secret and len(secret) >= 6:
            text = text.replace(secret, "[Schlüssel]")
    text = _SECRETS.sub("[Schlüssel]", text)
    home = str(Path.home())
    if len(home) > 1:
        text = text.replace(home, "~")
    return Privacy().text(text)


def _package(name: str) -> str:
    try:
        return metadata.version(name)
    except metadata.PackageNotFoundError:
        return "fehlt"


def _ago(seconds: float) -> str:
    if seconds < 90:
        return "gerade eben"
    if seconds < 5400:
        return f"vor {round(seconds / 60)} Minuten"
    return f"vor {round(seconds / 3600)} Stunden"


def _check(name: str, ok: bool | None, text: str) -> dict:
    return {"name": name, "ok": ok, "text": text}


def _excel_check(excel) -> dict:
    """Ob Gegenprobe das laufende Excel lesen darf. Ohne Dateinamen: Der kommt auch in den Fehlerbericht."""
    state = excel.status.state
    if state == "ok":
        return _check("Excel", True, "Liest das offene Excel direkt mit.")
    if state == "denied":
        where = (
            " Erlaube es in den Systemeinstellungen unter Datenschutz und Sicherheit, Automation: bei Terminal „Microsoft Excel“."
            if excel.platform == "darwin" else ""
        )
        return _check("Excel", False, f"Darf Excel nicht lesen.{where}")
    if state == "error":
        return _check("Excel", False, "Excel hat zuletzt nicht geantwortet. Details stehen im Fehlerbericht.")
    if state == "empty":
        return _check("Excel", None, "Excel ist offen, aber ohne Mappe.")
    if state == "busy":
        return _check("Excel", True, "Liest das offene Excel direkt mit. Gerade ist Excel beschäftigt, zum Beispiel weil in einer Zelle getippt wird.")
    return _check("Excel", None, "Wird direkt mitgelesen, sobald Excel offen ist.")


def status(settings, root: Path | None = None, excel=None) -> dict:
    """Der Zustand aller Bausteine, mit einem Satz je Baustein. ok: True, False oder None (geht auch ohne)."""
    from .update import ROOT, last_run, version

    info = version(root or ROOT)
    checks = []
    if info["commit"]:
        # Ohne Zweig: aus einer ZIP von GitHub.
        where = "" if info["branch"] in ("main", None) else f", Zweig {info['branch']}"
        text = f"Stand {datetime.strptime(info['date'], '%Y-%m-%d').strftime('%d.%m.%Y')} ({info['commit']}{where})."
    else:
        text = f"Version {info['version']}."
    update = last_run()
    if update.get("message"):
        text += f" Beim Start: {update['message']}"
    checks.append(_check("Version", True if info["branch"] in ("main", None) else None, text))
    if excel is not None:
        checks.append(_excel_check(excel))
    data_dir = Path(settings.data_dir)
    checks.append(_check("Daten", True, f"Ordner {scrub(str(data_dir.resolve()))}."))
    problems = [r for r in LOG.records if r[1] in ("ERROR", "CRITICAL")]
    if problems:
        checks.append(_check("Fehler", False, f"{len(problems)} Fehler seit dem Start, zuletzt {_ago(time.time() - problems[-1][0])}. Details stehen im Fehlerbericht."))
    return {"checks": checks, "version": info, "update": update, "warnings": len(LOG.records)}


def report(state: dict, settings, secrets: list[str] = ()) -> str:
    """Der Fehlerbericht als Text: Zustand, Technik, letzte Warnungen. Nichts Persönliches."""
    lines = [
        "Gegenprobe-Fehlerbericht",
        f"Erstellt: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        "",
        "## Zustand",
    ]
    for check in state["checks"]:
        mark = {True: "ok", False: "FEHLT", None: "-"}[check["ok"]]
        lines.append(f"[{mark}] {check['name']}: {check['text']}")
    lines += [
        "",
        "## Technik",
        f"System: {platform.system()} {platform.release()} ({platform.machine()})",
        f"Python: {sys.version.split()[0]}",
        "Pakete: " + ", ".join(f"{p} {_package(p)}" for p in PACKAGES),
        "",
        "## Letzte Warnungen und Fehler",
    ]
    if not LOG.records:
        lines.append("Keine seit dem Start.")
    for created, level, name, message in list(LOG.records)[-60:]:
        when = datetime.fromtimestamp(created).strftime("%H:%M:%S")
        lines.append(f"{when} {level} {name}: {message}")
    return scrub("\n".join(lines) + "\n", secrets)
