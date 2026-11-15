"""Die Übungsmappe aus dem Seitenfenster heraus in Excel öffnen.

Das Seitenfenster in Excel kann nichts herunterladen, ein Link auf die Datei tut dort nichts. Darum legt Gegenprobe
eine Kopie in den eigenen Datenordner und lässt sie vom System öffnen. An einer Kopie, die schon jemand geändert hat,
rührt es nichts: dann kommt eine frische daneben."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

DEMO = Path(__file__).parent / "web" / "demo" / "pruefer-uebung.xlsx"
NAME = "Übungsmappe"


def copy(folder: Path) -> Path:
    """Die Kopie, die geöffnet wird: eine unveränderte, wenn es sie schon gibt, sonst eine neue."""
    data = DEMO.read_bytes()
    folder.mkdir(parents=True, exist_ok=True)
    for number in range(1, 100):
        path = folder / (f"{NAME}.xlsx" if number == 1 else f"{NAME} {number}.xlsx")
        if not path.exists():
            path.write_bytes(data)
            return path
        if path.read_bytes() == data:
            return path
    raise OSError("Im Ordner liegen schon 99 geänderte Übungsmappen.")


def launch(path: Path) -> bool:
    """Die Datei in Excel öffnen (auf dem Mac ausdrücklich in Excel, sonst mit dem Programm für .xlsx)."""
    try:
        if sys.platform == "darwin":
            done = subprocess.run(["open", "-b", "com.microsoft.Excel", str(path)], capture_output=True, timeout=30)
            if done.returncode != 0:
                done = subprocess.run(["open", str(path)], capture_output=True, timeout=30)
            return done.returncode == 0
        if sys.platform == "win32":
            os.startfile(path)  # type: ignore[attr-defined]
            return True
        return subprocess.run(["xdg-open", str(path)], capture_output=True, timeout=30).returncode == 0
    except (OSError, subprocess.SubprocessError):
        return False


def open_copy(folder: Path) -> dict:
    """Für das Seitenfenster: Kopie anlegen, öffnen und sagen, wo sie liegt."""
    try:
        path = copy(folder)
    except OSError as err:
        return {"state": "error", "text": f"Die Übungsmappe konnte ich nicht anlegen ({err})."}
    return {"state": "ok" if launch(path) else "closed", "path": str(path), "name": path.name}
