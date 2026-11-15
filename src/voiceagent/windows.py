"""Das Seitenfenster unter Windows: was am Mac launchd, der Schlüsselbund und der wef-Ordner tun.

- Vertrauen: Die eigene Zertifizierungsstelle kommt mit `certutil -user -addstore Root` in die Stammzertifikate des
  Benutzers. Windows fragt dabei einmal nach („Sicherheitswarnung … Ja“), Adminrechte braucht es nicht.
- Anmelden bei Excel: Excel 2016 und neuer lesen Entwickler-Add-ins aus der Registry,
  HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer, Name = Add-in-ID, Wert = Pfad der Manifestdatei
  (so macht es auch Microsofts office-addin-dev-settings).
- Im Hintergrund: HKCU\\…\\CurrentVersion\\Run startet den Prüfer bei jeder Anmeldung mit pythonw.exe, also ohne
  Fenster und ohne Adminrechte. Gleich nach dem Einrichten startet er auch einmal so.

Alles mit der Standardbibliothek. Die Registry kommt als reg herein (sonst winreg), Befehle über run und popen:
So lässt sich das auch ohne Windows prüfen.
"""

from __future__ import annotations

import base64
import os
import ssl
import subprocess
import sys
import time
from pathlib import Path

from . import zertifikat

OFFICE_KEY = r"Software\Microsoft\Office\16.0"
WEF_KEY = OFFICE_KEY + r"\WEF\Developer"
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_NAME = "Gegenprobe Pruefer"
# Woran man den Prüfer im Hintergrund unter den Python-Prozessen erkennt.
MODULE = "voiceagent.dienst"
# Aus der Windows-API (Prozesserstellung); subprocess kennt sie nur unter Windows.
DETACHED_PROCESS = 0x00000008
CREATE_NEW_PROCESS_GROUP = 0x00000200
CREATE_BREAKAWAY_FROM_JOB = 0x01000000
# Kein Konsolenfenster für PowerShell und taskkill, auch wenn der Prüfer selbst (pythonw) sie aufruft.
CREATE_NO_WINDOW = 0x08000000


# --- Registry ---


def registry(reg=None):
    if reg is not None:
        return reg
    import winreg

    return winreg


def read(path: str, name: str, reg=None) -> str | None:
    reg = registry(reg)
    try:
        with reg.OpenKey(reg.HKEY_CURRENT_USER, path, 0, reg.KEY_READ) as key:
            value, _ = reg.QueryValueEx(key, name)
    except OSError:
        return None
    return value if isinstance(value, str) else None


def write(path: str, name: str, value: str, reg=None) -> None:
    reg = registry(reg)
    with reg.CreateKeyEx(reg.HKEY_CURRENT_USER, path, 0, reg.KEY_SET_VALUE) as key:
        reg.SetValueEx(key, name, 0, reg.REG_SZ, value)


def delete(path: str, name: str, reg=None) -> bool:
    reg = registry(reg)
    try:
        with reg.OpenKey(reg.HKEY_CURRENT_USER, path, 0, reg.KEY_SET_VALUE) as key:
            reg.DeleteValue(key, name)
    except OSError:
        return False
    return True


def exists(path: str, reg=None) -> bool:
    reg = registry(reg)
    try:
        with reg.OpenKey(reg.HKEY_CURRENT_USER, path, 0, reg.KEY_READ):
            return True
    except OSError:
        return False


# --- Excel ---


def office_seen(reg=None) -> bool:
    """Lief Office 2016 oder neuer schon einmal unter diesem Benutzer?"""
    return exists(OFFICE_KEY, reg)


def register(manifest: Path, addin_id: str, reg=None) -> None:
    """Das Add-in als Entwickler-Add-in bei Excel anmelden."""
    write(WEF_KEY, addin_id, str(manifest), reg)


def registered(addin_id: str, reg=None) -> str | None:
    """Der Pfad, unter dem Excel das Manifest sucht, oder None."""
    return read(WEF_KEY, addin_id, reg)


def unregister(addin_id: str, reg=None) -> bool:
    return delete(WEF_KEY, addin_id, reg)


# --- Vertrauen ---


def trusted(ca: Path, enum=None) -> bool:
    """Steht die Zertifizierungsstelle schon bei den Stammzertifikaten (des Benutzers oder des Rechners)?"""
    enum = enum or getattr(ssl, "enum_certificates", None)
    if enum is None:
        return False
    try:
        wanted = zertifikat.der(ca)
        return any(cert == wanted for cert, encoding, _ in enum("ROOT") if encoding == "x509_asn")
    except (OSError, ValueError, ssl.SSLError):
        return False


def trust(ca: Path, run=subprocess.run) -> tuple[bool, str]:
    """Mit certutil zu den Stammzertifikaten des Benutzers. Windows fragt dabei einmal nach. (geklappt, was certutil sagt)"""
    # certutil nimmt am sichersten die Binärform (.cer).
    cer = ca.with_suffix(".cer")
    cer.write_bytes(zertifikat.der(ca))
    try:
        done = run(["certutil", "-user", "-addstore", "Root", str(cer)], capture_output=True, text=True,
                   encoding="utf-8", errors="replace", timeout=600)
    except (OSError, subprocess.SubprocessError) as err:
        return False, str(err)
    said = " ".join(f"{done.stdout or ''} {done.stderr or ''}".split())[:300]
    return done.returncode == 0, said


# --- Der Prüfer im Hintergrund ---


def pythonw(python: str) -> str:
    """pythonw.exe neben python.exe startet ohne Fenster."""
    quiet = Path(python).with_name("pythonw.exe")
    return str(quiet) if Path(python).name.lower() == "python.exe" and quiet.exists() else python


def command(python: str, root: Path) -> list[str]:
    return [pythonw(python), "-m", MODULE, "--ordner", str(root)]


def autostart(python: str, root: Path, reg=None) -> None:
    """Bei jeder Anmeldung starten (HKCU …\\Run, ohne Adminrechte)."""
    write(RUN_KEY, RUN_NAME, subprocess.list2cmdline(command(python, root)), reg)


def autostarts(reg=None) -> str | None:
    """Die Befehlszeile, mit der Windows ihn bei der Anmeldung startet, oder None."""
    return read(RUN_KEY, RUN_NAME, reg)


_LIST = ("Get-CimInstance Win32_Process -Filter \"Name like 'python%'\" | "
         "ForEach-Object { \"$($_.ProcessId)`t$($_.CommandLine)\" }")


def processes(run=subprocess.run) -> list[int]:
    """Die Prozesse des Prüfers im Hintergrund (auch der Starter aus .venv, der das echte Python startet)."""
    # -EncodedCommand: so kommen Anführungszeichen und $ unverändert bei PowerShell an.
    script = base64.b64encode(_LIST.encode("utf-16-le")).decode("ascii")
    try:
        done = run(["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", script], capture_output=True,
                   text=True, encoding="utf-8", errors="replace", timeout=60, creationflags=CREATE_NO_WINDOW)
    except (OSError, subprocess.SubprocessError):
        return []
    found = []
    for line in (done.stdout or "").splitlines():
        pid, _, line_command = line.partition("\t")
        if pid.strip().isdigit() and MODULE in line_command and int(pid) != os.getpid():
            found.append(int(pid))
    return found


def stop(run=subprocess.run) -> bool:
    """Läuft er (etwa eine ältere Fassung), beenden. True, wenn einer lief."""
    found = processes(run)
    if not found:
        return False
    try:
        run(["taskkill", "/F", *[arg for pid in found for arg in ("/PID", str(pid))]], capture_output=True,
            text=True, encoding="utf-8", errors="replace", timeout=60, creationflags=CREATE_NO_WINDOW)
    except (OSError, subprocess.SubprocessError):
        pass
    return True


def start(args: list[str] | str, popen=subprocess.Popen) -> str:
    """Jetzt gleich starten, ohne Fenster und unabhängig von diesem Fenster. "" oder der Grund."""
    flags = DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
    last = ""
    # Läuft das Fenster in einem Job, der beim Schließen alles beendet, soll er davon frei sein; nicht jeder Job
    # erlaubt das, dann ohne.
    for extra in (CREATE_BREAKAWAY_FROM_JOB, 0):
        try:
            popen(args, creationflags=flags | extra, close_fds=True, stdin=subprocess.DEVNULL,
                  stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return ""
        except OSError as err:
            last = str(err)
    return last or "Windows hat ihn nicht gestartet"


def install(root: Path, python: str | None = None, reg=None, run=subprocess.run, popen=subprocess.Popen,
            pause=None, held=None) -> str:
    """Einrichten (oder erneuern) und gleich starten. "" oder der Grund."""
    from . import addin, dienst

    pause = pause or time.sleep
    python = python or sys.executable
    root = root.resolve()
    dienst.log_path(root).parent.mkdir(parents=True, exist_ok=True)
    # Eine ältere Fassung zuerst beenden, sonst liefe sie weiter und hielte den Port.
    stopped = stop(run)
    held = held or (lambda: dienst.busy(addin.PORT))
    for _ in range(20 if stopped else 2):
        if not held():
            break
        pause(0.5)
    else:
        return (f"Port {addin.PORT} ist belegt. Läuft noch VoiceAgent.bat? Schließ dessen Fenster und starte das hier "
                "noch einmal")
    try:
        autostart(python, root, reg)
    except OSError as err:
        return f"Windows lässt ihn nicht bei der Anmeldung starten ({err})"
    return start(command(python, root), popen)


def remove(reg=None, run=subprocess.run) -> None:
    delete(RUN_KEY, RUN_NAME, reg)
    stop(run)


def installed(reg=None) -> bool:
    return sys.platform == "win32" and autostarts(reg) is not None


def restart(reg=None, run=subprocess.run, popen=subprocess.Popen, pause=time.sleep) -> bool:
    """Nach einem Update: läuft er, mit der neuen Version neu starten, genau so, wie Windows ihn bei der Anmeldung
    startet."""
    from . import addin, dienst

    line = autostarts(reg) if sys.platform == "win32" else None
    if not line or not stop(run):
        return False
    for _ in range(20):
        if not dienst.busy(addin.PORT):
            break
        pause(0.5)
    return not start(line, popen)
