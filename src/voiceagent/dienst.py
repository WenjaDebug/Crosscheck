"""Der Prüfer im Hintergrund: Excel findet sein Seitenfenster, ohne dass ein Terminal oder ein Browser offen ist.

ExcelAddin.command richtet das am Mac einmal ein (launchd, ~/Library/LaunchAgents). Den Port 8766 hält dann macOS
selbst offen, schon ab der Anmeldung, und startet diesen kleinen Dienst, sobald Excel dort anklopft. Von selbst bei der
Anmeldung starten lässt macOS Hintergrunddienste nicht immer (wartet etwa ein macOS-Update, startet es sie nur noch
auf Anfrage); auf eine Anfrage hin startet es sie immer. Er liefert nur das Seitenfenster für Excel
(https://localhost:8766): ohne Browser und ohne Excel fernzusteuern, damit er wenig
Arbeitsspeicher braucht und nie nach Rechten fragt. Excel öffnen und unter Start auf Prüfer klicken genügt.

Unter Windows richtet ExcelAddin.bat ihn ein (windows.py): Windows startet ihn bei jeder Anmeldung mit
`pythonw.exe -m voiceagent.dienst --ordner <Gegenprobe>`, ohne Fenster. Dann schreibt er selbst ins Protokoll und
startet nach einem Absturz selbst neu, was am Mac launchd übernimmt.
"""

from __future__ import annotations

import errno
import os
import plistlib
import re
import socket
import ssl
import subprocess
import sys
import tempfile
import time
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path

LABEL = "io.github.wenjadebug.voiceagent.pruefer"
# Größer wird das Protokoll nicht, es steht ohnehin nur drin, was schiefging.
MAX_LOG = 1_000_000
WAIT = 30
# So lange wartet macOS, bevor es ihn nach einem Ende wieder startet (wie launchd von sich aus).
THROTTLE = 10
# Unter diesem Namen hält launchd den Port für ihn offen und gibt ihn beim Start weiter.
SOCKET = "Listeners"
# So beginnt jeder Start im Protokoll; daran sieht ExcelAddin.command, ob er nach dem Neustart von selbst lief.
STARTS = "Prüfer im Hintergrund startet"
# So beginnt, was ExcelAddin.command dazu ins Protokoll schreibt.
NOTE = "ExcelAddin.command:"
# Wo man macOS erlaubt, ihn im Hintergrund zu starten (Systemeinstellungen → Allgemein → Anmeldeobjekte).
LOGIN_ITEMS = "x-apple.systempreferences:com.apple.LoginItems-Settings.extension"


def plist_path() -> Path:
    return Path.home() / "Library" / "LaunchAgents" / f"{LABEL}.plist"


def log_path(root: Path | None = None) -> Path:
    """Im Datenordner von Gegenprobe, damit man es leicht findet; neben den Schlüsseln läge es zu versteckt."""
    return (root or Path.cwd()) / "data" / "pruefer-dienst.log"


def log_tail(root: Path | None = None, lines: int = 12) -> list[str]:
    try:
        text = log_path(root).read_text(encoding="utf-8", errors="replace")
    except OSError:
        return []
    # Was ExcelAddin.command dazuschreibt (note), ist zum Nachlesen; hier zählt, was der Dienst selbst sagt.
    return [line for line in text.splitlines() if line.strip() and NOTE not in line][-lines:]


def plist(python: str, root: Path, log: Path, port: int = 8766, on_demand: bool = True) -> bytes:
    """Die Beschreibung für macOS: den Port offen halten und ihn starten, sobald jemand dort anklopft.

    on_demand=False ist der Rückfall wie früher: bei der Anmeldung starten und am Laufen halten, den Port öffnet er
    selbst. Das hält macOS zurück, wenn es Dienste nur auf Anfrage startet."""
    if on_demand:
        # Kein RunAtLoad und kein KeepAlive: Wie auch immer er endet (Absturz, beendet in der Aktivitätsanzeige),
        # die nächste Anfrage von Excel startet ihn wieder, und bis dahin braucht er keinen Arbeitsspeicher.
        start = {"Sockets": {SOCKET: {"SockNodeName": "127.0.0.1", "SockServiceName": str(port), "SockFamily": "IPv4",
                                      "SockType": "stream"}}}
    else:
        start = {"RunAtLoad": True, "KeepAlive": True}
    return plistlib.dumps({
        "Label": LABEL,
        "ProgramArguments": [python, "-m", "voiceagent.dienst"],
        "WorkingDirectory": str(root),
        "EnvironmentVariables": {
            "PATH": "/usr/bin:/bin:/usr/sbin:/sbin:/opt/homebrew/bin:/usr/local/bin",
            # Das Seitenfenster liest die Mappe selbst über Excel; Excel fernsteuern braucht der Dienst nicht.
            "VOICEAGENT_NO_EXCEL": "1",
        },
        # Aus geht er nur mit ExcelAddin.command --entfernen.
        **start,
        "ThrottleInterval": THROTTLE,
        # Wie ein Programm, nicht gedrosselt wie Hintergrundarbeit: Er rechnet, wenn man in Excel auf Prüfen klickt.
        "ProcessType": "Interactive",
        "StandardOutPath": str(log),
        "StandardErrorPath": str(log),
    })


def _domain() -> str:
    return f"gui/{os.getuid()}"


def _call(*args: str, run=subprocess.run) -> tuple[bool, str]:
    try:
        done = run(["launchctl", *args], capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30)
    except (OSError, subprocess.SubprocessError) as err:
        return False, str(err)
    return done.returncode == 0, " ".join(f"{done.stderr or ''} {done.stdout or ''}".split())[:300]


def note(text: str, root: Path | None = None) -> None:
    """Ins Protokoll des Dienstes schreiben, was ExcelAddin.command von macOS erfährt, damit man es später nachlesen kann."""
    try:
        with log_path(root).open("a", encoding="utf-8") as file:
            for line in text.splitlines() or [""]:
                file.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {NOTE} {line}\n")
    except OSError:
        pass


def _launchctl(*args: str, run=subprocess.run) -> bool:
    return _call(*args, run=run)[0]


def _output(*command: str, run=subprocess.run) -> str:
    """Was ein Befehl ausgibt, ungekürzt; "" wenn er nicht geht."""
    try:
        done = run(list(command), capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30)
    except (OSError, subprocess.SubprocessError):
        return ""
    return (done.stdout or "") if done.returncode == 0 else ""


def install(root: Path | None = None, python: str | None = None, run=subprocess.run, pause=time.sleep,
            start: bool = True, held=None, on_demand: bool = True) -> str:
    """Den Dienst anlegen (oder erneuern) und laden. Gibt "" zurück, sonst was macOS dazu sagt.

    start=False lädt ihn so wie macOS bei der Anmeldung, ohne nachzuhelfen: Startet ihn dann eine Anfrage
    (answers), tut sie das auch nach einem Neustart. held sagt, ob jemand den Port schon offen hat."""
    from . import addin

    target = plist_path()
    root = (root or Path.cwd()).resolve()
    log = log_path(root)
    log.parent.mkdir(parents=True, exist_ok=True)
    target.parent.mkdir(parents=True, exist_ok=True)
    data = plist(python or sys.executable, root, log, addin.PORT, on_demand)
    with tempfile.NamedTemporaryFile("wb", dir=target.parent, delete=False, suffix=".tmp") as file:
        file.write(data)
    Path(file.name).replace(target)
    # Läuft schon eine ältere Fassung, erst beenden; sonst nimmt macOS die neue nicht. Das Abräumen dauert
    # einen Moment, und wer zu früh lädt, bekommt „Input/output error“.
    if _launchctl("bootout", f"{_domain()}/{LABEL}", run=run):
        for _ in range(20):
            if not _launchctl("print", f"{_domain()}/{LABEL}", run=run):
                break
            pause(0.25)
    # Den Port öffnet gleich macOS für ihn; hat ihn noch jemand offen, ginge das nicht. Ein alter Dienst braucht
    # nach dem Beenden noch einen Moment.
    held = held or (lambda: busy(addin.PORT))
    for _ in range(20):
        if not held():
            break
        pause(0.5)
    else:
        return (f"Port {addin.PORT} ist belegt. Läuft noch VoiceAgent.command? Dort Ctrl+C drücken und das hier "
                "noch einmal starten")
    ok, said = _call("bootstrap", _domain(), str(target), run=run)
    if not ok:
        # Ein abgeschalteter Dienst lässt sich nicht laden: wieder einschalten, dann der alte Weg.
        _launchctl("enable", f"{_domain()}/{LABEL}", run=run)
        ok, again = _call("bootstrap", _domain(), str(target), run=run)
        if not ok:
            _launchctl("load", "-w", str(target), run=run)
            said = said or again
        if ok:
            note(f"bootstrap ging erst nach enable ({said or 'ohne Grund'}).", root)
        else:
            note(f"bootstrap ging nicht ({said or again or 'ohne Grund'}), darum load -w.", root)
    # „load“ meldet auch dann Erfolg, wenn nichts geladen ist; darum nachsehen.
    if not _launchctl("print", f"{_domain()}/{LABEL}", run=run):
        return said or "macOS hat den Dienst nicht geladen"
    if start:
        # Geladen heißt noch nicht gestartet: jetzt gleich starten, nicht erst bei der nächsten Anmeldung.
        _launchctl("kickstart", f"{_domain()}/{LABEL}", run=run)
    return ""


def repair(run=subprocess.run) -> None:
    """Startet er nicht von selbst: einschalten, falls er abgeschaltet ist, und anstoßen."""
    _launchctl("enable", f"{_domain()}/{LABEL}", run=run)
    if not _launchctl("kickstart", f"{_domain()}/{LABEL}", run=run):
        target = plist_path()
        _launchctl("bootstrap", _domain(), str(target), run=run)
        _launchctl("kickstart", f"{_domain()}/{LABEL}", run=run)


# --- Nachsehen, was macOS mit ihm macht ---


def _described() -> dict:
    try:
        return plistlib.loads(plist_path().read_bytes())
    except (OSError, ValueError, plistlib.InvalidFileException):
        return {}


def state(run=subprocess.run) -> dict:
    """Was launchd über den Dienst weiß: geladen, läuft (pid), wie oft gestartet, wie er zuletzt endete, abgeschaltet,
    und ob er auf Anfrage startet (sockets; früher eingerichtet oder im Rückfall startet er bei der Anmeldung)."""
    info = {"plist": plist_path().exists(), "loaded": False, "pid": None, "state": "", "runs": None, "exit": "",
            "signal": "", "disabled": False, "missing": "", "sockets": False}
    data = _described()
    info["sockets"] = bool(data.get("Sockets"))
    # Ist der Ordner verschoben oder umbenannt (VoiceAgent → Gegenprobe), findet macOS Python nicht mehr.
    for needed in [*(data.get("ProgramArguments") or [])[:1], data.get("WorkingDirectory")]:
        if isinstance(needed, str) and needed and not Path(needed).exists():
            info["missing"] = needed
            break
    text = _output("launchctl", "print", f"{_domain()}/{LABEL}", run=run)
    if text:
        info["loaded"] = True
        for line in text.splitlines():
            key, sep, value = line.strip().partition(" = ")
            if not sep:
                continue
            # Die erste Angabe gehört zum Dienst selbst, spätere zu Unterblöcken.
            if key == "state" and not info["state"]:
                info["state"] = value
            elif key == "pid" and info["pid"] is None and value.isdigit():
                info["pid"] = int(value)
            elif key == "runs" and info["runs"] is None and value.isdigit():
                info["runs"] = int(value)
            elif key == "last exit code" and not info["exit"]:
                info["exit"] = value
            elif key == "last terminating signal" and not info["signal"]:
                info["signal"] = value
    for line in _output("launchctl", "print-disabled", _domain(), run=run).splitlines():
        name, sep, value = line.strip().partition("=>")
        if sep and name.strip().strip('"') == LABEL:
            info["disabled"] = value.strip() in ("disabled", "true")
    return info


def ready(info: dict) -> bool:
    """Läuft er, oder hält macOS den Port für ihn offen und startet ihn, sobald Excel anklopft?"""
    return bool(info["pid"]) or (info.get("sockets", False) and info["loaded"] and not info["disabled"]
                                 and not info.get("missing"))


def waits(run=subprocess.run) -> bool:
    """Hält macOS den Port schon für ihn offen? Dann darf VoiceAgent.command ihn nicht selbst öffnen, und anklopfen,
    um es herauszufinden, würde ihn unnötig starten."""
    return installed() and bool(_described().get("Sockets")) and ready(state(run))


def report(run=subprocess.run) -> str:
    """Was launchd über den Dienst ausgibt, gekürzt, für das Protokoll."""
    lines = _output("launchctl", "print", f"{_domain()}/{LABEL}", run=run).splitlines()
    return "\n".join(line.rstrip() for line in lines[:80] if line.strip())


def explain(info: dict) -> str:
    """Warum er nicht läuft, in Worten."""
    if not info["plist"]:
        return "er ist nicht eingerichtet"
    if info.get("missing"):
        return f"{info['missing']} gibt es nicht mehr, ist der Ordner verschoben oder umbenannt?"
    if info["disabled"]:
        return "er ist in macOS abgeschaltet"
    if not info["loaded"]:
        return "macOS hat ihn nicht geladen"
    if info["pid"]:
        return "er läuft"
    if info.get("sockets"):
        # So ist es gewollt: Er braucht erst Arbeitsspeicher, wenn Excel das Fenster öffnet.
        return "er wartet, bis Excel ihn braucht"
    if not info["runs"]:
        return "macOS hat ihn geladen, aber nie gestartet"
    if info["signal"]:
        return f"er wurde beendet ({info['signal']})"
    if info["exit"] and info["exit"] not in ("0", "(never exited)"):
        return f"er ist mit einem Fehler beendet (Code {info['exit']})"
    return "er ist beendet"


def boot_time(run=subprocess.run) -> datetime | None:
    """Wann der Mac zuletzt gestartet wurde."""
    match = re.search(r"sec = (\d+)", _output("sysctl", "-n", "kern.boottime", run=run))
    return datetime.fromtimestamp(int(match.group(1))) if match else None


def since_boot(root: Path | None = None, boot: datetime | None = None) -> bool | None:
    """Ist er nach dem letzten Neustart von selbst gestartet? None, wenn sich das nicht sagen lässt: noch kein
    Neustart, seit er eingerichtet ist (jedes ExcelAddin.command schreibt die Beschreibung neu)."""
    if boot is None:
        return None
    try:
        if plist_path().stat().st_mtime >= boot.timestamp():
            return None
        text = log_path(root).read_text(encoding="utf-8", errors="replace")
    except FileNotFoundError:
        return False if plist_path().exists() else None
    except OSError:
        return None
    # Die Uhr kann beim Hochfahren noch ein paar Sekunden nachgestellt werden.
    earliest = boot - timedelta(seconds=10)
    for line in text.splitlines():
        if STARTS not in line:
            continue
        try:
            when = datetime.strptime(line[:19], "%Y-%m-%d %H:%M:%S")
        except ValueError:
            continue
        if when >= earliest:
            return True
    return False


def before(root: Path | None = None, run=subprocess.run, boot: datetime | None = None) -> tuple[list[str], bool]:
    """Bevor ExcelAddin.command ihn neu lädt: Lief er nach dem letzten Neustart von selbst, und läuft er jetzt?
    Gibt die Sätze dazu zurück und ob er nach dem Neustart nicht von selbst startete."""
    if not installed():
        return [], False
    lines = []
    boot = boot or boot_time(run)
    seen = since_boot(root, boot)
    info = state(run)
    shown = boot.strftime("%d.%m. %H:%M") if boot else ""
    if info["sockets"]:
        # Er startet erst, wenn Excel ihn braucht: Nicht gelaufen ist kein Fehler, nicht bereit schon.
        if seen is not None:
            lines.append(f"Seit dem letzten Neustart des Macs ({shown}) hat Excel ihn "
                         f"{'schon gestartet' if seen else 'noch nicht gebraucht'}.")
        if not ready(info):
            lines.append(f"Eben war er nicht bereit: {explain(info)}.")
        return lines, seen is not None and not ready(info)
    if seen is not None:
        lines.append(f"Nach dem letzten Neustart des Macs ({shown}) ist er {'von selbst' if seen else 'nicht von selbst'} gestartet.")
    if not info["pid"]:
        lines.append(f"Eben lief er nicht: {explain(info)}.")
        tail = log_tail(root, 6)
        if tail and (info["exit"] or info["signal"]):
            lines.append("Zuletzt stand im Protokoll (data/pruefer-dienst.log):")
            lines += [f"    {line}" for line in tail]
    return lines, seen is False


def remove(run=subprocess.run) -> None:
    target = plist_path()
    if not _launchctl("bootout", f"{_domain()}/{LABEL}", run=run) and target.exists():
        _launchctl("unload", "-w", str(target), run=run)
    target.unlink(missing_ok=True)


def installed() -> bool:
    return sys.platform == "darwin" and plist_path().exists()


def restart(run=subprocess.run) -> bool:
    """Nach einem Update: den Dienst mit der neuen Version neu starten, wenn er eingerichtet ist und läuft."""
    if sys.platform == "win32":
        from . import windows

        return windows.restart()
    if not installed():
        return False
    info = state(run)
    # Wartet er auf Excel, startet er beim nächsten Anklopfen ohnehin mit der neuen Version.
    if info["sockets"] and not info["pid"]:
        return False
    return _launchctl("kickstart", "-k", f"{_domain()}/{LABEL}", run=run)


def answers(timeout: float = 20, url: str | None = None, ca: Path | None = None) -> str:
    """Nach dem Einrichten: Liefert der Prüfer das Seitenfenster über https mit seinem Zertifikat aus?
    Gibt "" zurück, wenn ja, sonst den Grund."""
    from . import addin

    url = url or f"{addin.ORIGIN}/addin/pruefer.html"
    ca = ca or addin.folder() / "ca.pem"
    if not ca.exists():
        return "das Zertifikat fehlt"
    context = ssl.create_default_context(cafile=str(ca))
    # Excel prüft das Zertifikat mit macOS, nicht nach Pythons strengeren Regeln.
    context.verify_flags &= ~ssl.VERIFY_X509_STRICT
    deadline = time.monotonic() + timeout
    reason = "keine Antwort"
    while True:
        try:
            # Die erste Anfrage weckt ihn erst: Bis er antwortet, dauert es ein paar Sekunden.
            with urllib.request.urlopen(url, context=context, timeout=10) as response:
                # Ein Dienst von vor der Umbenennung antwortet noch mit dem alten Namen.
                page = response.read()
                if response.status == 200 and (b"Gegenprobe" in page or b"VoiceAgent" in page):
                    return ""
                reason = f"Antwort {response.status}"
        except ssl.SSLError as err:
            reason = f"Zertifikat passt nicht ({err.reason or err})"
        except OSError as err:
            reason = f"keine Verbindung ({getattr(err, 'reason', None) or err})"
        if time.monotonic() >= deadline:
            return reason
        time.sleep(1)


def busy(port: int, host: str = "127.0.0.1") -> bool:
    """Antwortet auf dem Port schon jemand (der Dienst oder VoiceAgent.command)?"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        probe.settimeout(1)
        return probe.connect_ex((host, port)) == 0


def activated(lib=None) -> list[socket.socket]:
    """Hat macOS ihn gestartet, weil jemand am Port anklopft? Dann hat launchd den Port offen und gibt ihn hier her
    (launch_activate_socket). Sonst [], etwa von Hand gestartet."""
    if sys.platform != "darwin" and lib is None:
        return []
    import ctypes

    # Es steckt in libSystem, die jedes Programm am Mac schon geladen hat.
    for path in ((None, "/usr/lib/libSystem.B.dylib") if lib is None else ()):
        try:
            lib = ctypes.CDLL(path)
        except OSError:
            continue
        if hasattr(lib, "launch_activate_socket"):
            break
        lib = None
    if lib is None:
        print("launch_activate_socket gibt es hier nicht.", flush=True)
        return []
    call = lib.launch_activate_socket
    call.argtypes = [ctypes.c_char_p, ctypes.POINTER(ctypes.POINTER(ctypes.c_int)), ctypes.POINTER(ctypes.c_size_t)]
    call.restype = ctypes.c_int
    fds, count = ctypes.POINTER(ctypes.c_int)(), ctypes.c_size_t(0)
    failed = call(SOCKET.encode(), ctypes.byref(fds), ctypes.byref(count))
    if failed or not fds:
        # ESRCH heißt nur: nicht von macOS gestartet, etwa von Hand. Alles andere gehört ins Protokoll.
        if failed and failed != errno.ESRCH:
            print(f"launch_activate_socket ging nicht: {os.strerror(failed)} ({failed}).", flush=True)
        return []
    try:
        return [socket.socket(fileno=fds[i]) for i in range(count.value)]
    finally:
        lib.free(fds)


def _trim(log: Path) -> None:
    try:
        if log.stat().st_size > MAX_LOG:
            os.truncate(log, 0)
    except OSError:
        pass


def main(pause=time.sleep) -> int:
    from . import addin

    _trim(log_path())
    # Im Rückfall (oder von Hand gestartet) öffnet er den Port selbst.
    on_demand = installed() and bool(_described().get("Sockets"))
    sockets = activated() if on_demand else []
    woken = ", geweckt über den Port" if sockets else ""
    print(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {STARTS} ({sys.executable}, {Path.cwd()}{woken}).", flush=True)
    if not (addin.folder() / "ca.pem").exists():
        print("Das Excel-Add-in ist nicht eingerichtet (ExcelAddin.command). Der Dienst wartet.", flush=True)
        # macOS startet ihn danach wieder; nur selten nachsehen, damit das Protokoll klein bleibt.
        pause(WAIT * 10)
        return 0
    found = addin.ready()
    if found is None:
        # Der Grund steht schon im Protokoll; macOS versucht es gleich noch einmal.
        return 1
    if on_demand and not sockets:
        # Den Port hat macOS offen und hätte ihn übergeben sollen; selbst öffnen geht dann nicht. ExcelAddin.command
        # fällt in dem Fall auf den Start bei der Anmeldung zurück.
        print(f"macOS hat den Port {addin.PORT} nicht übergeben.", flush=True)
        return 1
    # Von Hand gestartet (oder im Rückfall): Läuft gerade VoiceAgent.command, liefert es das Seitenfenster selbst;
    # danach übernimmt wieder der Dienst.
    while not sockets and busy(addin.PORT):
        time.sleep(WAIT)
    import uvicorn

    from .app import create_app
    from .config import Settings

    app = create_app(Settings())
    cert, key = found
    config = uvicorn.Config(app, host="127.0.0.1", port=addin.PORT, log_level="warning",
                            ssl_certfile=str(cert), ssl_keyfile=str(key))
    uvicorn.Server(config).run(sockets=sockets or None)
    return 0


def background(folder: Path, pause=time.sleep, run=None, tries: int | None = None) -> int:
    """Wie Windows ihn bei der Anmeldung startet (pythonw, ohne Fenster, irgendwo): im Ordner von Gegenprobe, mit
    dem Protokoll in data/pruefer-dienst.log und nach einem Absturz nach THROTTLE Sekunden wieder, wie launchd."""
    os.chdir(folder)
    # Das Seitenfenster liest die Mappe selbst über Excel; Excel fernsteuern braucht der Dienst nicht.
    os.environ.setdefault("VOICEAGENT_NO_EXCEL", "1")
    log = log_path()
    log.parent.mkdir(parents=True, exist_ok=True)
    _trim(log)
    # pythonw hat weder stdout noch stderr; ohne sie bricht uvicorn ab.
    stream = open(log, "a", encoding="utf-8", errors="replace", buffering=1)  # noqa: SIM115 (bleibt bis zum Ende offen)
    sys.stdout = sys.stderr = stream
    run = run or main
    count = 0
    while True:
        count += 1
        try:
            code = run()
        except SystemExit as stop:
            code = stop.code if isinstance(stop.code, int) else 1
        except Exception as err:  # noqa: BLE001 (jeder Fehler gehört ins Protokoll, dann neu starten)
            print(f"{time.strftime('%Y-%m-%d %H:%M:%S')} Abgebrochen: {err!r}", flush=True)
            code = 1
        if code == 0 or (tries is not None and count >= tries):
            return code
        pause(THROTTLE)
        _trim(log)


def cli(argv: list[str] | None = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(description="Der Prüfer im Hintergrund")
    parser.add_argument("--ordner", type=Path, help="Ordner von Gegenprobe (so startet ihn Windows bei der Anmeldung)")
    args = parser.parse_args(argv)
    if args.ordner is not None:
        return background(args.ordner)
    return main()


if __name__ == "__main__":
    sys.exit(cli())
