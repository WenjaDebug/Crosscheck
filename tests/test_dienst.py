"""Der Prüfer im Hintergrund: macOS hält seinen Port offen und startet ihn, sobald Excel anklopft. Excel findet sein
Seitenfenster ohne Terminal, auch nach einem Neustart."""

import ctypes
import os
import plistlib
import shutil
import socket
import ssl
import sys
import tempfile
import threading
import time
import urllib.request
from types import SimpleNamespace

import pytest

from voiceagent import addin, dienst, update


class Launchctl:
    """Wie launchctl: merkt sich, ob der Dienst geladen ist und läuft, und die Aufrufe; was in fail steht, geht
    schief. starts=False: macOS lädt ihn, startet ihn aber nicht gleich; demand=False: auch nicht, wenn jemand am
    Port anklopft (knock). Dazu sysctl für die Startzeit des Macs."""

    def __init__(self, fail=(), loaded=False, starts=True, disabled=False, boot=None, demand=True):
        self.calls, self.fail, self.loaded = [], set(fail), loaded
        self.starts, self.running, self.disabled, self.boot, self.demand = starts, False, disabled, boot, demand

    def knock(self):
        """Excel (oder ExcelAddin.command) klopft am Port an, den macOS für ihn offen hält."""
        if self.loaded and self.demand and not self.disabled:
            self.running = True

    def __call__(self, command, **_):
        if command[0] == "sysctl":
            stamp = int(self.boot.timestamp()) if self.boot else 0
            return SimpleNamespace(returncode=0 if self.boot else 1, stdout=f"{{ sec = {stamp}, usec = 0 }} Mon Oct 27", stderr="")
        verb = command[1]
        self.calls.append(command[1:])
        failed = verb in self.fail
        if not failed and verb in ("bootstrap", "load"):
            self.loaded = True
            self.running = self.starts and not self.disabled
        elif not failed and verb in ("bootout", "unload"):
            failed, self.loaded, self.running = not self.loaded, False, False
        elif not failed and verb in ("print", "kickstart"):
            failed = not self.loaded
            if verb == "kickstart" and not failed:
                self.running = True
        elif not failed and verb == "enable":
            self.disabled = False
        out = ""
        if verb == "print" and not failed:
            now = "running" if self.running else "not running"
            pid = "\tpid = 4242\n" if self.running else ""
            out = (f"gui/501/{dienst.LABEL} = {{\n\tstate = {now}\n\truns = {1 if self.running else 0}\n{pid}"
                   "\tlast exit code = (never exited)\n\tendpoints = {\n\t\tstate = active\n\t}\n}\n")
        elif verb == "print-disabled":
            out = f'disabled services = {{\n\t"{dienst.LABEL}" => {"disabled" if self.disabled else "enabled"}\n}}\n'
        return SimpleNamespace(returncode=1 if failed else 0, stdout=out,
                               stderr="Bootstrap failed: 5: Input/output error" if failed else "")


OUTPUT = dienst._output


def answering(run):
    """Wie dienst.answers: klopft an und sagt, ob er antwortet."""
    def answers(**_):
        run.knock()
        return "" if run.running else "keine Verbindung (refused)"
    return answers


def wired(monkeypatch, run):
    """dienst ruft launchctl und sysctl über run auf."""
    monkeypatch.setattr(dienst, "_call", lambda *args, **_: (run(["launchctl", *args]).returncode == 0, ""))
    monkeypatch.setattr(dienst, "_output", lambda *command, **_: OUTPUT(*command, run=run))


@pytest.fixture
def home(tmp_path, monkeypatch):
    monkeypatch.setattr(dienst.Path, "home", lambda: tmp_path)
    monkeypatch.setenv("VOICEAGENT_ADDIN_DIR", str(tmp_path / ".voiceagent" / "addin"))
    monkeypatch.setattr(dienst.sys, "platform", "darwin")
    monkeypatch.setattr(dienst.os, "getuid", lambda: 501, raising=False)
    # Niemand sonst hat den Port offen (auch kein VoiceAgent.command, das gerade läuft).
    monkeypatch.setattr(dienst, "busy", lambda port, host="127.0.0.1": False)
    return tmp_path


def test_macos_holds_the_port_and_starts_the_service_when_excel_knocks(tmp_path):
    data = plistlib.loads(dienst.plist("/v/.venv/bin/python", tmp_path, tmp_path / "dienst.log", 8766))
    assert data["Label"] == dienst.LABEL and data["ProgramArguments"] == ["/v/.venv/bin/python", "-m", "voiceagent.dienst"]
    assert data["WorkingDirectory"] == str(tmp_path) and data["EnvironmentVariables"]["VOICEAGENT_NO_EXCEL"] == "1"
    assert data["Sockets"] == {"Listeners": {"SockNodeName": "127.0.0.1", "SockServiceName": "8766",
                                             "SockFamily": "IPv4", "SockType": "stream"}}
    # Was macOS zurückhält, wenn es Dienste nur auf Anfrage startet, steht nicht drin.
    assert not {"RunAtLoad", "KeepAlive", "StartInterval"} & set(data)
    assert data["ProcessType"] == "Interactive" and data["ThrottleInterval"] == 10
    assert "/usr/bin" in data["EnvironmentVariables"]["PATH"] and data["StandardErrorPath"].endswith("dienst.log")


def test_install_replaces_an_old_service_and_remove_takes_it_away(home, tmp_path):
    run = Launchctl()
    assert dienst.install(root=tmp_path, python="/v/python", run=run, pause=lambda _: None) == ""
    target = home / "Library" / "LaunchAgents" / f"{dienst.LABEL}.plist"
    job = f"gui/501/{dienst.LABEL}"
    assert plistlib.loads(target.read_bytes())["ProgramArguments"][0] == "/v/python"
    assert run.calls == [["bootout", job], ["bootstrap", "gui/501", str(target)], ["print", job], ["kickstart", job]]
    # Eine neue Fassung: erst warten, bis die alte weg ist, dann laden und starten.
    run.calls.clear()
    assert dienst.install(root=tmp_path, python="/v/python2", run=run, pause=lambda _: None) == ""
    assert [c[0] for c in run.calls] == ["bootout", "print", "bootstrap", "print", "kickstart"]
    assert dienst.installed() and dienst.restart(run=run) and run.calls[-1] == ["kickstart", "-k", job]
    # Wartet er auf Excel, nimmt er beim nächsten Start ohnehin die neue Version: nicht unnötig starten.
    run.running = False
    assert not dienst.restart(run=run) and run.calls[-1][0] != "kickstart"
    dienst.remove(run=run)
    assert not target.exists() and not dienst.installed() and not dienst.restart(run=run)


def test_a_port_that_is_still_open_is_said(home, tmp_path):
    """Hat noch jemand den Port offen (VoiceAgent.command), kann macOS ihn nicht für den Dienst öffnen."""
    run = Launchctl()
    said = dienst.install(root=tmp_path, python="/v/python", run=run, pause=lambda _: None, held=lambda: True)
    assert said.startswith("Port 8766 ist belegt. Läuft noch VoiceAgent.command? Dort Ctrl+C drücken")
    assert "bootstrap" not in [c[0] for c in run.calls]
    # Ein alter Dienst gibt ihn nach dem Beenden nach einem Moment frei.
    frees = iter([True, True, False])
    assert dienst.install(root=tmp_path, python="/v/python", run=run, pause=lambda _: None, held=lambda: next(frees)) == ""


def test_an_older_mac_gets_launchctl_load(home, tmp_path):
    run = Launchctl(fail={"bootstrap"})
    assert dienst.install(root=tmp_path, python="/v/python", run=run) == ""
    assert [c[0] for c in run.calls] == ["bootout", "bootstrap", "enable", "bootstrap", "load", "print", "kickstart"]
    # „load“ sagt auch ohne Erfolg 0: Ob er wirklich da ist, zeigt erst „print“, und dann sagt er, warum nicht.
    said = dienst.install(root=tmp_path, python="/v/python", run=Launchctl(fail={"bootstrap", "print"}))
    assert said == "Bootstrap failed: 5: Input/output error"
    assert (tmp_path / "data").is_dir()


def test_setup_reports_when_the_service_does_not_answer(home, tmp_path):
    log = dienst.log_path(tmp_path)
    log.parent.mkdir(parents=True)
    log.write_text("2026-10-27 Prüfer im Hintergrund startet\nModuleNotFoundError: No module named 'x'\n", encoding="utf-8")
    assert dienst.log_tail(tmp_path)[-1] == "ModuleNotFoundError: No module named 'x'"
    with socket.socket() as server:
        server.bind(("127.0.0.1", 0))
        port = server.getsockname()[1]
    assert dienst.answers(timeout=0, url=f"https://localhost:{port}/x") == "das Zertifikat fehlt"
    if shutil.which("openssl"):
        ca, _, _ = addin.certificates()
        assert dienst.answers(timeout=0, url=f"https://localhost:{port}/x", ca=ca).startswith("keine Verbindung")


def test_busy_sees_a_running_server():
    with socket.socket() as server:
        server.bind(("127.0.0.1", 0))
        server.listen()
        port = server.getsockname()[1]
        assert dienst.busy(port)
    assert not dienst.busy(port)


def test_without_the_addin_the_service_waits(home, capsys):
    waited = []
    assert dienst.main(pause=waited.append) == 0 and "nicht eingerichtet" in capsys.readouterr().out
    assert waited == [dienst.WAIT * 10]


def test_a_new_version_restarts_the_service(monkeypatch):
    restarted = []
    monkeypatch.setattr(update, "pull", lambda: "Neue Version geladen (3 Änderungen).")
    monkeypatch.setattr(update, "sync_packages", lambda: False)
    monkeypatch.setattr(update, "drop_retired", lambda: [])
    monkeypatch.setattr(update, "record", lambda *_: None)
    monkeypatch.setattr(dienst, "restart", lambda: restarted.append(True) or True)
    assert update.main() == 0 and restarted == [True]
    monkeypatch.setattr(update, "pull", lambda: "Gegenprobe ist aktuell.")
    assert update.main() == 0 and restarted == [True]


@pytest.mark.skipif(shutil.which("openssl") is None or sys.platform == "win32", reason="der Weg mit openssl ist der vom Mac")
def test_setup_ends_with_the_service_running(home, monkeypatch, capsys):
    (home / "Library" / "Containers" / "com.microsoft.Excel").mkdir(parents=True)
    monkeypatch.setattr(addin.sys, "platform", "darwin")
    monkeypatch.setattr(addin, "_trusted", lambda cert: True)
    run = Launchctl(starts=False)
    wired(monkeypatch, run)
    monkeypatch.setattr(dienst, "answers", answering(run))
    monkeypatch.chdir(home)
    assert addin.install() == 0
    out = capsys.readouterr().out
    assert "4/4" in out and "Er startet von selbst, sobald Excel ihn braucht, also auch nach einem Neustart." in out
    assert "Terminal und Browser brauchst du dafür nicht" in out
    # So wie nach einem Neustart geladen, ohne Anstoß: Gestartet hat ihn das Anklopfen.
    assert "kickstart" not in [c[0] for c in run.calls] and run.running
    assert "Probe: auf die Anfrage hin von selbst gestartet" in dienst.log_path().read_text(encoding="utf-8")
    monkeypatch.setattr(dienst, "answers", lambda **_: "keine Verbindung (refused)")
    assert addin.install() == 1 and "antwortet nicht: keine Verbindung" in capsys.readouterr().out
    assert (addin.wef_folder() / "voiceagent-pruefer.xml").exists() and dienst.installed()
    assert addin.install(remove=True) == 0 and not dienst.installed()
    assert not (addin.wef_folder() / "voiceagent-pruefer.xml").exists()


@pytest.mark.skipif(shutil.which("openssl") is None or sys.platform == "win32", reason="der Weg mit openssl ist der vom Mac")
def test_setup_opens_full_disk_access_and_later_leaves_excel_alone(home, monkeypatch, capsys):
    (home / "Library" / "Containers" / "com.microsoft.Excel").mkdir(parents=True)
    monkeypatch.setattr(addin.sys, "platform", "darwin")
    monkeypatch.setattr(addin, "_trusted", lambda cert: True)
    monkeypatch.setattr(dienst, "answers", lambda **_: "")
    run = Launchctl()
    wired(monkeypatch, run)
    opened, real = [], addin.subprocess.run
    monkeypatch.setattr(addin.subprocess, "run", lambda command, **kw: opened.append(command) if command[0] == "open"
                        else real(command, **kw))

    def refused(*_, **__):
        raise PermissionError(1, "Operation not permitted")

    monkeypatch.setattr(addin, "tempfile", SimpleNamespace(NamedTemporaryFile=refused))
    assert addin.install() == 1
    out = capsys.readouterr().out
    assert "Festplattenvollzugriff" in out and "Befehl-Q" in out and opened == [["open", addin.FULL_DISK_ACCESS]]
    assert not dienst.installed()
    # Mit der Erlaubnis meldet es das Add-in an; danach braucht es den Ordner von Excel nicht mehr.
    monkeypatch.setattr(addin, "tempfile", tempfile)
    assert addin.install() == 0 and "3/4 Ich melde das Add-in bei Excel an" in capsys.readouterr().out
    monkeypatch.setattr(addin, "tempfile", SimpleNamespace(NamedTemporaryFile=refused))
    assert addin.install() == 0 and "3/4 Das Add-in ist bei Excel schon angemeldet." in capsys.readouterr().out
    assert addin.install(remove=True) == 0 and not (addin.folder() / "manifest.sha256").exists()


def test_state_reads_what_launchd_knows(home, monkeypatch):
    run = Launchctl()
    wired(monkeypatch, run)
    info = dienst.state()
    assert not info["plist"] and not info["loaded"] and dienst.explain(info) == "er ist nicht eingerichtet"
    dienst.install(root=home, python=sys.executable, run=run, pause=lambda _: None, start=False)
    info = dienst.state()
    # Der Zustand im Unterblock („endpoints“) zählt nicht.
    assert info["loaded"] and info["pid"] == 4242 and info["state"] == "running" and dienst.explain(info) == "er läuft"
    quiet = Launchctl(starts=False)
    wired(monkeypatch, quiet)
    dienst.install(root=home, python=sys.executable, run=quiet, pause=lambda _: None, start=False)
    # Bis Excel anklopft, läuft er nicht, und so ist es gewollt.
    info = dienst.state()
    assert info["sockets"] and dienst.ready(info) and dienst.explain(info) == "er wartet, bis Excel ihn braucht"
    # Eingerichtet, bevor macOS den Port hielt: Dann hätte er laufen sollen.
    old = {**info, "sockets": False}
    assert not dienst.ready(old) and dienst.explain(old) == "macOS hat ihn geladen, aber nie gestartet"
    assert dienst.explain({**old, "runs": 4, "exit": "1"}) == "er ist mit einem Fehler beendet (Code 1)"
    assert dienst.explain({**old, "runs": 4, "signal": "Killed: 9"}) == "er wurde beendet (Killed: 9)"
    quiet.disabled = True
    info = dienst.state()
    assert info["disabled"] and not dienst.ready(info) and dienst.explain(info) == "er ist in macOS abgeschaltet"
    # Abgeschaltet und nie gestartet: einschalten und anstoßen.
    dienst.repair(run=quiet)
    assert not quiet.disabled and quiet.running and quiet.calls[-2:] == [["enable", f"gui/501/{dienst.LABEL}"],
                                                                         ["kickstart", f"gui/501/{dienst.LABEL}"]]
    # Der Ordner umbenannt: macOS findet Python nicht mehr.
    dienst.install(root=home, python=str(home / "VoiceAgent" / ".venv" / "bin" / "python"), run=quiet,
                   pause=lambda _: None, start=False)
    assert dienst.explain(dienst.state()) == (f"{home / 'VoiceAgent' / '.venv' / 'bin' / 'python'} gibt es nicht mehr, "
                                              "ist der Ordner verschoben oder umbenannt?")
    assert not dienst.ready(dienst.state())


def test_whether_it_started_by_itself_after_the_last_reboot(home, monkeypatch):
    from datetime import datetime, timedelta

    target = dienst.plist_path()
    target.parent.mkdir(parents=True)
    target.write_bytes(b"x")
    set_up = datetime(2026, 10, 27, 15, 0)
    os.utime(target, (set_up.timestamp(), set_up.timestamp()))
    boot = datetime(2026, 10, 28, 8, 12)
    log = dienst.log_path(home)
    log.parent.mkdir(parents=True)
    assert dienst.since_boot(home, None) is None
    # Ohne Protokoll ist er seit dem Neustart nie gestartet.
    assert dienst.since_boot(home, boot) is False
    log.write_text(f"2026-10-27 15:00:02 {dienst.STARTS} (/v/python, /v).\n", encoding="utf-8")
    assert dienst.since_boot(home, boot) is False
    with log.open("a", encoding="utf-8") as file:
        file.write(f"2026-10-28 08:12:41 {dienst.STARTS} (/v/python, /v).\n")
    assert dienst.since_boot(home, boot) is True
    # Seit dem Neustart neu eingerichtet: Dann lässt sich nichts sagen.
    later = boot + timedelta(minutes=5)
    os.utime(target, (later.timestamp(), later.timestamp()))
    assert dienst.since_boot(home, boot) is None


def test_before_says_what_happened_since_the_reboot(home, monkeypatch):
    from datetime import datetime

    monkeypatch.chdir(home)
    run = Launchctl(starts=False, boot=datetime(2026, 10, 28, 8, 12))
    wired(monkeypatch, run)
    assert dienst.before() == ([], False)
    dienst.install(root=home, python=sys.executable, run=run, pause=lambda _: None, start=False)
    old = datetime(2026, 10, 27, 15, 0).timestamp()
    os.utime(dienst.plist_path(), (old, old))
    # Er startet erst, wenn Excel ihn braucht: noch nicht gelaufen ist kein Fehler.
    assert dienst.before() == (["Seit dem letzten Neustart des Macs (28.10. 08:12) hat Excel ihn noch nicht gebraucht."],
                               False)
    run.disabled = True
    lines, missed = dienst.before()
    assert missed and lines[-1] == "Eben war er nicht bereit: er ist in macOS abgeschaltet."
    # Eingerichtet, bevor macOS den Port hielt: Dann hätte er nach dem Neustart von selbst starten sollen.
    run.disabled = False
    data = plistlib.loads(dienst.plist_path().read_bytes())
    del data["Sockets"]
    dienst.plist_path().write_bytes(plistlib.dumps({**data, "RunAtLoad": True, "KeepAlive": True}))
    os.utime(dienst.plist_path(), (old, old))
    lines, missed = dienst.before()
    assert missed and lines == ["Nach dem letzten Neustart des Macs (28.10. 08:12) ist er nicht von selbst gestartet.",
                                "Eben lief er nicht: macOS hat ihn geladen, aber nie gestartet."]


@pytest.mark.skipif(shutil.which("openssl") is None or sys.platform == "win32", reason="der Weg mit openssl ist der vom Mac")
def test_setup_nudges_a_service_that_does_not_start_by_itself(home, monkeypatch, capsys):
    (home / "Library" / "Containers" / "com.microsoft.Excel").mkdir(parents=True)
    monkeypatch.setattr(addin.sys, "platform", "darwin")
    monkeypatch.setattr(addin, "_trusted", lambda cert: True)
    # macOS startet ihn auch nicht, wenn jemand anklopft.
    run = Launchctl(starts=False, demand=False)
    wired(monkeypatch, run)
    monkeypatch.chdir(home)
    monkeypatch.setattr(dienst, "answers", answering(run))
    opened, real = [], addin.subprocess.run
    monkeypatch.setattr(addin.subprocess, "run", lambda command, **kw: opened.append(command) if command[0] == "open"
                        else real(command, **kw))
    assert addin.install() == 0
    out = capsys.readouterr().out
    assert ("Von selbst startete er nicht (macOS hat ihn auf die Anfrage hin nicht gestartet). Angestoßen antwortet er "
            "jetzt.") in out
    assert "Anmeldeobjekte" in out and "Im Hintergrund erlauben" in out and opened == [["open", dienst.LOGIN_ITEMS]]
    assert [c[0] for c in run.calls][-2:] == ["enable", "kickstart"]
    # Was macOS dazu sagte, steht zum Nachlesen im Protokoll.
    log = dienst.log_path()
    text = log.read_text(encoding="utf-8")
    assert ("ExcelAddin.command: Probe: nicht von selbst gestartet (macOS hat ihn auf die Anfrage hin nicht gestartet)"
            in text)
    assert "ExcelAddin.command: \tstate = not running" in text
    # Startet er gar nicht, auch angestoßen nicht: Das Protokoll sagt, warum (ohne das, was dazugeschrieben ist).
    run.fail = {"kickstart"}
    with log.open("a", encoding="utf-8") as file:
        file.write("ModuleNotFoundError: No module named 'uvicorn'\n")
    assert addin.install() == 1
    out = capsys.readouterr().out
    assert "antwortet nicht: keine Verbindung" in out and "No module named 'uvicorn'" in out
    assert "launchctl print" not in out and "state = " not in out


@pytest.mark.skipif(shutil.which("openssl") is None or sys.platform == "win32", reason="der Weg mit openssl ist der vom Mac")
def test_setup_falls_back_to_starting_at_login_when_the_port_is_not_handed_over(home, monkeypatch, capsys):
    """Startet er, bekommt aber den Port von macOS nicht: dann wie früher, angestoßen und mit eigenem Port."""
    (home / "Library" / "Containers" / "com.microsoft.Excel").mkdir(parents=True)
    monkeypatch.setattr(addin.sys, "platform", "darwin")
    monkeypatch.setattr(addin, "_trusted", lambda cert: True)
    monkeypatch.chdir(home)
    run = Launchctl(starts=False)
    wired(monkeypatch, run)
    monkeypatch.setattr(dienst, "answers", lambda **_: "" if run.running and not dienst.state()["sockets"]
                        else "keine Verbindung (refused)")
    assert addin.install() == 0
    out = capsys.readouterr().out
    assert "Auf Anfrage startete macOS ihn nicht. Er läuft jetzt wie bisher, angestoßen" in out
    data = plistlib.loads(dienst.plist_path().read_bytes())
    assert data["RunAtLoad"] is True and data["KeepAlive"] is True and "Sockets" not in data
    assert "Rückfall auf den Start bei der Anmeldung" in dienst.log_path().read_text(encoding="utf-8")


def test_voiceagent_command_leaves_the_port_to_macos(home, monkeypatch):
    run = Launchctl(starts=False)
    wired(monkeypatch, run)
    assert not dienst.waits()
    dienst.install(root=home, python=sys.executable, run=run, pause=lambda _: None, start=False)
    assert dienst.waits() and not run.running
    run.disabled = True
    assert not dienst.waits()


@pytest.mark.skipif(sys.platform == "win32", reason="launchd gibt es nur am Mac")
def test_the_port_macos_held_open_is_taken_over():
    """launch_activate_socket gibt dem Dienst den Port, an dem Excel anklopft."""
    freed = []
    with socket.socket() as server:
        server.bind(("127.0.0.1", 0))
        server.listen()
        port = server.getsockname()[1]
        handed = (ctypes.c_int * 1)(os.dup(server.fileno()))

        def launch_activate_socket(name, fds, count):
            assert name == b"Listeners"
            fds._obj.contents = ctypes.c_int.from_buffer(handed)
            count._obj.value = 1
            return 0

        [taken] = dienst.activated(SimpleNamespace(launch_activate_socket=launch_activate_socket, free=freed.append))
        with taken, socket.create_connection(("127.0.0.1", port), timeout=2):
            assert taken.getsockname()[1] == port and taken.type == socket.SOCK_STREAM
    assert len(freed) == 1
    # Nicht von macOS gestartet (ESRCH): nichts zu übernehmen.
    assert dienst.activated(SimpleNamespace(launch_activate_socket=lambda *_: 3, free=freed.append)) == []


@pytest.mark.skipif(shutil.which("openssl") is None or sys.platform == "win32", reason="der Weg mit openssl ist der vom Mac")
def test_woken_through_the_port_it_answers_excel_there(home, monkeypatch, capsys):
    """Den Port hat macOS geöffnet: Der Dienst öffnet ihn nicht selbst, sondern antwortet dort mit https."""
    monkeypatch.chdir(home)
    monkeypatch.setenv("VOICEAGENT_NO_EXCEL", "1")
    ca, _, _ = addin.certificates()
    dienst.install(root=home, python=sys.executable, run=Launchctl(starts=False), pause=lambda _: None, start=False)
    # Gibt macOS den Port nicht her, kann er ihn auch nicht selbst öffnen: Das steht im Protokoll.
    monkeypatch.setattr(dienst, "activated", lambda: [])
    assert dienst.main() == 1 and "macOS hat den Port 8766 nicht übergeben." in capsys.readouterr().out
    server = socket.socket()
    server.bind(("127.0.0.1", 0))
    server.listen()
    port = server.getsockname()[1]
    monkeypatch.setattr(dienst, "activated", lambda: [server])
    monkeypatch.setattr(dienst, "busy", lambda *_: pytest.fail("Er klopft nicht selbst an."))
    threading.Thread(target=dienst.main, daemon=True).start()
    context = ssl.create_default_context(cafile=str(ca))
    context.verify_flags &= ~ssl.VERIFY_X509_STRICT
    for _ in range(150):
        try:
            with urllib.request.urlopen(f"https://localhost:{port}/addin/pruefer.html", context=context, timeout=2) as page:
                assert page.status == 200 and b"Gegenprobe Pr" in page.read()
                assert page.headers["Cache-Control"] == "no-cache"
            break
        except OSError:
            time.sleep(0.1)
    else:
        pytest.fail("https antwortet nicht")
    assert "geweckt über den Port" in capsys.readouterr().out
