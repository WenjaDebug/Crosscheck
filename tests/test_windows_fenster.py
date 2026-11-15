"""Das Seitenfenster unter Windows: Zertifikat ohne openssl, Vertrauen mit certutil, Anmelden bei Excel in der Registry
und der Prüfer im Hintergrund über HKCU …\\Run. Windows gibt es hier nicht: Registry, certutil, PowerShell und das
Starten von Prozessen sind nachgebaut, das Zertifikat ist echt und wird mit einer echten TLS-Verbindung geprüft."""

import os
import shutil
import socket
import ssl
import subprocess
import sys
import threading
import time
import urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path
from types import SimpleNamespace

import pytest
from test_excel_live import FakeExcel, client_with

from voiceagent import addin, dienst, windows, zertifikat

ROOT = Path(__file__).resolve().parents[1]


class FakeRegistry:
    """Wie winreg, nur HKEY_CURRENT_USER: Schlüssel mit ihren Zeichenfolgen."""

    HKEY_CURRENT_USER = "HKCU"
    KEY_READ, KEY_SET_VALUE, REG_SZ = 0x20019, 0x0002, 1

    def __init__(self, *keys):
        self.keys: dict[str, dict[str, str]] = {}
        for key in keys:
            self._create(key)

    class Key:
        def __init__(self, path):
            self.path = path

        def __enter__(self):
            return self

        def __exit__(self, *_):
            return False

    def _create(self, path):
        parts = path.split("\\")
        # Wie die echte Registry: Die Schlüssel darüber entstehen mit.
        for i in range(1, len(parts) + 1):
            self.keys.setdefault("\\".join(parts[:i]).lower(), {})

    def OpenKey(self, root, path, reserved=0, access=0):
        assert root == self.HKEY_CURRENT_USER
        if path.lower() not in self.keys:
            raise FileNotFoundError(2, "Das System kann die angegebene Datei nicht finden")
        return self.Key(path.lower())

    def CreateKeyEx(self, root, path, reserved=0, access=0):
        assert root == self.HKEY_CURRENT_USER
        self._create(path)
        return self.Key(path.lower())

    def SetValueEx(self, key, name, reserved, kind, value):
        assert kind == self.REG_SZ and isinstance(value, str)
        self.keys[key.path][name] = value

    def QueryValueEx(self, key, name):
        if name not in self.keys[key.path]:
            raise FileNotFoundError(2, "Das System kann die angegebene Datei nicht finden")
        return self.keys[key.path][name], self.REG_SZ

    def DeleteValue(self, key, name):
        if name not in self.keys[key.path]:
            raise FileNotFoundError(2, "Das System kann die angegebene Datei nicht finden")
        del self.keys[key.path][name]

    def value(self, path, name):
        return self.keys.get(path.lower(), {}).get(name)


class FakeWindows:
    """certutil, PowerShell (Prozessliste), taskkill und das Starten eines Prozesses; dazu die Stammzertifikate,
    wie ssl.enum_certificates sie liefert. refuse: In der Sicherheitswarnung wird Nein geklickt."""

    def __init__(self, refuse=False):
        self.calls, self.store, self.running, self.started = [], [], [], []
        self.refuse = refuse

    def __call__(self, command, **_):
        self.calls.append(command)
        if command[0] == "certutil":
            if self.refuse:
                return SimpleNamespace(returncode=1, stdout="CertUtil: -addstore-Befehl ist FEHLGESCHLAGEN: "
                                       "0x800704c7 (WIN32: 1223 ERROR_CANCELLED)", stderr="")
            self.store.append(Path(command[-1]).read_bytes())
            return SimpleNamespace(returncode=0, stdout="CertUtil: -addstore-Befehl wurde erfolgreich ausgeführt.", stderr="")
        if command[0] == "powershell":
            lines = [f'{pid}\t"C:\\Python313\\pythonw.exe" -m voiceagent.dienst --ordner C:\\Gegenprobe' for pid in self.running]
            # Andere Python-Programme bleiben in Ruhe.
            lines += ["77\tpython -m voiceagent", "78\t"]
            return SimpleNamespace(returncode=0, stdout="\n".join(lines) + "\n", stderr="")
        if command[0] == "taskkill":
            pids = [int(command[i + 1]) for i, arg in enumerate(command) if arg == "/PID"]
            self.running = [pid for pid in self.running if pid not in pids]
            return SimpleNamespace(returncode=0, stdout="", stderr="")
        raise AssertionError(command)

    def enum(self, name):
        assert name == "ROOT"
        return [(cert, "x509_asn", True) for cert in self.store]

    def popen(self, args, **kw):
        self.started.append((args, kw))
        self.running.append(4242 + len(self.started))
        return SimpleNamespace(pid=self.running[-1])


@pytest.fixture
def pure(tmp_path, monkeypatch):
    """Zertifikate wie unter Windows: ohne openssl."""
    monkeypatch.setenv("VOICEAGENT_ADDIN_DIR", str(tmp_path / "addin"))
    monkeypatch.setattr(addin, "_pure", lambda: True)
    monkeypatch.setattr(addin, "_openssl", lambda *a, **k: pytest.fail("openssl darf hier nicht laufen"))
    return tmp_path


# --- Zertifikat ohne openssl ---


def test_miller_rabin_knows_primes_and_carmichael_numbers():
    assert zertifikat._probably_prime(2**127 - 1) and zertifikat._probably_prime(65537)
    # 561 besteht den Fermat-Test zu jeder Basis, Miller-Rabin nicht; 2^128+1 ist keine Primzahl.
    assert not zertifikat._probably_prime(561) and not zertifikat._probably_prime(2**128 + 1)
    assert not zertifikat._probably_prime(1) and zertifikat._probably_prime(2)


def test_the_windows_way_is_taken_without_openssl(monkeypatch):
    monkeypatch.setattr(addin.sys, "platform", "win32")
    assert addin._pure()
    monkeypatch.setattr(addin.sys, "platform", "darwin")
    monkeypatch.setattr(addin.shutil, "which", lambda name: "/usr/bin/openssl")
    # Am Mac bleibt es bei openssl.
    assert not addin._pure()
    monkeypatch.setattr(addin.shutil, "which", lambda name: None)
    assert addin._pure()


def _serve_tls(cert, key, connections):
    """Ein echter TLS-Server auf 127.0.0.1, der jede Verbindung mit „ok“ beantwortet."""
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain(str(cert), str(key))
    server = socket.create_server(("127.0.0.1", 0))
    server.settimeout(10)

    def answer():
        with server:
            for _ in range(connections):
                try:
                    raw, _ = server.accept()
                except OSError:
                    return
                try:
                    with context.wrap_socket(raw, server_side=True) as tls:
                        tls.recv(10)
                        tls.sendall(b"ok")
                except (OSError, ssl.SSLError):
                    pass

    threading.Thread(target=answer, daemon=True).start()
    return server.getsockname()[1]


def _strict(ca):
    """Ein Client, der nur dieser Stelle vertraut, den Namen prüft und so streng ist wie Python ab 3.13."""
    context = ssl.create_default_context(cafile=str(ca))
    context.verify_flags |= ssl.VERIFY_X509_STRICT
    assert context.check_hostname and context.verify_mode == ssl.CERT_REQUIRED
    return context


def test_a_certificate_without_openssl_passes_a_strict_tls_client(pure):
    ca, cert, key = addin.certificates()
    assert {p.name for p in ca.parent.iterdir()} == {"ca.pem", "ca-key.pem", "cert.pem", "key.pem"}
    assert "BEGIN PRIVATE KEY" in key.read_text() and "BEGIN CERTIFICATE" in cert.read_text()
    if sys.platform != "win32":
        assert oct(key.stat().st_mode & 0o777) == "0o600" and oct((ca.parent / "ca-key.pem").stat().st_mode & 0o777) == "0o600"
    port = _serve_tls(cert, key, 3)
    context = _strict(ca)
    for name in ("localhost", "127.0.0.1"):
        with socket.create_connection(("127.0.0.1", port), timeout=10) as raw:
            with context.wrap_socket(raw, server_hostname=name) as tls:
                tls.sendall(b"hallo")
                assert tls.recv(10) == b"ok"
    # Für einen anderen Namen gilt es nicht.
    with socket.create_connection(("127.0.0.1", port), timeout=10) as raw:
        with pytest.raises(ssl.SSLCertVerificationError):
            context.wrap_socket(raw, server_hostname="example.com")


def test_a_certificate_without_openssl_reads_like_openssl_and_is_renewed(pure, monkeypatch):
    ca, cert, key = addin.certificates()
    expires = addin._expires(cert)
    assert timedelta(days=addin.CERT_DAYS - 1) < expires - datetime.now(timezone.utc) <= timedelta(days=addin.CERT_DAYS)
    assert addin._linked(cert) and zertifikat.expires(ca).year >= datetime.now().year + 9
    authority, first = ca.read_bytes(), cert.read_bytes()
    # Noch lange gültig: Es bleibt.
    assert addin.ready() == (cert, key) and cert.read_bytes() == first
    # Bald abgelaufen: ein neues Zertifikat von derselben Stelle, der Windows schon vertraut.
    monkeypatch.setattr(addin, "RENEW_DAYS", addin.CERT_DAYS + 1)
    assert addin.certificates() == (ca, cert, key)
    assert cert.read_bytes() != first and ca.read_bytes() == authority
    port = _serve_tls(cert, key, 1)
    with socket.create_connection(("127.0.0.1", port), timeout=10) as raw:
        with _strict(ca).wrap_socket(raw, server_hostname="localhost") as tls:
            tls.sendall(b"x")
            assert tls.recv(10) == b"ok"


@pytest.mark.skipif(shutil.which("openssl") is None, reason="nur zum Vergleich mit openssl")
def test_openssl_agrees_with_the_certificate_without_openssl(pure, tmp_path):
    ca, cert, key = addin.certificates()
    done = subprocess.run(["openssl", "verify", "-x509_strict", "-CAfile", str(ca), str(cert)], capture_output=True, text=True)
    assert done.returncode == 0, done.stderr
    # Neuere openssl schreiben „CN=…“, ältere „CN = …“.
    text = subprocess.run(["openssl", "x509", "-noout", "-text", "-in", str(cert)], capture_output=True, text=True).stdout
    text = text.replace(" = ", "=")
    for wanted in ("CA:FALSE", "Digital Signature, Key Encipherment", "TLS Web Server Authentication",
                   "DNS:localhost, IP Address:127.0.0.1", "Subject Key Identifier", "Authority Key Identifier",
                   "Issuer: CN=VoiceAgent lokal (nur dieser Rechner)", "Subject: CN=localhost", "sha256WithRSAEncryption"):
        assert wanted in text, wanted
    authority = subprocess.run(["openssl", "x509", "-noout", "-text", "-in", str(ca)], capture_output=True, text=True).stdout
    assert "CA:TRUE" in authority and "Certificate Sign, CRL Sign" in authority and "critical" in authority
    assert subprocess.run(["openssl", "rsa", "-check", "-noout", "-in", str(key)], capture_output=True).returncode == 0
    # Umgekehrt liest der Weg ohne openssl auch, was openssl angelegt hat (etwa ein Zertifikat vom Mac).
    made = tmp_path / "openssl"
    made.mkdir()
    subprocess.run(["openssl", "req", "-x509", "-new", "-newkey", "rsa:2048", "-nodes", "-sha256", "-days", "30",
                    "-keyout", "k.pem", "-out", "c.pem", "-subj", "/CN=localhost"], cwd=made, capture_output=True, check=True)
    end = subprocess.run(["openssl", "x509", "-enddate", "-noout", "-in", "c.pem"], cwd=made, capture_output=True, text=True).stdout
    assert zertifikat.expires(made / "c.pem") == datetime.fromtimestamp(
        ssl.cert_time_to_seconds(end.strip().partition("=")[2]), timezone.utc)
    loaded = zertifikat.Key.load(zertifikat.unpem((made / "k.pem").read_text()))
    assert loaded.p * loaded.q == loaded.n and pow(pow(5, loaded.e, loaded.n), loaded.d, loaded.n) == 5


def test_uvicorn_serves_the_pane_with_the_certificate_without_openssl(pure, tmp_path):
    ca, cert, key = addin.certificates()
    app, _ = client_with(tmp_path, FakeExcel(None))
    with socket.socket() as probe:
        probe.bind(("127.0.0.1", 0))
        port = probe.getsockname()[1]
    addin.serve(app, cert, key, port)
    for _ in range(100):
        try:
            with urllib.request.urlopen(f"https://localhost:{port}/addin/pruefer.html", context=_strict(ca), timeout=2) as response:
                assert response.status == 200 and b"Gegenprobe" in response.read()
            break
        except OSError:
            time.sleep(0.1)
    else:
        pytest.fail("https antwortet nicht")


# --- Einrichten unter Windows ---


@pytest.fixture
def win(pure, monkeypatch):
    """Ein Windows-PC mit Excel, auf dem noch nichts eingerichtet ist."""
    reg = FakeRegistry(windows.OFFICE_KEY + r"\Excel")
    system = FakeWindows()
    monkeypatch.setattr(dienst, "busy", lambda port, host="127.0.0.1": False)
    monkeypatch.setattr(dienst, "answers", lambda **_: "" if system.running else "keine Verbindung (refused)")
    python = pure / ".venv" / "Scripts" / "python.exe"
    python.parent.mkdir(parents=True)
    python.write_bytes(b"")
    (python.parent / "pythonw.exe").write_bytes(b"")
    monkeypatch.setattr(windows.sys, "executable", str(python))
    return SimpleNamespace(reg=reg, system=system, root=pure, python=python,
                           install=lambda remove=False: addin.install_windows(remove, reg=reg, run=system, popen=system.popen,
                                                                              enum=system.enum, root=pure))


def test_setup_on_windows_trusts_registers_and_starts_in_the_background(win, capsys):
    assert win.install() == 0
    out = capsys.readouterr().out
    for step in ("1/4", "2/4 Jetzt fragt Windows einmal", "Klick in der Sicherheitswarnung auf Ja", "3/4 Ich melde",
                 "4/4", "Er antwortet", "Excel ganz schließen", "Entwickler-Add-ins", "Internet Explorer",
                 "Dieses Fenster und den Browser brauchst du dafür nicht."):
        assert step in out, step
    assert "Mac" not in out and "Befehl-Q" not in out
    # Vertrauen nur für diesen Benutzer, ohne Adminrechte, mit der Binärform der eigenen Stelle.
    certutil = [c for c in win.system.calls if c[0] == "certutil"]
    assert certutil == [["certutil", "-user", "-addstore", "Root", str(addin.folder() / "ca.cer")]]
    assert win.system.store == [zertifikat.der(addin.folder() / "ca.pem")]
    # Bei Excel angemeldet: Name = Add-in-ID, Wert = Pfad des Manifests im eigenen Ordner.
    target = addin.folder() / "voiceagent-pruefer.xml"
    assert win.reg.value(windows.WEF_KEY, addin.ADDIN_ID) == str(target)
    assert target.read_text(encoding="utf-8") == addin.manifest()
    # Bei jeder Anmeldung ohne Fenster, im Ordner von Gegenprobe; und gleich jetzt gestartet, frei vom Fenster.
    line = win.reg.value(windows.RUN_KEY, windows.RUN_NAME)
    assert line == subprocess.list2cmdline([str(win.python.parent / "pythonw.exe"), "-m", "voiceagent.dienst",
                                            "--ordner", str(win.root.resolve())])
    (args, kw), = win.system.started
    assert args[0].endswith("pythonw.exe") and args[1:] == ["-m", "voiceagent.dienst", "--ordner", str(win.root.resolve())]
    assert kw["creationflags"] & windows.DETACHED_PROCESS and kw["stdout"] == subprocess.DEVNULL

    # Noch einmal: nichts fragt, die alte Fassung wird beendet und neu gestartet.
    assert win.install() == 0
    out = capsys.readouterr().out
    assert "2/4 Windows vertraut dem Zertifikat schon." in out and "3/4 Das Add-in ist bei Excel schon angemeldet." in out
    assert len([c for c in win.system.calls if c[0] == "certutil"]) == 1
    assert [c[:3] for c in win.system.calls if c[0] == "taskkill"] == [["taskkill", "/F", "/PID"]]
    assert len(win.system.started) == 2 and len(win.system.running) == 1

    # Entfernen: aus Excel und aus dem Autostart, der Prüfer läuft nicht mehr; das Zertifikat bleibt (und das sagt es).
    assert win.install(remove=True) == 0
    out = capsys.readouterr().out
    assert "aus Excel entfernt" in out and "Benutzerzertifikate verwalten" in out
    assert win.reg.value(windows.WEF_KEY, addin.ADDIN_ID) is None and win.reg.value(windows.RUN_KEY, windows.RUN_NAME) is None
    assert not target.exists() and not win.system.running and win.system.store
    # Zweimal entfernen schadet nicht.
    assert win.install(remove=True) == 0


def test_setup_on_windows_says_when_the_security_warning_was_declined(win, capsys):
    win.system.refuse = True
    assert win.install() == 1
    out = capsys.readouterr().out
    assert "Das hat nicht geklappt" in out and "ExcelAddin.bat" in out and "ERROR_CANCELLED" in out
    assert win.reg.value(windows.WEF_KEY, addin.ADDIN_ID) is None and not win.system.started


def test_setup_on_windows_says_when_the_background_does_not_answer(win, monkeypatch, capsys):
    monkeypatch.setattr(dienst, "answers", lambda **_: "keine Verbindung (refused)")
    log = dienst.log_path(win.root)
    log.parent.mkdir(parents=True, exist_ok=True)
    log.write_text("2026-11-10 09:00:00 Prüfer im Hintergrund startet (pythonw).\nModuleNotFoundError: uvicorn\n", encoding="utf-8")
    assert win.install() == 1
    out = capsys.readouterr().out
    assert "antwortet nicht: keine Verbindung" in out and "ModuleNotFoundError: uvicorn" in out and "Bildschirmfoto" in out


def test_setup_on_windows_waits_for_voiceagent_bat_to_give_up_the_port(win, monkeypatch, capsys):
    waited = []
    monkeypatch.setattr(windows.time, "sleep", waited.append)
    monkeypatch.setattr(dienst, "busy", lambda port, host="127.0.0.1": True)
    assert win.install() == 0
    out = capsys.readouterr().out
    assert "Port 8766 ist belegt. Läuft noch VoiceAgent.bat?" in out and "starte VoiceAgent.bat, solange" in out
    assert not win.system.started and waited


def test_setup_on_windows_without_excel_still_registers(win, capsys):
    win.reg.keys.clear()
    assert win.install() == 0
    out = capsys.readouterr().out
    assert "noch nicht gefunden" in out and win.reg.value(windows.WEF_KEY, addin.ADDIN_ID)


def test_setup_picks_the_windows_way_and_other_systems_are_told(monkeypatch, capsys):
    calls = []
    monkeypatch.setattr(addin, "install_windows", lambda remove=False: calls.append(remove) or 0)
    monkeypatch.setattr(addin.sys, "platform", "win32")
    assert addin.install() == 0 and addin.install(remove=True) == 0 and calls == [False, True]
    monkeypatch.setattr(addin.sys, "platform", "linux")
    assert addin.install() == 1 and "nur am Mac und unter Windows" in capsys.readouterr().out


def test_trust_is_read_from_the_root_store(pure):
    ca, _, _ = addin.certificates()
    other = FakeWindows()
    assert not windows.trusted(ca, other.enum)
    other.store.append(b"\x30\x00")
    assert not windows.trusted(ca, other.enum)
    other.store.append(zertifikat.der(ca))
    assert windows.trusted(ca, other.enum)
    assert len(zertifikat.thumbprint(ca)) == 40


# --- Der Prüfer im Hintergrund unter Windows ---


def test_only_the_background_checker_is_stopped():
    system = FakeWindows()
    system.running = [11, 12]
    assert windows.processes(system) == [11, 12]
    # PowerShell bekommt das Skript verschlüsselt, ohne Ärger mit Anführungszeichen.
    assert system.calls[0][:4] == ["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand"]
    assert windows.stop(system) and system.running == []
    assert system.calls[-1] == ["taskkill", "/F", "/PID", "11", "/PID", "12"]
    assert not windows.stop(system)


def test_a_new_version_restarts_the_background_on_windows(monkeypatch):
    reg, system = FakeRegistry(), FakeWindows()
    monkeypatch.setattr(windows.sys, "platform", "win32")
    monkeypatch.setattr(dienst, "busy", lambda port, host="127.0.0.1": False)
    assert not windows.restart(reg, system, system.popen)
    windows.write(windows.RUN_KEY, windows.RUN_NAME, '"C:\\G\\.venv\\Scripts\\pythonw.exe" -m voiceagent.dienst --ordner C:\\G', reg)
    # Eingerichtet, läuft aber nicht: Er startet bei der nächsten Anmeldung ohnehin mit der neuen Version.
    assert not windows.restart(reg, system, system.popen)
    system.running = [31]
    assert windows.restart(reg, system, system.popen)
    # Genau so, wie Windows ihn bei der Anmeldung startet.
    assert system.started[-1][0] == windows.read(windows.RUN_KEY, windows.RUN_NAME, reg) and 31 not in system.running
    restarted = []
    monkeypatch.setattr(windows, "restart", lambda: restarted.append(True) or True)
    assert dienst.restart() and restarted == [True]


def test_starting_falls_back_when_the_window_may_not_let_go():
    started = []

    def popen(args, **kw):
        if kw["creationflags"] & windows.CREATE_BREAKAWAY_FROM_JOB:
            raise PermissionError(5, "Zugriff verweigert")
        started.append(kw["creationflags"])

    assert windows.start(["pythonw.exe"], popen) == "" and started == [windows.DETACHED_PROCESS | windows.CREATE_NEW_PROCESS_GROUP]

    def never(args, **kw):
        raise OSError(2, "nicht gefunden")

    assert "nicht gefunden" in windows.start(["pythonw.exe"], never)


def test_in_the_background_it_writes_its_log_and_starts_again_after_a_crash(tmp_path, monkeypatch):
    monkeypatch.chdir(ROOT)
    monkeypatch.setattr(sys, "stdout", sys.stdout)
    monkeypatch.setattr(sys, "stderr", sys.stderr)
    monkeypatch.setattr(os, "environ", {k: v for k, v in os.environ.items() if k != "VOICEAGENT_NO_EXCEL"})
    runs, pauses = [], []

    def run():
        runs.append((Path.cwd(), os.environ.get("VOICEAGENT_NO_EXCEL")))
        print(f"Lauf {len(runs)}")
        if len(runs) == 1:
            raise RuntimeError("abgestürzt")
        if len(runs) == 2:
            raise SystemExit(1)
        return 0

    try:
        assert dienst.background(tmp_path, pause=pauses.append, run=run) == 0
    finally:
        sys.stdout.close()
    assert runs == [(tmp_path, "1")] * 3 and pauses == [dienst.THROTTLE] * 2
    text = dienst.log_path(tmp_path).read_text(encoding="utf-8")
    assert "Lauf 1" in text and "Abgebrochen: RuntimeError('abgestürzt')" in text and "Lauf 3" in text


def test_the_command_line_of_the_service(monkeypatch, tmp_path):
    called = []
    monkeypatch.setattr(dienst, "main", lambda: called.append("main") or 0)
    monkeypatch.setattr(dienst, "background", lambda folder: called.append(folder) or 0)
    # So startet ihn launchd am Mac: ohne Angaben, wie bisher.
    assert dienst.cli([]) == 0 and called == ["main"]
    assert dienst.cli(["--ordner", str(tmp_path)]) == 0 and called == ["main", tmp_path]


def test_the_start_page_on_windows(tmp_path, monkeypatch):
    from voiceagent import app as app_module

    monkeypatch.setenv("VOICEAGENT_ADDIN_DIR", str(tmp_path / "addin"))
    _, client = client_with(tmp_path, FakeExcel(None))
    monkeypatch.setattr(app_module.sys, "platform", "win32")
    monkeypatch.setattr(windows, "installed", lambda reg=None: False)
    data = client.get("/api/start").json()
    assert data["windows"] is True and data["dienst"] is False and data["mac"] is False
    monkeypatch.setattr(windows, "installed", lambda reg=None: True)
    monkeypatch.setattr(windows, "processes", lambda run=None: [])
    data = client.get("/api/start").json()
    assert data["dienst"] and data["laeuft"] is False and "Windows" in data["grund"]
    monkeypatch.setattr(windows, "processes", lambda run=None: [4242])
    assert client.get("/api/start").json()["laeuft"] is True
    script = (ROOT / "src" / "voiceagent" / "web" / "start.js").read_text(encoding="utf-8")
    assert "ExcelAddin.bat doppelklicken" in script and "nur am Mac und unter Windows" in script
    assert "bisher nur am Mac" not in script


def test_excel_addin_bat_is_there_and_keeps_its_window_open():
    data = (ROOT / "ExcelAddin.bat").read_bytes()
    # cmd liest .bat in der Codepage der Konsole: nur ASCII, mit Windows-Zeilenenden.
    assert data.isascii() and b"\r\n" in data and b"\n" not in data.replace(b"\r\n", b"")
    text = data.decode("ascii")
    assert "install-windows.ps1" in text and "-m voiceagent.addin %*" in text and "\r\npause\r\n" in text
