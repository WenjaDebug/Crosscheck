"""Excel live mitlesen: die Antworten von Excel (Mac und Windows), wo die Datei liegt,
und wie die Aufnahme damit arbeitet, ohne Sprechtaste."""

import json
import os
import shutil
import subprocess
import sys
import time
from types import SimpleNamespace

import pytest
from fastapi.testclient import TestClient
from test_excel_structure import WAHL, mappe

from voiceagent import diagnose, excel
from voiceagent.app import create_app
from voiceagent.config import Settings
from voiceagent.excel import ExcelLink, Snapshot, Status, local_path, parse_mac, parse_windows
from voiceagent.xlsx import Cell

MAC_ANSWER = {
    "running": True, "front": True, "book": "Wahl.xlsx", "path": "/Users/anna/Documents/Wahl.xlsx",
    "sheets": ["Tabelle1", "Grafik"], "sheet": "Tabelle1", "row": 2, "col": 2,
    "formulas": [["Parteien", "Anteil", ""], ["RN", "33.2", "=C3*577/100"], ["", "", ""]],
    "values": [["Parteien", "Anteil", ""], ["RN", 33.2, 191.564], ["", {"d": "20.10.2026"}, True]],
    "at": [3, 4],
}


def test_mac_answer_becomes_cells():
    snapshot, status = parse_mac(json.dumps(MAC_ANSWER))
    assert status.state == "ok"
    assert (snapshot.book, snapshot.sheet, snapshot.sheets, snapshot.front) == ("Wahl.xlsx", "Tabelle1", ["Tabelle1", "Grafik"], True)
    assert snapshot.cells[(2, 2)] == Cell("Parteien")
    assert snapshot.cells[(3, 3)] == Cell("33,2")
    assert snapshot.cells[(3, 4)] == Cell("191,564", "C3*577/100")
    assert snapshot.cells[(4, 3)] == Cell("20.10.2026") and snapshot.cells[(4, 4)] == Cell("WAHR")
    assert (4, 2) not in snapshot.cells
    assert snapshot.at == (3, 4) and snapshot.window == (2, 2, 4, 4)


def test_mac_answers_without_cells():
    assert parse_mac(json.dumps({"running": False}))[1].state == "closed"
    assert parse_mac(json.dumps({"running": True}))[1].state == "empty"
    denied = {"running": True, "error": "Error: Not authorized to send Apple events to Microsoft Excel. (-1743)"}
    assert parse_mac(json.dumps(denied))[1].state == "denied"
    assert parse_mac("kaputt")[1].state == "error"
    # Nur die Zellen fehlen: Die Datei bleibt.
    snapshot, status = parse_mac(json.dumps({"running": True, "book": "Mappe1", "error": "Error: -1728"}))
    assert status.state == "ok" and snapshot.cells is None


def test_windows_answer_becomes_cells():
    output = "\n".join([
        "running\t1", "book\tWahl.xlsx", "path\tC:\\Users\\anna\\Wahl.xlsx", "sheets\tTabelle1", "sheets\tGrafik",
        "sheet\tTabelle1", "row\t2", "col\t2", "size\t2 3",
        "cell\t1\t1\tParteien\ts:Parteien", "cell\t2\t2\t33.2\tn:33.2", "cell\t2\t3\t=C3*577/100\tn:191.564",
        "cell\t1\t3\t\td:20.10.2026 00:00", "cell\t1\t2\t\tb:0", "at\t3 4",
    ])
    snapshot, status = parse_windows(output)
    assert status.state == "ok" and snapshot.sheets == ["Tabelle1", "Grafik"]
    assert snapshot.cells == {
        (2, 2): Cell("Parteien"), (3, 3): Cell("33,2"), (3, 4): Cell("191,564", "C3*577/100"),
        (2, 4): Cell("20.10.2026"), (2, 3): Cell("FALSCH"),
    }
    assert snapshot.at == (3, 4) and snapshot.window == (2, 2, 3, 4)
    assert parse_windows("running\t0")[1].state == "closed"
    assert parse_windows("running\t1")[1].state == "empty"


class Run:
    """Statt osascript oder PowerShell: gibt der Reihe nach die Antworten zurück und merkt sich die Aufrufe."""

    def __init__(self, *answers):
        self.answers = list(answers)
        self.calls = []

    def __call__(self, command, **options):
        self.calls.append(command)
        answer = self.answers.pop(0)
        if isinstance(answer, Exception):
            raise answer
        if isinstance(answer, int):
            return SimpleNamespace(returncode=answer, stdout=b"", stderr=b"")
        if isinstance(answer, tuple):
            return SimpleNamespace(returncode=answer[0], stdout=b"", stderr=answer[1].encode())
        return SimpleNamespace(returncode=0, stdout=answer.encode(), stderr=b"")


def test_link_on_the_mac():
    run = Run(0, json.dumps(MAC_ANSWER))
    link = ExcelLink("darwin", run)
    assert link.running()
    assert run.calls[0] == ["pgrep", "-x", "Microsoft Excel"]
    snapshot = link.read()
    assert run.calls[1][:3] == ["osascript", "-l", "JavaScript"] and run.calls[1][-1] == "used"
    assert snapshot.cells[(3, 4)].formula == "C3*577/100"
    # Das ganze Blatt gelesen: Was fehlt, ist gelöscht.
    assert snapshot.window is None and link.connected and link.status.state == "ok"


def test_link_denied_busy_and_slow():
    link = ExcelLink("darwin", Run((1, "execution error: Not authorized to send Apple events to Microsoft Excel. (-1743)")))
    assert link.read() is None and link.status.state == "denied"
    late = subprocess.TimeoutExpired("osascript", 20)
    run = Run(json.dumps(MAC_ANSWER), late, json.dumps(MAC_ANSWER), late, late, late, json.dumps(MAC_ANSWER))
    link = ExcelLink("darwin", run)
    link.read()
    # Sie tippt gerade in einer Zelle: Excel antwortet erst danach.
    assert link.read() is None and link.status.state == "busy"
    assert link.read() is not None and run.calls[-1][-1] == "used"
    for _ in range(3):
        link.read()
    # Immer wieder zu langsam: ab jetzt nur noch, was im Fenster zu sehen ist.
    snapshot = link.read()
    assert run.calls[-1][-1] == "visible" and snapshot.window == (2, 2, 4, 4)
    rejected = "running\t1\nerror\tDer Aufruf wurde vom Aufgerufenen abgelehnt. (0x80010001 (RPC_E_CALL_REJECTED))"
    assert parse_windows(rejected)[1].state == "busy"
    assert parse_mac(json.dumps({"running": True, "nobook": "Error: Can't get object."}))[1].state == "empty"


def test_link_on_windows_uses_powershell_without_a_window():
    run = Run("EXCEL.EXE   1234 Console", "running\t1\nbook\tWahl.xlsx\nsheet\tTabelle1\nrow\t1\ncol\t1\nsize\t1 1\ncell\t1\t1\t\ts:x")
    link = ExcelLink("win32", run)
    assert link.running() and link.read().cells == {(1, 1): Cell("x")}
    assert run.calls[1][:4] == ["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand"]
    assert ExcelLink("linux").read() is None and not ExcelLink("linux").supported


@pytest.mark.skipif(not sys.platform.startswith("win"), reason="PowerShell gibt es nur unter Windows")
def test_powershell_script_runs():
    # Ohne Excel (wie auf dem Testrechner): Das Skript läuft durch und sagt, Excel sei zu.
    link = ExcelLink()
    assert link.read() is None and link.status.state == "closed"


@pytest.mark.skipif(shutil.which("node") is None, reason="ohne Node")
def test_mac_script_with_a_pretend_excel(tmp_path):
    """Das AppleScript (JavaScript) mit einem nachgebauten Excel, wie es macOS ausführt."""
    fake = """
    const sheet = {name: () => "Tabelle1",
      usedRange: {firstRowIndex: () => 2, firstColumnIndex: () => 2,
        formula: () => [["Parteien", "=1+1"]], value: () => [["Parteien", new Date(2026, 9, 20)]]}};
    const xl = {running: () => true, frontmost: () => false, activeSheet: sheet,
      activeWorkbook: {name: () => "Wahl.xlsx", fullName: () => "/Users/b/Wahl.xlsx", worksheets: {name: () => ["Tabelle1"]}},
      activeCell: {firstRowIndex: () => 2, firstColumnIndex: () => 3}};
    function Application(name) { return xl; }
    """
    script = tmp_path / "excel.js"
    script.write_text(fake + excel._MAC_SCRIPT + '\nconsole.log(run(["used"]));\n', encoding="utf-8")
    output = subprocess.run(["node", str(script)], capture_output=True, text=True, encoding="utf-8", timeout=30, check=True).stdout
    snapshot, status = parse_mac(output)
    assert status.state == "ok" and snapshot.front is False
    assert snapshot.cells == {(2, 2): Cell("Parteien"), (2, 3): Cell("20.10.2026", "1+1")}
    assert snapshot.at == (2, 3)


# --- Wo die Datei liegt ---


@pytest.fixture
def onedrive(tmp_path, monkeypatch):
    for variable in ("OneDrive", "OneDriveCommercial", "OneDriveConsumer"):
        monkeypatch.delenv(variable, raising=False)
    folder = tmp_path / "Library" / "CloudStorage" / "OneDrive-Persönlich" / "Dokumente" / "Uni"
    folder.mkdir(parents=True)
    path = folder / "Wahl 2024.xlsx"
    path.write_bytes(mappe(("Tabelle1", WAHL, "", {})))
    return tmp_path, path


def test_onedrive_address_is_the_local_file(onedrive):
    home, path = onedrive
    assert local_path("https://d.docs.live.net/4f2a9c/Dokumente/Uni/Wahl%202024.xlsx", home) == path
    assert local_path("https://firma-my.sharepoint.com/personal/b_firma_ch/Documents/Dokumente/Uni/Wahl%202024.xlsx", home) == path
    # Anderer Ordner in der Adresse: nach dem Namen gesucht.
    assert local_path("https://d.docs.live.net/4f2a9c/Woanders/Wahl%202024.xlsx", home) == path
    assert local_path("https://d.docs.live.net/4f2a9c/Fehlt.xlsx", home) is None
    assert local_path(str(path), home) == path
    assert local_path("Mappe1", home) is None


def test_old_mac_path(onedrive):
    _, path = onedrive
    assert excel._hfs("Macintosh HD:Users:anna:Wahl.xlsx") == "/Users/anna/Wahl.xlsx"
    assert excel._hfs(str(path)) == str(path)


def test_saved_file_is_read_once_per_save(onedrive):
    _, path = onedrive
    old = time.time() - 60
    os.utime(path, (old, old))
    link = ExcelLink("darwin", Run())
    snapshot = Snapshot("Wahl 2024.xlsx", str(path))
    book = link.book(snapshot)
    assert book.sheets[0].name == "Tabelle1" and link.file_note == ""
    assert link.book(snapshot) is book
    os.utime(path, (old + 10, old + 10))
    assert link.book(snapshot) is not book
    assert link.book(Snapshot("Mappe1", "Mappe1")) is None and link.file_note == "unsaved"
    csv = path.with_suffix(".csv")
    csv.write_text("a;b", encoding="utf-8")
    assert link.book(Snapshot("Wahl.csv", str(csv))) is None and link.file_note == "format"


# --- In der Aufnahme ---


class FakeExcel:
    """Ein Excel, das zeigt, was der Test hineinlegt."""

    supported = True
    platform = "darwin"
    interval = 0.02

    def __init__(self, snapshot: Snapshot | None, book=None):
        self.snapshot, self._book = snapshot, book
        self.open, self.connected = True, False
        self.status, self.file_note = Status("closed"), "unsaved"

    def running(self):
        return self.open

    def read(self):
        if self.snapshot is None:
            self.status = Status("denied", "-1743")
            return None
        self.connected, self.status = True, Status("ok")
        return self.snapshot

    def book(self, snapshot):
        return self._book


def wahl_cells(extra: dict | None = None) -> dict:
    cells = {(4, 2): Cell("Parteien"), (4, 3): Cell("Wähleranteil(%)"), (4, 4): Cell("Anzahl Sitze")}
    for row, (name, share) in enumerate([("RN", "33,2"), ("NFP", "28"), ("Ensemble", "20,8")], start=6):
        cells[(row, 2)], cells[(row, 3)] = Cell(name), Cell(share)
    cells.update(extra or {})
    return cells


def client_with(tmp_path, link):
    app = create_app(Settings(data_dir=tmp_path), excel=link)
    return app, TestClient(app)


def test_diagnose_says_how_excel_is():
    link = ExcelLink("darwin", Run())
    link.status = Status("denied", "-1743")
    check = diagnose._excel_check(link)
    assert not check["ok"] and "Automation" in check["text"]
    link.status = Status("ok")
    assert diagnose._excel_check(link)["ok"]

