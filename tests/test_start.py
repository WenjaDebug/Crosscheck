"""Die Startseite gehört dem Prüfer; unter „Mehr“ stehen nur noch der Excel-Helfer und der Zustand."""

import re
import sys
from pathlib import Path

import pytest
from test_excel_live import FakeExcel, client_with

WEB = Path(__file__).resolve().parents[1] / "src" / "voiceagent" / "web"


def test_the_start_page_is_the_checker_and_only_helper_and_status_remain(tmp_path, monkeypatch):
    monkeypatch.setenv("VOICEAGENT_ADDIN_DIR", str(tmp_path / "addin"))
    _, client = client_with(tmp_path, FakeExcel(None))
    start = client.get("/").text
    assert "Gegenprobe Prüfer" in start and 'id="drop"' in start and "/demo/pruefer-uebung.xlsx" in start
    more = start[start.index("<summary>Mehr</summary>"):]
    for page in ("/helfer.html", "/zustand.html", "/pruefer.html"):
        assert client.get(page).status_code == 200, page
    assert "/helfer.html" in more and "/zustand.html" in more
    for page in ("/lehrling.html", "/map.html", "/handbuch.html", "/rules.html", "/lernstand.html", "/ki.html",
                 "/daten.html", "/app.js", "/ptt.js", "/demo/notebook.html"):
        assert page not in start, page
        assert client.get(page).status_code == 404, page
    assert client.get("/api/start").json()["addin"] is False


def test_the_start_page_says_when_the_background_checker_is_not_running(tmp_path, monkeypatch):
    from voiceagent import dienst

    monkeypatch.setenv("VOICEAGENT_ADDIN_DIR", str(tmp_path / "addin"))
    # Der Weg am Mac (launchd), auch wenn die Tests unter Windows laufen; Windows prüft test_windows_fenster.py.
    monkeypatch.setattr(sys, "platform", "darwin")
    monkeypatch.setattr(dienst, "installed", lambda: True)
    monkeypatch.setattr(dienst, "state", lambda: {"plist": True, "loaded": True, "pid": None, "state": "not running",
                                                   "runs": 0, "exit": "", "signal": "", "disabled": False})
    monkeypatch.setattr(dienst, "boot_time", lambda: None)
    _, client = client_with(tmp_path, FakeExcel(None))
    data = client.get("/api/start").json()
    assert data["dienst"] and data["laeuft"] is False and data["neustart"] is None
    assert data["grund"] == "macOS hat ihn geladen, aber nie gestartet"
    script = (WEB / "start.js").read_text(encoding="utf-8")
    assert "läuft gerade nicht" in script and "Nach dem letzten Neustart des Macs" in script


def test_the_start_page_says_the_checker_starts_when_excel_needs_it(tmp_path, monkeypatch):
    from voiceagent import dienst

    monkeypatch.setenv("VOICEAGENT_ADDIN_DIR", str(tmp_path / "addin"))
    # Der Weg am Mac (launchd), auch wenn die Tests unter Windows laufen; Windows prüft test_windows_fenster.py.
    monkeypatch.setattr(sys, "platform", "darwin")
    monkeypatch.setattr(dienst, "installed", lambda: True)
    monkeypatch.setattr(dienst, "boot_time", lambda: pytest.fail("Ein Start nach dem Neustart sagt hier nichts."))
    waiting = {"plist": True, "loaded": True, "pid": None, "state": "not running", "runs": 0, "exit": "", "signal": "",
               "disabled": False, "missing": "", "sockets": True}
    monkeypatch.setattr(dienst, "state", lambda: waiting)
    _, client = client_with(tmp_path, FakeExcel(None))
    data = client.get("/api/start").json()
    assert data["laeuft"] is True and data["abruf"] is True and data["grund"] == "" and data["neustart"] is None
    waiting["disabled"] = True
    data = client.get("/api/start").json()
    assert data["laeuft"] is False and data["grund"] == "er ist in macOS abgeschaltet"
    script = (WEB / "start.js").read_text(encoding="utf-8")
    assert "sobald du ihn in Excel öffnest" in script and "ist gerade nicht bereit" in script


def test_every_page_links_back_to_a_page_that_exists():
    for page in WEB.glob("*.html"):
        for target in re.findall(r'class="back"><a href="([^"]+)"', page.read_text(encoding="utf-8")):
            path = "index.html" if target == "/" else target.lstrip("/")
            assert (WEB / path).exists(), (page.name, target)
