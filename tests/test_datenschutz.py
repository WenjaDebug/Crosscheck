"""Was Gegenprobe über Grenzen und Datenschutz verspricht, muss stimmen: Nach einer Prüfung steht kein Inhalt der
Mappe im Protokoll oder auf der Festplatte, ein Prüfbericht verschwindet nach einem Tag, und der Hinweis ist da."""

import logging
import os
import time
from pathlib import Path

from test_addin import payload
from test_excel_live import FakeExcel, client_with

from voiceagent import __version__, addin, bericht

SECRET = "Zürcher-Geheim"
WEB = Path(__file__).resolve().parents[1] / "src" / "voiceagent" / "web"


def _secret_book() -> dict:
    body = payload()
    sheet = body["sheets"][0]
    sheet["values"][1][0] = SECRET
    body["book"] = f"{SECRET}.xlsx"
    return body


def test_a_check_leaves_nothing_of_the_book_in_the_log_or_on_disk(tmp_path, caplog, capsys):
    caplog.set_level(logging.DEBUG)
    _, client = client_with(tmp_path, FakeExcel(None))
    headers = {"X-VoiceAgent": "1"}
    assert client.post("/api/addin/pruefen", json=_secret_book(), headers=headers).json()["state"] == "ok"
    client.post("/api/addin/herleitung", json={"sheet": "Wahl", "cell": "A3"}, headers=headers)
    # Pivot-Tabellen schrieb der Prüfer früher mit ihren Feldnamen ins Protokoll.
    addin._pivots([{"name": SECRET, "rows": [SECRET], "columns": [], "filters": [], "values": [], "skipped": {}}])
    out, err = capsys.readouterr()
    assert SECRET not in caplog.text and SECRET not in out and SECRET not in err
    for path in tmp_path.rglob("*"):
        if path.is_file():
            assert SECRET.encode() not in path.read_bytes(), path


def test_a_report_is_deleted_after_a_day(tmp_path):
    old = tmp_path / "20261101-120000-0123abcd-pruefbericht.html"
    new = tmp_path / "20261110-120000-4567abcd-pruefbericht-ohne-daten.html"
    other = tmp_path / "eigene-notiz.html"
    for path in (old, new, other):
        path.write_text("x", encoding="utf-8")
    now = time.time()
    os.utime(old, (now - 2 * 24 * 3600, now - 2 * 24 * 3600))
    bericht.tidy(tmp_path, now)
    assert not old.exists() and new.exists() and other.exists()
    bericht.tidy(tmp_path / "gibt-es-nicht", now)


def test_the_pane_shows_limits_and_privacy_with_its_version(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    page = client.get("/addin/pruefer.html").text
    assert f"Version {__version__}" in page and "__VERSION__" not in page
    for said in ("Gegenprobe kann sich irren", "Deine Daten bleiben bei dir", "Ohne Gewähr", 'id="about-open"'):
        assert said in page
    start = (WEB / "index.html").read_text(encoding="utf-8")
    assert "Gegenprobe kann sich irren" in start and "Ohne Gewähr" in start
