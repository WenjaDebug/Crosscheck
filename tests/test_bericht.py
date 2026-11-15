"""Prüfbericht zum Weitergeben: mit Werten, und ohne Daten (keine Werte, Namen, Blatt- und Mappennamen)."""

from test_addin import payload
from test_excel_live import FakeExcel, client_with

from voiceagent import bericht


def make(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    headers = {"X-VoiceAgent": "1"}
    return client, headers


def test_report_with_and_without_data(tmp_path):
    client, headers = make(tmp_path)
    assert client.post("/api/addin/bericht", json={"daten": True}, headers=headers).json()["state"] == "again"
    secret = payload(book="Befunde Müller.xlsx")
    secret["sheets"][0]["name"] = "Patient Huber"
    secret["active"] = "Patient Huber"
    client.post("/api/addin/pruefen", json=secret, headers=headers)
    full = client.post("/api/addin/bericht", json={"daten": True}, headers=headers).json()
    page = client.get(full["url"]).text
    assert full["url"].startswith("/berichte/") and "Befunde Müller.xlsx" in page and "Patient Huber" in page
    assert "#DIV/0!" in page and "C8" in page and "Fehler" in page
    plain = client.post("/api/addin/bericht", json={"daten": False}, headers=headers).json()
    assert plain["url"].endswith("-ohne-daten.html")
    page = client.get(plain["url"]).text
    for private in ("Müller", "Huber", "RN", "Ensemble", "577", "191", "Wahl"):
        assert private not in page, private
    assert "Blatt 1" in page and "C8" in page and "Fehlerwert in einer Zelle" in page
    assert client.post("/api/addin/bericht", json={}).status_code == 403


def test_report_files_are_checked_and_pruned(tmp_path):
    client, _ = make(tmp_path)
    assert client.get("/berichte/../ki.json").status_code == 404
    assert client.get("/berichte/20261027-010203-pruefbericht.html").status_code == 404
    from datetime import datetime, timedelta

    folder = tmp_path / "berichte"
    start = datetime(2026, 10, 27, 1, 0, 0)
    for i in range(bericht.KEEP + 5):
        bericht.save(folder, "<p>x</p>", i % 2 == 0, start + timedelta(seconds=i))
    assert len(list(folder.iterdir())) == bericht.KEEP


def test_every_kind_has_a_plain_name():
    import re
    from pathlib import Path

    names = ("analyse.py", "auswertung.py", "task.py", "nachschlagen.py", "summen.py", "statistik.py", "statistiktests.py",
             "zeitraeume.py", "pivotdiagramm.py", "bereiche.py", "quoten.py", "alstext.py")
    source = "".join((Path(bericht.__file__).parent / name).read_text(encoding="utf-8") for name in names)
    kinds = set(re.findall(r'f"([a-z]+):\{', source)) | set(re.findall(r'_finding\("([a-z]+)"', source))
    # Arten, deren Schlüssel aus einer Variablen entsteht (kind = "points").
    kinds |= set(re.findall(r'level, kind = .*, "([a-z]+)"$', source, re.MULTILINE))
    assert kinds <= set(bericht.KINDS), kinds - set(bericht.KINDS)
