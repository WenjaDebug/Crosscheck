"""Die Übungsmappe aus dem Seitenfenster öffnen: eine Kopie im Datenordner, geänderte Kopien bleiben stehen."""

from pathlib import Path

from test_excel_live import FakeExcel, client_with

from voiceagent import uebung

HEADERS = {"X-VoiceAgent": "1"}


def test_the_copy_is_the_practice_workbook(tmp_path):
    path = uebung.copy(tmp_path)
    assert path == tmp_path / "Übungsmappe.xlsx"
    assert path.read_bytes() == uebung.DEMO.read_bytes()


def test_an_unchanged_copy_is_opened_again(tmp_path):
    first = uebung.copy(tmp_path)
    assert uebung.copy(tmp_path) == first
    assert sorted(p.name for p in tmp_path.iterdir()) == ["Übungsmappe.xlsx"]


def test_a_changed_copy_stays_and_a_fresh_one_comes_beside_it(tmp_path):
    first = uebung.copy(tmp_path)
    first.write_bytes(b"geaendert")
    second = uebung.copy(tmp_path)
    assert second == tmp_path / "Übungsmappe 2.xlsx"
    assert first.read_bytes() == b"geaendert" and second.read_bytes() == uebung.DEMO.read_bytes()


def test_the_pane_endpoint_opens_the_copy(tmp_path, monkeypatch):
    opened: list[Path] = []
    monkeypatch.setattr(uebung, "launch", lambda path: opened.append(path) or True)
    _, client = client_with(tmp_path, FakeExcel(None))
    answer = client.post("/api/addin/uebung", json={}, headers=HEADERS).json()
    assert answer["state"] == "ok" and answer["name"] == "Übungsmappe.xlsx"
    assert opened == [tmp_path / "uebungsmappen" / "Übungsmappe.xlsx"]
    assert client.post("/api/addin/uebung", json={}).status_code == 403


def test_when_nothing_opens_it_the_pane_learns_where_it_is(tmp_path, monkeypatch):
    monkeypatch.setattr(uebung, "launch", lambda path: False)
    _, client = client_with(tmp_path, FakeExcel(None))
    answer = client.post("/api/addin/uebung", json={}, headers=HEADERS).json()
    assert answer["state"] == "closed" and answer["path"].endswith("Übungsmappe.xlsx")


def test_the_pane_no_longer_relies_on_a_download():
    pane = (Path(uebung.__file__).parent / "web" / "addin" / "pruefer.html").read_text(encoding="utf-8")
    script = (Path(uebung.__file__).parent / "web" / "addin" / "pruefer.js").read_text(encoding="utf-8")
    assert 'id="practice"' in pane and " download>" not in pane
    assert '"/api/addin/uebung"' in script
