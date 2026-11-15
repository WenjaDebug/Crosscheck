"""Excel-Prüfer: die offene Mappe nachrechnen, ohne KI und ohne Aufnahme."""

from test_excel_live import FakeExcel, client_with, wahl_cells

from voiceagent.excel import Snapshot
from voiceagent.pruefer import Checker
from voiceagent.xlsx import Cell


def sheet_with_mistakes() -> dict:
    cells = wahl_cells()
    for row, (formula, value) in enumerate([("C6/100*577", "191,564"), ("C7/100*577", "161,56"), ("C8/100*577", "120,016")], start=6):
        cells[(row, 4)] = Cell(value, formula)
    cells[(9, 2)], cells[(9, 3)], cells[(9, 4)] = Cell("Rest"), Cell("18"), Cell("103,86", "C9/100*577")
    cells[(11, 4)] = Cell("472,576", "SUM(D6:D8)")  # vergisst „Rest“
    cells[(12, 3)] = Cell("#DIV/0!", "C6/0")
    return cells


def test_checker_finds_what_python_can_recompute():
    link = FakeExcel(Snapshot("Wahl.xlsx", "/Users/b/Wahl.xlsx", "Tabelle1", ["Tabelle1"], sheet_with_mistakes()))
    result = Checker(link).check()
    assert result["state"] == "ok" and result["book"] == "Wahl.xlsx" and result["sheet"] == "Tabelle1"
    texts = [h["text"] for h in result["hints"]]
    assert any("#DIV/0!" in t for t in texts)
    assert any("Die Summe in D11" in t and "Rest" in t for t in texts)
    assert any("577" in t and "fest" in t for t in texts)
    [error] = [h for h in result["hints"] if h["key"] == "error:C12"]
    assert error["level"] == "fehler" and "teilt durch die Zahl 0" in error["text"]
    levels = [h["level"] for h in result["hints"]]
    assert result["text"] == f"{levels.count('fehler')} Fehler und {levels.count('pruefen')} Stellen zum Prüfen."


def test_the_count_says_how_bad_it_is():
    """Drei Fehler und eine offene Frage sind nicht „4 Stellen zum Prüfen“."""
    from voiceagent.pruefer import tally

    assert tally([]) == "Mir fällt nichts auf."
    assert tally([{"level": "fehler"}]) == "1 Fehler."
    assert tally([{"level": "fehler"}] * 3 + [{"level": "pruefen"}]) == "3 Fehler und 1 Stelle zum Prüfen."
    assert tally([{"level": "pruefen"}] * 2 + [{"level": "info"}]) == "2 Stellen zum Prüfen und 1 Hinweis."
    assert tally([{"level": "fehler"}, {"level": "pruefen"}, {"level": "info"}, {"level": "info"}]) == \
        "1 Fehler, 1 Stelle zum Prüfen und 2 Hinweise."


def test_checker_says_what_is_missing():
    link = FakeExcel(None)
    assert Checker(link).check()["text"].startswith("Ich darf Excel nicht lesen.")
    link.open = False
    assert Checker(link).check() == {"state": "closed", "text": "Öffne eine Mappe in Excel, dann prüfe ich sie."}
    assert Checker(None).check()["state"] == "off"


def test_clean_sheet_has_nothing_to_say(tmp_path):
    cells = wahl_cells()
    for row in (6, 7, 8):
        cells[(row, 4)] = Cell("1", f"C{row}/C{row}")
    link = FakeExcel(Snapshot("Ok.xlsx", "/Users/b/Ok.xlsx", "Tabelle1", ["Tabelle1"], cells))
    _, client = client_with(tmp_path, link)
    result = client.get("/api/pruefen").json()
    assert result["state"] == "ok" and result["hints"] == [] and result["text"] == "Mir fällt nichts auf."
    assert client.get("/pruefer.html").status_code == 200
