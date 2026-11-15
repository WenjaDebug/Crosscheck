"""Große Mappen: gleiche Ergebnisse, nur schneller, und ehrlich, wenn ein Blatt zu groß ist.

Zerlegte Formeln aus Vorlagen, abgeschnittene Blätter (Seitenfenster und Datei), unveränderte Blätter nur mit
Stempel, große übrige Blätter im Hintergrund und Fehlalarme, die erst auf großen Mappen auffielen."""

import time

import pytest
from test_excel_live import FakeExcel, client_with
from test_workbook import xlsx

from voiceagent import addin, analyse, formel, pruefer, xlsx as xlsx_module
from voiceagent.einsetzen import applied
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Cells, Sheet, cut_of, read_xlsx

HEADERS = {"X-VoiceAgent": "1"}


def grid(rows):
    """Zellen ab A1: Text, Zahl oder (Formel mit „=“, Wert wie Excel ihn zeigt)."""
    cells = {}
    for r, row in enumerate(rows, 1):
        for c, value in enumerate(row, 1):
            if value is None or value == "":
                continue
            if isinstance(value, tuple):
                cells[(r, c)] = Cell(_text(value[1]), value[0][1:])
            else:
                cells[(r, c)] = Cell(_text(value))
    return cells


def _text(value):
    return repr(value).replace(".", ",") if isinstance(value, float) else str(value)


def hints(rows, name="T"):
    return pruefer.sheet_result("B", name, {name: grid(rows)})


# --- Formeln aus Vorlagen ---


@pytest.mark.parametrize("formulas", [
    ["D2*1.081", "D3*1.081", "D17*1.081", "$D$2*1.081"],
    ["SUMIFS(Daten!D:D,Daten!B:B,A2)", "SUMIFS(Daten!D:D,Daten!B:B,A3)", "SUMIFS(Daten!D2:D20,Daten!B2:B20,A4)"],
    ["IF(B2>0,B2/C2,\"B2\")", "IF(B3>0,B3/C3,\"B3\")", "IF(B4>=0,B4/C4,\"x\")"],
    ["'Daten 2024'!A1+A2", "'Daten 2024'!A9+A10", "Tabelle2!A9+A10"],
    ["VLOOKUP(A2,Tarife!$A$2:$B$9,2,FALSE)", "VLOOKUP(A3,Tarife!$A$2:$B$9,2,FALSE)", "VLOOKUP(A4,Tarife!$A$2:$B$9,2,0)"],
    ["SUM(2:2)", "SUM(3:3)", "SUM(B:B)-B1", "C2/(B2/100)^2", "C3/(B3/100)^2", "-B2%"],
])
def test_formulas_from_a_template_are_the_same_as_read_on_their_own(formulas):
    """Heruntergezogene Formeln gleicher Form zerlegt der Prüfer einmal und setzt nur die Adressen ein."""
    with formel.remembered():
        for text in formulas * 2:
            assert formel.parse(text) == formel._Parser(text).parse(), text


def test_a_formula_that_cannot_be_read_stays_unreadable_from_a_template():
    with formel.remembered():
        formel.parse("SUM(B2:B9)")
        for text in ("SUM(B2:B9", "B2+", "SUM(B2:B9)+)"):
            with pytest.raises(formel.Unknown):
                formel.parse(text)


# --- Abgeschnittene Blätter ---


def daten(rows=10):
    values = [["Fall", "Region", "Produkt", "Umsatz"]] + [[24000 + r, "Ost" if r % 2 else "West", "Maus", r * 10]
                                                          for r in range(2, rows + 2)]
    return {"name": "Daten", "row": 1, "col": 1, "values": values, "formulas_at": [], "kinds_at": [], "stamp": "d1"}


def auswertung(total):
    values = [["Kennzahl", "Wert"], ["Umsatz Summe", total]]
    return {"name": "Auswertung", "row": 1, "col": 1, "values": values, "formulas_at": [[1, 1, "=SUM(Daten!D:D)"]],
            "kinds_at": [], "stamp": "a1"}


def test_the_pane_reads_whole_rows_and_says_where_it_stopped(monkeypatch):
    monkeypatch.setattr(addin, "MAX_CELLS", 30)
    _, _, sheets = addin.sheets_from({"book": "B", "active": "Daten", "sheets": [daten()]})
    cells = sheets[0].cells
    # 4 Zellen je Zeile: Zeile 8 hätte 32 Zellen ergeben.
    assert cut_of(cells) == 7 and max(r for r, _ in cells) == 7 and len(cells) == 28
    # Hat das Fenster selbst nur den oberen Teil gelesen, gilt das auch ohne Grenze hier.
    monkeypatch.undo()
    item = daten()
    item["values"], item["cut"] = item["values"][:6], True
    _, _, sheets = addin.sheets_from({"book": "B", "active": "Daten", "sheets": [item]})
    assert cut_of(sheets[0].cells) == 6


def test_the_active_sheet_gets_its_cells_first(monkeypatch):
    monkeypatch.setattr(addin, "MAX_BOOK_CELLS", 30)
    payload = {"book": "B", "active": "Auswertung", "sheets": [daten(), auswertung(650)]}
    _, _, sheets = addin.sheets_from(payload)
    assert [s.name for s in sheets] == ["Daten", "Auswertung"]
    assert cut_of(sheets[1].cells) is None and len(sheets[1].cells) == 4
    assert cut_of(sheets[0].cells) == 6 and len(sheets[0].cells) == 24


def test_nothing_is_said_about_formulas_that_reach_below_the_cut(monkeypatch):
    """„Excel zeigt etwas anderes“ wäre falsch: Excel rechnet mit allen Zeilen, der Prüfer kennt nur die oberen."""
    full = sum(r * 10 for r in range(2, 12))
    _, _, sheets = addin.sheets_from({"book": "B", "active": "Auswertung", "sheets": [daten(), auswertung(full + 5)]})
    cells_of = {s.name: s.cells for s in sheets}
    assert [h["cell"] for h in pruefer.sheet_result("B", "Auswertung", cells_of)["hints"]
            if h["key"].startswith("recalc:")] == ["B2"]
    monkeypatch.setattr(addin, "MAX_BOOK_CELLS", 30)
    _, _, sheets = addin.sheets_from({"book": "B", "active": "Auswertung", "sheets": [daten(), auswertung(full + 5)]})
    done = pruefer.sheet_result("B", "Auswertung", {s.name: s.cells for s in sheets})
    assert not [h for h in done["hints"] if "B2" in h["cell"]]
    assert done["cut"] == ("„Daten“ ist sehr groß: Davon lese ich nur die Zeilen bis 6. Eine Formel hier rechnet mit "
                           "Zeilen weiter unten, die rechne ich nicht nach.")


def test_a_cut_own_sheet_is_checked_up_to_its_last_read_row():
    rows = [["Fall", "Umsatz", "Mit MwSt", "", "Total"]]
    for r in range(2, 12):
        rows.append([f"F{r}", r * 10, (f"=B{r}*1.081", r * 10 * 1.081)])
    rows[1].extend(["", ("=SUM(B:B)", 99999)])
    rows[6][2] = ("=B5*1.081", 50 * 1.081)  # C7 rechnet mit der Zeile darüber
    cells = grid(rows)
    cut = Cells({p: c for p, c in cells.items() if p[0] <= 8}, 8)
    done = pruefer.sheet_result("B", "T", {"T": cut})
    keys = [h["key"] for h in done["hints"]]
    assert any(k.startswith("pattern:C7") or "C7" in k for k in keys), keys
    assert not [h for h in done["hints"] if "E2" in h["cell"]]
    assert done["cut"].startswith("Das Blatt ist sehr groß: Ich prüfe die Zeilen bis 8.")


def test_a_file_reads_only_the_top_of_a_huge_sheet(monkeypatch):
    monkeypatch.setattr(xlsx_module, "MAX_CELLS", 20)
    monkeypatch.setattr(xlsx_module, "MAX_BOOK_CELLS", 30)
    rows = {f"A{r}": r for r in range(1, 16)} | {f"B{r}": ("=", f"A{r}*2", r * 2) for r in range(1, 16)}
    rest = {f"A{r}": r for r in range(1, 7)} | {f"B{r}": r for r in range(1, 7)}
    sheets = read_xlsx(xlsx(("Daten", rows), ("Klein", {"A1": "x", "A2": 1}), ("Rest", rest)))
    by = {s.name: s.cells for s in sheets}
    assert cut_of(by["Daten"]) == 10 and len(by["Daten"]) == 20 and by["Daten"][(10, 2)].formula == "A10*2"
    assert cut_of(by["Klein"]) is None and len(by["Klein"]) == 2
    # Die Mappe hat 30 Zellen Platz: Von „Rest“ passen noch 8, also die Zeilen 1 bis 4.
    assert cut_of(by["Rest"]) == 4 and len(by["Rest"]) == 8


def test_a_file_with_a_doctype_after_the_first_bytes_is_refused():
    data = xlsx(("Daten", {"A1": 1}))
    import io
    import zipfile

    source = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as archive:
        for info in source.infolist():
            raw = source.read(info)
            if info.filename.endswith("sheet1.xml"):
                raw = b'<?xml version="1.0"?>' + b" " * 5000 + b'<!DOCTYPE x [<!ENTITY a "b">]>' + raw
            archive.writestr(info, raw)
    with pytest.raises(ValueError):
        read_xlsx(out.getvalue())


def test_cut_note_names_a_sheet_that_was_not_read_at_all():
    report = analyse.Report()
    report.cuts = {"Daten": 0}
    assert pruefer.cut_note("T", report) == "Von „Daten“ lese ich nichts mehr, die Mappe ist sehr groß."


def test_inserting_a_fix_keeps_the_cut():
    cells = Cells({(1, 1): Cell("Umsatz"), (2, 1): Cell("10"), (3, 1): Cell("20", "A2*2")}, 3)
    changed = applied(cells, [{"cells": {"A3": "=A2*3"}}])
    assert cut_of(changed) == 3 and changed[(3, 1)].formula == "A2*3"


# --- Unveränderte Blätter und der Hintergrund ---


def test_an_unchanged_sheet_comes_only_with_its_stamp(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    total = sum(r * 10 for r in range(2, 12))
    body = {"book": "B.xlsx", "active": "Auswertung", "sheets": [daten(), auswertung(total + 5)]}
    first = client.post("/api/addin/pruefen", json=body, headers=HEADERS).json()
    assert [h["cell"] for h in first["hints"] if h["key"].startswith("recalc:")] == ["B2"]
    again = {"book": "B.xlsx", "active": "Auswertung", "sheets": [{"name": "Daten", "same": "d1"}, auswertung(total)]}
    second = client.post("/api/addin/pruefen", json=again, headers=HEADERS).json()
    assert second["state"] == "ok" and not [h for h in second["hints"] if h["key"].startswith("recalc:")]
    # Ein Stempel, den Gegenprobe nicht kennt (nach einem Neustart): Das Fenster schickt alles noch einmal.
    stale = {"book": "B.xlsx", "active": "Auswertung", "sheets": [{"name": "Daten", "same": "d0"}, auswertung(total)]}
    assert client.post("/api/addin/pruefen", json=stale, headers=HEADERS).json()["state"] == "again"


def test_the_old_pane_with_whole_grids_gives_the_same_cells():
    new = daten(4)
    new["formulas_at"] = [[2, 3, "=D2*2"]]
    new["kinds_at"] = [[1, 0, "d"]]
    old = {k: v for k, v in new.items() if not k.endswith("_at")}
    old["formulas"] = [[None] * 4 for _ in new["values"]]
    old["formulas"][2][3] = "=D2*2"
    old["kinds"] = [[None] * 4 for _ in new["values"]]
    old["kinds"][1][0] = "d"
    _, _, a = addin.sheets_from({"book": "B", "active": "Daten", "sheets": [new]})
    _, _, b = addin.sheets_from({"book": "B", "active": "Daten", "sheets": [old]})
    assert dict(a[0].cells) == dict(b[0].cells) and a[0].cells[(3, 4)].formula == "D2*2"
    assert a[0].cells[(2, 1)].value == "17.09.1965"


def _wait(checker, seconds=20):
    end = time.monotonic() + seconds
    while time.monotonic() < end and checker._idle is not None and checker._idle.is_alive():
        time.sleep(0.05)


def test_big_other_sheets_are_checked_in_the_background(tmp_path, monkeypatch):
    monkeypatch.setattr(pruefer, "OTHERS_CELLS", 10)
    monkeypatch.setattr(pruefer, "IDLE_WAIT", 0.0)
    checker = pruefer.Checker(tmp_path)
    item = daten()
    item["values"][3][3] = "dreißig"  # Text zwischen den Zahlen
    first = checker.check_payload({"book": "B", "active": "Auswertung", "sheets": [item, auswertung(0)]})
    assert first["others"] == [{"name": "Daten", "fehler": None, "pruefen": None, "pending": True}]
    _wait(checker)
    later = checker.others()
    assert later["sheet"] == "Auswertung" and later["others"][0]["name"] == "Daten"
    assert later["others"][0]["fehler"] is not None and "pending" not in later["others"][0]
    # Das Ergebnis im Hintergrund ist dasselbe wie direkt geprüft.
    direct = checker.check_payload({"book": "B", "active": "Daten",
                                    "sheets": [{"name": "Daten", "same": "d1"}, {"name": "Auswertung", "same": "a1"}]})
    assert pruefer._levels(direct) == {k: later["others"][0][k] for k in ("fehler", "pruefen")}
    assert direct["hints"]


def test_a_new_request_stops_the_background(tmp_path, monkeypatch):
    monkeypatch.setattr(pruefer, "OTHERS_CELLS", 10)
    monkeypatch.setattr(pruefer, "IDLE_WAIT", 0.3)
    checker = pruefer.Checker(tmp_path)
    checker.check_payload({"book": "B", "active": "Auswertung", "sheets": [daten(), auswertung(0)]})
    checker._generation += 1  # wie eine neue Anfrage, bevor der Hintergrund anfängt
    _wait(checker)
    assert checker.others()["others"][0].get("pending") is True


def test_analysis_stops_between_steps_when_asked():
    cells = grid([["Fall", "Wert"], ["A", 1], ["B", 2], ["Summe", ("=SUM(B2:B3)", 3)]])
    with analyse.stoppable(lambda: True), pytest.raises(analyse.Stopped):
        analyse.analyse(cells, layout(Sheet("T", cells)))
    # Danach rechnet sie wieder ganz.
    assert analyse.analyse(cells, layout(Sheet("T", cells))).stats["formeln"] == 1


def test_the_others_endpoint_answers_without_checking(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    assert client.post("/api/addin/andere", json={}, headers=HEADERS).json()["others"] == []
    client.post("/api/addin/pruefen", json={"book": "B", "active": "Auswertung", "sheets": [daten(), auswertung(0)]},
                headers=HEADERS)
    answer = client.post("/api/addin/andere", json={}, headers=HEADERS).json()
    assert answer["book"] == "B" and answer["sheet"] == "Auswertung" and answer["others"][0]["name"] == "Daten"
    assert client.post("/api/addin/andere", json={}).status_code == 403


# --- Fehlalarme, die große Mappen zeigten ---


def test_a_squared_height_in_the_bmi_is_no_typed_number():
    rows = [["Fall", "Größe cm", "Gewicht kg", "Visiten", "BMI"]]
    for r, (h, w, v) in enumerate([(170, 70, 1), (182, 90, 2), (165, 55, 3), (175, 80, 2), (160, 62, 4), (190, 95, 1)], 2):
        rows.append([f"P{r}", h, w, v, (f"=C{r}/(B{r}/100)^2", w / (h / 100) ** 2)])
    assert hints(rows)["hints"] == []


def test_the_same_factor_in_every_row_is_no_number_copied_from_the_row():
    rows = [["Fall", "Anzahl", "Betrag", "Doppelt"]]
    for r, (n, b) in enumerate([(1, 10), (2, 20), (3, 30), (4, 40), (5, 50), (6, 60)], 2):
        rows.append([f"F{r}", n, b, (f"=C{r}*2", b * 2)])
    keys = [h["key"] for h in hints(rows)["hints"]]
    assert not [k for k in keys if k.startswith("copied:")]


def test_digits_of_a_code_are_not_the_fixed_number():
    rows = [["Code", "Betrag", "Mit Zuschlag"]]
    for r, b in enumerate([10, 20, 30, 40, 50], 2):
        rows.append([f"PAT-10{r}", b, (f"=B{r}*1.05", b * 1.05)])
    rows += [["MW-1,05"], ["Zuschlag (1,05)"]]
    text = next(h["text"] for h in hints(rows)["hints"] if h["key"] == "constant:1,05")
    assert "A8 „Zuschlag (1,05)“" in text and "MW-1,05" not in text


def test_a_sum_does_not_miss_what_stands_below_it():
    rows = [["Station", "Fälle", "Tage"]]
    for r, (f, t) in enumerate([(10, 50), (12, 70), (8, 40), (15, 90), (11, 60)], 2):
        rows.append([f"S{r}", f, t])
    rows += [["Summe", ("=SUM(B2:B6)", 56), ("=SUM(C2:C6)", 310)], [], ["Ziel", 60], ["Vorjahr", 52]]
    assert hints(rows)["hints"] == []


def test_a_mean_column_in_the_total_row_is_the_mean_over_all():
    rows = [["Station", "Fälle", "Mittlere Liegedauer"]]
    for r, (f, t) in enumerate([(10, 5.0), (12, 5.5), (8, 4.0), (15, 6.0), (11, 5.5)], 2):
        rows.append([f"S{r}", f, t])
    rows.append(["Summe", ("=SUM(B2:B6)", 56), ("=AVERAGE(C2:C6)", 5.2)])
    assert hints(rows)["hints"] == []
    # Eine Summenspalte rechnet in der Summenzeile weiterhin die Summe.
    rows[-1][1] = ("=AVERAGE(B2:B6)", 11.2)
    assert [h["key"] for h in hints(rows)["hints"] if h["cell"] == "B7"]
