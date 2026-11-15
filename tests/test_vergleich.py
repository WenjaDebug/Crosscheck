"""Zwei Versionen vergleichen: was sich geändert hat und warum ein Ergebnis anders ist."""

from test_addin import payload
from test_excel_live import FakeExcel, client_with
from test_workbook import xlsx

from voiceagent.vergleich import Axis, align, compare, translate
from voiceagent.xlsx import Cell

RATES = {"K": 0.1, "A": 0.2, "M": 0.3}


def tarife(rates=RATES):
    cells = {(1, 1): Cell("Code"), (1, 2): Cell("Satz")}
    for r, (code, rate) in enumerate(rates.items(), start=2):
        cells[(r, 1)], cells[(r, 2)] = Cell(code), Cell(f"{rate * 100:g} %")
    return cells


def rechnung(rows, rates=RATES, total=None, sheet="Tarife", formula="B{r}*C{r}"):
    """rows: (Kunde, Code, Betrag). Satz per SVERWEIS, Total, darunter die Summe."""
    cells = {(1, 1): Cell("Kunde"), (1, 2): Cell("Betrag"), (1, 3): Cell("Satz"), (1, 4): Cell("Total"), (1, 5): Cell("Code")}
    s = 0.0
    for r, (who, code, amount) in enumerate(rows, start=2):
        rate = rates[code]
        cells[(r, 1)], cells[(r, 2)], cells[(r, 5)] = Cell(who), Cell(f"{amount:g}"), Cell(code)
        cells[(r, 3)] = Cell(f"{rate * 100:g} %", f"VLOOKUP(E{r},{sheet}!$A$2:$B$4,2,FALSE)")
        value = amount * rate * (1.1 if "1.1" in formula else 1)
        cells[(r, 4)] = Cell(f"{value:g}", formula.format(r=r))
        s += value
    last = len(rows) + 2
    cells[(last, 1)], cells[(last, 4)] = Cell("Summe"), Cell(f"{total if total is not None else s:g}", f"SUM(D2:D{last - 1})")
    return cells


ROWS = [("Huber", "K", 100), ("Meier", "A", 200), ("Keller", "M", 300)]


def items(found, sheet="Rechnung", kind=None):
    s = next(s for s in found["sheets"] if s["name"] == sheet)
    return [i for i in s["items"] if kind is None or i["kind"] == kind]


def test_nothing_changed():
    book = {"Tarife": tarife(), "Rechnung": rechnung(ROWS)}
    found = compare(book, {k: dict(v) for k, v in book.items()})
    assert found["state"] == "ok" and found["counts"] == {"blatt": 0} and "nichts geändert" in found["text"]


def test_an_inserted_row_and_a_changed_amount_explain_the_new_total():
    old = {"Tarife": tarife(), "Rechnung": rechnung(ROWS)}
    rows = [ROWS[0], ("Brunner", "A", 50), ("Meier", "A", 250), ROWS[2]]
    new = {"Tarife": tarife(), "Rechnung": rechnung(rows)}
    found = compare(old, new, "alt.xlsx", "neu.xlsx")
    # Die Formeln unter der neuen Zeile sind verschoben, aber nicht geändert; SUMME(D2:D4) heißt jetzt D2:D5.
    assert items(found, kind="formel") == [] and items(found, kind="ersetzt") == []
    lines = items(found, kind="zeilen")
    assert len(lines) == 1 and "1 Zeile neu: Zeile 3" in lines[0]["text"] and "„Brunner“" in lines[0]["text"]
    change = items(found, kind="eingabe")
    assert len(change) == 1 and change[0]["cell"] == "B4" and "200 → 250" in change[0]["text"]
    result = items(found, kind="ergebnis")
    assert [r["cell"] for r in result] == ["D6"]
    assert "Summe" in result[0]["text"] and "140 → 160 (+14,3 %)" in result[0]["text"]
    assert "Weil B4 (Betrag Meier) von 200 auf 250 geändert ist und neue Zeilen mitzählen." == result[0]["why"]
    assert result[0]["cells"] == ["B4"]
    assert found["text"] == "1 Ergebnis ist anders, 1 Eingabe ist geändert und 1 Zeile ist neu."


def test_a_formula_replaced_by_a_number_and_a_changed_column():
    old = {"Tarife": tarife(), "Rechnung": rechnung(ROWS)}
    new = {"Tarife": tarife(), "Rechnung": rechnung(ROWS, formula="B{r}*C{r}*1.1")}
    new["Rechnung"][(3, 4)] = Cell("44")
    found = compare(old, new)
    replaced = items(found, kind="ersetzt")
    assert len(replaced) == 1 and replaced[0]["cell"] == "D3" and replaced[0]["level"] == "fehler"
    assert "feste Zahl 44" in replaced[0]["text"]
    changed = items(found, kind="formel")
    assert len(changed) == 1 and "Spalte D (Total) in 2 Zeilen (D2 bis D4)" in changed[0]["text"]
    assert "vorher =Betrag * Satz, jetzt =Betrag * Satz * 1,1" in changed[0]["text"]
    result = items(found, kind="ergebnis")[0]
    assert result["cell"] == "D5" and "die Formeln in D2, D3, D4 anders sind" in result["why"]


def test_sorted_rows_and_a_renamed_sheet_are_no_changes():
    old = {"Tarife": tarife(), "Rechnung": rechnung(ROWS)}
    new = {"Sätze": tarife(), "Rechnung": rechnung([ROWS[2], ROWS[0], ROWS[1]], sheet="Sätze")}
    found = compare(old, new)
    assert items(found, kind="formel") == [] and items(found, kind="eingabe") == [] and items(found, kind="ergebnis") == []
    assert "umsortiert" in items(found, kind="zeilen")[0]["text"]
    renamed = next(s for s in found["sheets"] if s["name"] == "Sätze")
    assert renamed["old"] == "Tarife" and renamed["state"] == "gleich" and found["counts"]["blatt"] == 1


def test_a_changed_rate_on_another_sheet_and_a_deleted_row():
    old = {"Tarife": tarife(), "Rechnung": rechnung(ROWS)}
    rates = {**RATES, "M": 0.4}
    new = {"Tarife": tarife(rates), "Rechnung": rechnung(ROWS[:1] + ROWS[2:], rates)}
    found = compare(old, new)
    assert "B4" in items(found, "Tarife", "eingabe")[0]["text"]
    lines = items(found, kind="zeilen")
    assert len(lines) == 1 and "gibt es nicht mehr: „Meier“" in lines[0]["text"]
    total = [r for r in items(found, kind="ergebnis") if r["cell"] == "D4"][0]
    assert "Zeilen aus dem Bereich gelöscht sind" in total["why"] and "auf Blatt Tarife B4 geändert ist" in total["why"]


def test_translate_moves_references_like_excel():
    rows = align([("a",), ("b",), ("c",), ("d",)], [("a",), ("x",), ("b",), ("c",), ("d",)])
    cols = Axis(pairs={c: c for c in range(1, 6)}).done()
    axes = {"s": (rows, cols, "S")}
    assert rows.added == [2] and rows.new(2) == 3
    assert translate("SUM(D2:D4)+$B$4+'Daten 2024'!B2", "S", axes) == "SUM(D3:D5)+$B$5+'Daten 2024'!B2"
    gone = align([("a",), ("b",), ("c",), ("d",)], [("a",), ("c",), ("d",)])
    assert gone.removed == [2] and gone.new(2) == 2 and gone.new(2, end=True) == 1 and gone.new(9) == 8


def test_compare_in_the_addin(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    headers = {"X-VoiceAgent": "1"}
    old = xlsx(("Wahl", {
        "A2": "Partei", "B2": "Anteil", "C2": "Sitze", "D2": "Stand",
        "A3": "RN", "B3": 0.3, "C3": ("=", "B3*577", 173.1), "D3": 45566,
        "A4": "NFP", "B4": 0.28, "C4": ("=", "B4*577", 161.56), "D4": 45566,
        "A5": "Ensemble", "B5": 0.208, "C5": ("=", "B5*577", 120.016), "D5": 45566,
        "A6": "Rest", "B6": 0.18, "C6": ("=", "B6*577", 103.86),
        "A7": "Summe", "B7": ("=", "SUM(B3:B6)", 0.968), "C7": ("=", "SUM(C3:C5)", 454.676),
        "C8": ("=", "C3/0", "#DIV/0!"),
    }))
    first = client.post("/api/addin/vergleich?name=Wahl%20alt.xlsx", content=old, headers=headers).json()
    assert first["state"] == "again"
    client.post("/api/addin/pruefen", json=payload(), headers=headers)
    found = client.post("/api/addin/vergleich?name=Wahl%20alt.xlsx", content=old, headers=headers).json()
    assert found["state"] == "ok" and found["before"] == "Wahl alt.xlsx" and found["now"] == "Wahl.xlsx"
    wahl = next(s for s in found["sheets"] if s["name"] == "Wahl")
    kinds = [i["kind"] for i in wahl["items"]]
    assert kinds.count("eingabe") == 1 and "formel" not in kinds and "ersetzt" not in kinds
    results = {i["cell"]: i for i in wahl["items"] if i["kind"] == "ergebnis"}
    assert set(results) == {"B7", "C7"} and "B3 (Anteil RN) von 0,3 auf 33,2 % geändert ist" in results["C7"]["why"]
    assert next(s for s in found["sheets"] if s["name"] == "Leer")["state"] == "neu"
    bad = client.post("/api/addin/vergleich?name=x.xlsx", content=b"kein zip", headers=headers).json()
    assert bad["state"] == "error" and "Excel" in bad["text"]
    assert client.post("/api/addin/vergleich", content=old).status_code == 403


def test_a_new_value_without_a_change_before_it_points_at_the_recalculation():
    old = {"Tarife": tarife(), "Rechnung": rechnung(ROWS)}
    new = {"Tarife": tarife(), "Rechnung": rechnung(ROWS)}
    look = new["Rechnung"][(4, 3)].formula
    new["Rechnung"][(4, 3)], new["Rechnung"][(4, 4)] = Cell("20 %", look), Cell("60", "B4*C4")
    new["Rechnung"][(5, 4)] = Cell("110", "SUM(D2:D4)")
    found = compare(old, new)
    total = items(found, kind="ergebnis")
    assert [r["cell"] for r in total] == ["D5"]
    assert total[0]["why"] == "Weil C4 (Satz Keller) in Excel 20 % zeigt, nachgerechnet aber 30 % ergibt."
    assert "2 weitere Formeln" in items(found, kind="folge")[0]["text"]
