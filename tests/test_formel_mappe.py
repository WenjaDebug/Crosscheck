"""Der eigene Formelrechner für ganze Mappen: andere Blätter, Nachschlagen, Text und Datum."""

import pytest

from test_workbook import xlsx

from voiceagent import addin, formel
from voiceagent.formel import Err, Evaluator, foreign, parse, references, serial_of

TABLE = {  # Daten!A1:C5: Kürzel, Name, Satz
    (1, 1): "Kürzel", (1, 2): "Name", (1, 3): "Satz",
    (2, 1): "A", (2, 2): "Allgemein", (2, 3): 0.1,
    (3, 1): "B", (3, 2): "Bau", (3, 3): 0.2,
    (4, 1): "C", (4, 2): "Chemie", (4, 3): 0.3,
    (5, 1): "D", (5, 2): "Dienste", (5, 3): None,
}
STEPS = {(1, 5): 0.0, (2, 5): 1000.0, (3, 5): 5000.0, (1, 6): 0.0, (2, 6): 0.05, (3, 6): 0.08}  # Stufen in E:F
UNSORTED = {(1, 8): 30.0, (2, 8): 10.0, (3, 8): 20.0, (1, 9): "x", (2, 9): "y", (3, 9): "z"}


def calc(formula, here=None, sheets=None):
    here = here or {}
    sheets = {k.casefold(): v for k, v in (sheets or {"Daten": {**TABLE, **STEPS, **UNSORTED}}).items()}

    def other(name, pos):
        if name.casefold() not in sheets:
            raise formel.Unknown(name)
        return sheets[name.casefold()].get(pos)

    def extent(name):
        cells = here if name is None else sheets[name.casefold()]
        return max((r for r, _ in cells), default=1), max((c for _, c in cells), default=1)

    return Evaluator(lambda p: here.get(p), other=other, extent=extent).run(parse(formula))


def test_other_sheets_and_whole_columns():
    assert calc("Daten!C2*100") == pytest.approx(10)
    assert calc("SUM(Daten!C2:C5)") == pytest.approx(0.6)
    assert calc("SUM('Daten'!C:C)") == pytest.approx(0.6)
    assert calc("SUM(B:B)", {(1, 2): 1.0, (9, 2): 2.0}) == 3
    assert calc("'Q1 2024'!A1+1", sheets={"Q1 2024": {(1, 1): 41.0}}) == 42
    assert calc("Übersicht!A1", sheets={"Übersicht": {(1, 1): "ok"}}) == "ok"
    assert calc("'It''s'!A1", sheets={"It's": {(1, 1): 7.0}}) == 7
    node = parse("Daten!B3+C3+SUM(Daten!$A$1:$C$5)")
    assert foreign(node) == {"daten"} and foreign(node, own="DATEN") == set()
    assert references(node, own="Daten")[0] == {(3, 2), (3, 3)} and references(node)[0] == {(3, 3)}


def test_lookups_like_excel():
    assert calc('VLOOKUP("b",Daten!A1:C5,2,FALSE)') == "Bau"
    assert calc('VLOOKUP("C",Daten!$A$2:$C$5,3,0)') == pytest.approx(0.3)
    assert calc('VLOOKUP("X",Daten!A2:C5,2,FALSE)') == Err("#N/A")
    assert calc('IFERROR(VLOOKUP("X",Daten!A2:C5,2,FALSE),"fehlt")') == "fehlt"
    assert calc('VLOOKUP("B",Daten!A2:C5,4,FALSE)') == Err("#REF!")
    assert calc('VLOOKUP("Ch*",Daten!B2:C5,2,FALSE)') == pytest.approx(0.3)
    assert calc('VLOOKUP("D",Daten!A2:C5,3,FALSE)') is None  # leer: Excel zeigt 0
    # Ungefähr in sortierten Stufen: der größte Wert, der nicht größer ist.
    assert calc("VLOOKUP(2500,Daten!E1:F3,2)") == pytest.approx(0.05)
    assert calc("VLOOKUP(9999,Daten!E1:F3,2,TRUE)") == pytest.approx(0.08)
    assert calc("VLOOKUP(-1,Daten!E1:F3,2)") == Err("#N/A")
    assert calc('HLOOKUP("Name",Daten!A1:C3,3,FALSE)') == "Bau"
    assert calc('MATCH("C",Daten!A1:A5,0)') == 4 and calc("MATCH(1000,Daten!E1:E3)") == 2
    assert calc('INDEX(Daten!B1:B5,MATCH("C",Daten!A1:A5,0))') == "Chemie"
    assert calc("INDEX(Daten!A1:C5,3,2)") == "Bau" and calc("INDEX(Daten!A1:C1,2)") == "Name"
    assert calc("INDEX(Daten!A1:C5,9,1)") == Err("#REF!")
    assert calc('XLOOKUP("B",Daten!A2:A5,Daten!B2:B5)') == "Bau"
    assert calc('XLOOKUP("Q",Daten!A2:A5,Daten!B2:B5,"keiner")') == "keiner"
    assert calc("XLOOKUP(2500,Daten!E1:E3,Daten!F1:F3,,-1)") == pytest.approx(0.05)
    assert calc("XLOOKUP(2500,Daten!E1:E3,Daten!F1:F3,,1)") == pytest.approx(0.08)
    assert calc('XMATCH("z",Daten!I1:I3)') == 3 and calc("XMATCH(20,Daten!H1:H3,0,-1)") == 3


def test_approximate_lookup_in_unsorted_values_is_not_guessed():
    with pytest.raises(formel.Unknown):
        Evaluator(lambda p: UNSORTED.get(p)).eval(parse("VLOOKUP(15,H1:I3,2)"))
    with pytest.raises(formel.Unknown):
        Evaluator(lambda p: UNSORTED.get(p)).eval(parse("MATCH(15,H1:H3,1)"))


def test_text_date_and_statistics():
    v = {(1, 1): "  Müller   AG ", (2, 1): "01.03.2024", (3, 1): "15.04.2024", (4, 1): "31.01.2024"}
    assert calc("TRIM(A1)", v) == "Müller AG" and calc("LEN(TRIM(A1))", v) == 9
    assert calc('LEFT("Kardiologie",5)&"-"&UPPER(RIGHT("abc",2))') == "Kardi-BC"
    assert calc('MID("Station 12",9,2)') == "12" and calc('VALUE("3,5")') == 3.5
    assert calc('SUBSTITUTE("a-b-c","-","")') == "abc" and calc('FIND("b","abc")') == 2
    assert calc('TEXTJOIN(", ",TRUE,"a","","b")') == "a, b" and calc('CONCAT("x",1)') == "x1"
    assert calc("A3-A2", v) == 45 and calc("YEAR(A2)", v) == 2024 and calc("MONTH(A3)", v) == 4
    assert calc('DATEDIF(A2,A3,"d")', v) == 45 and calc('DATEDIF(DATE(1980,5,20),A3,"y")', v) == 43
    assert calc("EOMONTH(A4,1)", v) == serial_of("29.02.2024") and calc("EDATE(A4,1)", v) == serial_of("29.02.2024")
    assert calc("WEEKDAY(A2,2)", v) == 5  # Freitag
    assert calc("A3>A2", v) is True
    assert serial_of("01.10.2024 06:00") == pytest.approx(45566.25) and serial_of("12:00") == 0.5 and serial_of("31.02.2024") is None
    nums = {(r, 1): float(x) for r, x in enumerate([3, 1, 4, 1, 5, 9, 2, 6], start=1)}
    assert calc("LARGE(A1:A8,2)", nums) == 6 and calc("SMALL(A1:A8,1)", nums) == 1
    assert calc("PERCENTILE.INC(A1:A8,0.5)", nums) == 3.5 and calc("QUARTILE(A1:A8,1)", nums) == pytest.approx(1.75)
    assert calc("RANK.EQ(5,A1:A8)", nums) == 3 and calc("RANK(5,A1:A8,1)", nums) == 6
    assert calc('IFS(A1>5,"hoch",A1>2,"mittel",TRUE,"tief")', nums) == "mittel"
    assert calc('SWITCH(A2,1,"eins",2,"zwei","andere")', nums) == "eins" and calc("CHOOSE(2,10,20,30)") == 20
    assert calc("ISNUMBER(A1)", nums) is True and calc("ISNA(MATCH(99,A1:A8,0))", nums) is True
    assert calc("ISBLANK(B1)", nums) is True and calc("N(TRUE)+N(\"x\")") == 1
    assert calc("MROUND(17,5)") == 15 and calc("TRUNC(-2.789,1)") == pytest.approx(-2.7) and calc("SIGN(-3)") == -1


# --- Prüfen über die ganze Mappe ---

from voiceagent.analyse import analyse, value_of  # noqa: E402
from voiceagent.workbook import layout  # noqa: E402
from voiceagent.xlsx import Cell, Sheet, kept_text, read_book  # noqa: E402


def text(value: str) -> Cell:
    """Eine Textzelle, wie die Datei oder das Add-in sie liefern (mit dem genauen Text, wo es zählt)."""
    return Cell(" ".join(value.split()), None, kept_text(value))


def tarife(codes=("K", "A", "M", "B", "B")):
    cells = {(1, 1): Cell("Code"), (1, 2): Cell("Satz")}
    for r, (code, rate) in enumerate(zip(codes, ["0,1", "0,2", "0,3", "0,15", "0,25"]), start=2):
        cells[(r, 1)], cells[(r, 2)] = (text(code) if isinstance(code, str) else Cell(str(code))), Cell(rate)
    return cells


def rechnung(column: str, rows=("K", "A", "M", "B", "X", "Y")):
    cells = {(1, 1): Cell("Code"), (1, 2): Cell("Betrag"), (1, 3): Cell("Satz")}
    for r, code in enumerate(rows, start=2):
        cells[(r, 1)], cells[(r, 2)] = text(code), Cell("100")
        cells[(r, 3)] = Cell("0", column.format(r=r, e=r + 4))
    return cells


def report_for(cells, book=None):
    return analyse(cells, layout(Sheet("Rechnung", cells)), name="Rechnung", book=book or {"Tarife": tarife()})


def keys(report):
    return {f.key.split(":")[0]: f for f in report.findings}


def test_approximate_lookup_on_unsorted_table_is_an_error():
    found = keys(report_for(rechnung("VLOOKUP(A{r},Tarife!$A$2:$B$6,2)")))
    assert "approx" not in found and found["lookexact"].level == "fehler"
    assert found["lookexact"].text.startswith("SVERWEIS in C2 sucht ungefähr, weil das vierte Argument FALSCH fehlt, und "
                                              "Tarife!A2:B6 ist nicht aufsteigend sortiert.")
    assert found["lookexact"].tip == ("Korrigiert: =SVERWEIS(A2;Tarife!$A$2:$B$6;2;FALSCH), nachgerechnet 0,1. "
                                      "Danach wieder bis unten herunterziehen.")


def test_iferror_hides_missing_keys_and_says_why():
    found = keys(report_for(rechnung("IFERROR(VLOOKUP(A{r},Tarife!$A$2:$B$6,2,FALSE),0)")))
    hint = found["hidden"]
    assert "In 2 Formeln" in hint.text and "„X“" in hint.text and "WENNFEHLER setzt still" in hint.text
    assert "approx" not in found


@pytest.mark.parametrize("rows, codes, reason, fix", [
    (("K", "A "), ("K", "A", "M"), "weil in A3 ein Leerzeichen am Ende steht. In der Tabelle steht „A“ (Tarife!A3).",
     "=WENNFEHLER(SVERWEIS(GLÄTTEN(A3);Tarife!$A$2:$B$4;2;FALSCH);0), nachgerechnet 0,2"),
    (("K", "123"), ("K", 123, "M"), "weil in A3 die Zahl als Text steht, in der Tabelle aber als Zahl (Tarife!A3).",
     "=WENNFEHLER(SVERWEIS(WERT(A3);Tarife!$A$2:$B$4;2;FALSCH);0), nachgerechnet 0,2"),
])
def test_missing_key_says_why(rows, codes, reason, fix):
    cells = rechnung("IFERROR(VLOOKUP(A{r},Tarife!$A$2:$B$4,2,FALSE),0)", rows=rows)
    found = keys(report_for(cells, {"Tarife": tarife(codes)}))
    assert "hidden" not in found and reason in found["looknear"].text and "Darum zeigt C3 0 statt 0,2." in found["looknear"].text
    assert fix in found["looknear"].tip


def test_number_as_text_in_the_table():
    cells = rechnung("IFERROR(VLOOKUP(A{r},Tarife!$A$2:$B$4,2,FALSE),0)", rows=("K",))
    cells[(3, 1)], cells[(3, 2)], cells[(3, 3)] = Cell("7"), Cell("100"), Cell("0", "IFERROR(VLOOKUP(A3,Tarife!$A$2:$B$4,2,FALSE),0)")
    found = keys(report_for(cells, {"Tarife": tarife(("K", "7", "M"))}))
    assert "weil die Tabelle die Zahlen als Text enthält (Tarife!A3), A3 aber als Zahl" in found["looknear"].text
    assert '=WENNFEHLER(SVERWEIS(A3&"";Tarife!$A$2:$B$4;2;FALSCH);0), nachgerechnet 0,2' in found["looknear"].tip
    # Steht nur ein Schlüssel der Tabelle als Text, hilft &"" den übrigen nicht: dann die Zelle selbst korrigieren.
    cells[(2, 1)] = Cell("9")
    cells[(2, 3)] = Cell("0", "IFERROR(VLOOKUP(A2,Tarife!$A$2:$B$4,2,FALSE),0)")
    found = keys(report_for(cells, {"Tarife": tarife((9, "7", "M"))}))
    assert found["looknear"].tip.startswith("Trag in Tarife!A3 die Zahl ohne Apostroph ein") and "dann zeigt C3 0,2" in found["looknear"].tip


def test_exact_text_survives_the_addin_and_the_file():
    _, _, sheets = addin.sheets_from({"book": "B", "active": "S", "sheets": [
        {"name": "S", "row": 1, "col": 1, "values": [["A ", "123", 123, "Bern", "1. Wahl"]]}]})
    cells = sheets[0].cells
    assert [cells[(1, c)].text for c in range(1, 6)] == ["A ", "123", None, None, None]
    assert cells[(1, 1)].value == "A" and value_of(cells[(1, 1)]) == "A " and value_of(cells[(1, 3)]) == 123.0
    book = read_book(xlsx(("Tabelle1", {"A1": " A", "B1": "007", "C1": 7, "D1": "Bern"})))
    cells = book.sheets[0].cells
    assert [cells[(1, c)].text for c in range(1, 5)] == [" A", "007", None, None]
    assert value_of(cells[(1, 2)]) == "007" and value_of(cells[(1, 3)]) == 7.0


def test_duplicate_key_and_sliding_table():
    found = keys(report_for(rechnung("VLOOKUP(A{r},Tarife!A{r}:B{e},2,FALSE)", rows=("K", "A", "M", "B"))))
    assert "sliding" in found and "C2 sucht in A2:B6, C5 schon in A5:B9" in found["sliding"].text
    found = keys(report_for(rechnung("VLOOKUP(A{r},Tarife!$A$2:$B$6,2,FALSE)", rows=("K", "B"))))
    assert "twice" in found and "„B“ steht 2-mal in Tarife!A2:B6 (Zeilen 5, 6)" in found["twice"].text
    assert "sliding" not in found


def test_nearest_text_key_without_false():
    stufen = {(1, 1): Cell("A"), (2, 1): Cell("C"), (3, 1): Cell("E"), (1, 2): Cell("1"), (2, 2): Cell("3"), (3, 2): Cell("5")}
    cells = rechnung("VLOOKUP(A{r},Stufen!$A$1:$B$3,2)", rows=("C", "D"))
    found = keys(report_for(cells, {"Stufen": stufen}))
    assert "nearest" not in found
    assert found["lookexact"].text == ("SVERWEIS in C3 findet „D“ nicht in Stufen!A1:B3 und nimmt still die Zeile von „C“ "
                                       "(Stufen!A2), weil das vierte Argument FALSCH fehlt. So zeigt C3 3. Mit genauer Suche "
                                       "zeigt Excel #NV, solange „D“ in der Tabelle fehlt.")
    assert found["lookexact"].tip.startswith("Korrigiert: =SVERWEIS(A3;Stufen!$A$1:$B$3;2;FALSCH).")


def test_recalculates_across_sheets_and_unknown_without_them():
    cells = rechnung("B{r}*Tarife!B{r}", rows=("K", "A", "M"))
    cells[(2, 3)], cells[(3, 3)], cells[(4, 3)] = Cell("10", "B2*Tarife!B2"), Cell("20", "B3*Tarife!B3"), Cell("99", "B4*Tarife!B4")
    report = report_for(cells)
    wrong = [f for f in report.findings if f.key.startswith("recalc:")]
    assert len(wrong) == 1 and "C4" in wrong[0].text and "30" in wrong[0].text
    assert report.stats["nicht_nachrechenbar"] == 0
    alone = analyse(cells, layout(Sheet("Rechnung", cells)))
    assert alone.stats["nicht_nachrechenbar"] == 3 and not [f for f in alone.findings if f.key.startswith("recalc:")]


def test_dates_are_recalculated_as_dates():
    cells = {(1, 1): Cell("Eintritt"), (1, 2): Cell("Austritt"), (1, 3): Cell("Tage"), (1, 4): Cell("Kontrolle"),
             (2, 1): Cell("01.03.2024"), (2, 2): Cell("15.04.2024"), (2, 3): Cell("45", "B2-A2"), (2, 4): Cell("15.05.2024", "B2+30"),
             (3, 1): Cell("01.03.2024"), (3, 2): Cell("11.03.2024"), (3, 3): Cell("10", "B3-A3"), (3, 4): Cell("12.04.2024", "B3+30")}
    report = analyse(cells, layout(Sheet("T", cells)))
    wrong = [f.text for f in report.findings if f.key.startswith("recalc:")]
    assert wrong == ["D3 (Kontrolle) zeigt 12.04.2024, nachgerechnet ergibt =Austritt + 30 aber 10.04.2024."] or (
        len(wrong) == 1 and "12.04.2024" in wrong[0] and "10.04.2024" in wrong[0])


def test_many_lookups_stay_fast():
    import time

    table = {(1, 1): Cell("ID"), (1, 2): Cell("Wert")}
    for r in range(2, 3002):
        table[(r, 1)], table[(r, 2)] = Cell(f"P{r:05d}"), Cell(str(r))
    cells = {(1, 1): Cell("ID"), (1, 2): Cell("Wert")}
    for r in range(2, 2002):
        cells[(r, 1)] = Cell(f"P{r:05d}")
        cells[(r, 2)] = Cell(str(r), f"VLOOKUP(A{r},Liste!$A$2:$B$3001,2,FALSE)")
    start = time.monotonic()
    report = analyse(cells, layout(Sheet("Rechnung", cells)), name="Rechnung", book={"Liste": table})
    assert time.monotonic() - start < 4
    assert report.stats["formeln"] == 2000 and not [f for f in report.findings if f.level == "fehler"]


def test_many_wrong_values_in_one_column_are_one_hint():
    cells = rechnung("B{r}*Tarife!B{r}", rows=("K", "A", "M", "B"))
    for r in range(2, 6):
        cells[(r, 3)] = Cell("0", f"B{r}*Tarife!B{r}")
    wrong = [f for f in report_for(cells).findings if f.key.startswith("recalc:")]
    assert len(wrong) == 1 and "In 4 Zellen von Satz (C2 bis C5) zeigt Excel etwas anderes" in wrong[0].text


def test_repeated_codes_are_groups_not_duplicates():
    cells = rechnung("B{r}*2", rows=("K", "A", "K", "A", "K", "A", "K", "A"))
    for r in range(2, 10):
        cells[(r, 3)] = Cell("200", f"B{r}*2")
    assert "dupe" not in keys(report_for(cells))
    cells = rechnung("B{r}*2", rows=("K", "A", "M", "B", "X", "K"))
    for r in range(2, 8):
        cells[(r, 3)] = Cell("200", f"B{r}*2")
    assert "dupe" in keys(report_for(cells))
