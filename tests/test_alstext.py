from test_bereiche import table

from voiceagent.alstext import day, number
from voiceagent.analyse import analyse
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Sheet


def hints(cells):
    return {f.key: f for f in analyse(cells, layout(Sheet("T", cells)), name="T").findings}


def with_figures(cells):
    cells[(2, 6)], cells[(2, 7)] = Cell("Umsatz Februar"), Cell("4597.1", 'SUMIFS(D:D,A:A,">="&DATE(2026,2,1),A:A,"<"&DATE(2026,3,1))')
    cells[(3, 6)], cells[(3, 7)] = Cell("Umsatz Nord"), Cell("21833.52", 'SUMIF(B:B,"Nord",D:D)')
    return cells


def test_reading_numbers_and_dates_written_as_text():
    assert number("4855.45") == 4855.45 and number("4.855,45") == 4855.45 and number("4855,45") == 4855.45
    assert number("1 234") == 1234 and number("1,234.5") == 1234.5 and number("12 Stück") is None
    assert day("23.02.2026") == 46076 and day("2026-02-23") == 46076 and day("31.02.2026") is None


def test_a_date_as_text_and_what_it_changes():
    cells = with_figures(table())
    cells[(5, 1)] = Cell("23.02.2026", None, "23.02.2026")
    found = hints(cells)["astext:A5"]
    assert found.text == ("In Datum steht bei A5 („23.02.2026“) ein Datum als Text. Excel rechnet nicht damit: Bedingungen "
                          "mit Datumsgrenzen (SUMMEWENNS, ZÄHLENWENNS), MONAT() und Sortieren übersehen es. Dadurch ergibt "
                          "G2 (Umsatz Februar) 4597,1 statt 9452,55.")
    assert found.level == "fehler" and "Text in Spalten" in found.tip


def test_a_number_as_text_replaces_the_general_hint():
    cells = with_figures(table())
    cells[(5, 4)] = Cell("4855.45", None, "4855.45")
    found = hints(cells)
    assert "text:Umsatz (€)" not in found
    assert found["astext:D5"].text.endswith("Dadurch ergibt G2 (Umsatz Februar) 4597,1 statt 9452,55, G3 (Umsatz Nord) "
                                            "21833,52 statt 26688,97.")
    assert "In eine Zahl umwandeln" in found["astext:D5"].tip


def test_both_in_one_row_do_not_hide_each_other():
    cells = with_figures(table())
    cells[(5, 1)] = Cell("23.02.2026", None, "23.02.2026")
    cells[(5, 4)] = Cell("4855.45", None, "4855.45")
    found = hints(cells)
    assert "G2 (Umsatz Februar) 4597,1 statt 9452,55." in found["astext:A5"].text
    assert "G3 (Umsatz Nord) 21833,52 statt 26688,97" in found["astext:D5"].text


def test_numbers_meant_as_text_stay_quiet():
    cells = table()
    assert not [k for k in hints(cells) if k.startswith("astext")]
    # Nummern mit Null vorn und Fallnummern stehen mit Absicht als Text da.
    cells[(5, 4)] = Cell("00123", None, "00123")
    assert not [k for k in hints(cells) if k.startswith("astext")]
