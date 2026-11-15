"""Summen gegen ihre Teile: Summenzeilen, Summenspalten, Gesamtsummen und Teile, die nicht alle Zeilen erfassen."""

import pytest

from voiceagent import formel
from voiceagent.analyse import Sheet as Checked
from voiceagent.analyse import analyse
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Sheet

# Erfundene Zahlen: jede Region dreimal.
SALES = [("Nord", 100.0), ("Ost", 200.0), ("Süd", 300.0), ("West", 400.0), ("Nord", 50.0), ("Ost", 25.0),
         ("Süd", 10.0), ("West", 40.0), ("Nord", 5.0), ("Ost", 2.5), ("Süd", 1.0), ("West", 4.0)]
REGIONS = [("Umsatz Nord", 'SUMIF(A:A,"Nord",B:B)'), ("Umsatz Ost", 'SUMIF(A:A,"Ost",B:B)'),
           ("Umsatz Süd", 'SUMIF(A:A,"Süd",B:B)'), ("Umsatz West", 'SUMIF(A:A,"West",B:B)')]


def filled(cells):
    """Was Excel in den Formelzellen zeigt."""
    for _ in range(4):
        calc = Checked(cells, layout(Sheet("T", cells)), "T", {}).evaluator()[0]
        for pos, cell in list(cells.items()):
            if cell.formula:
                value = calc.run(formel.parse(cell.formula))
                cells[pos] = Cell(value.code if isinstance(value, formel.Err) else repr(value), cell.formula)
    return cells


def overview(lines, total, sales=SALES, label="Umsatz gesamt"):
    """Die Tabelle in A:B, rechts daneben die Übersicht in D:E mit der Summe darunter."""
    cells = {(1, 1): Cell("Region"), (1, 2): Cell("Umsatz")}
    for r, (region, amount) in enumerate(sales, start=2):
        cells[(r, 1)], cells[(r, 2)] = Cell(region), Cell(repr(amount))
    for r, (name, f) in enumerate([*lines, (label, total)], start=2):
        cells[(r, 4)] = Cell(name)
        cells[(r, 5)] = Cell(f) if f.replace(".", "").isdigit() else Cell("0", f)
    return filled(cells)


def grid(rows, total_row=None, total_col=None, corner=None):
    """Eine Übersicht Region mal Produkt mit Zeilensummen rechts und Spaltensummen unten."""
    cells = {(1, 1): Cell("Umsatz"), (1, 2): Cell("Maus"), (1, 3): Cell("Dock"), (1, 4): Cell("Monitor"), (1, 5): Cell("Summe")}
    for r, (region, *values) in enumerate(rows, start=2):
        cells[(r, 1)] = Cell(region)
        for c, v in enumerate(values, start=2):
            cells[(r, c)] = Cell(repr(v))
        cells[(r, 5)] = Cell("0", (total_col or {}).get(r, f"SUM(B{r}:D{r})"))
    last = len(rows) + 2
    cells[(last, 1)] = Cell("Summe")
    for c, letter in enumerate("BCDE", start=2):
        cells[(last, c)] = Cell("0", (total_row or {}).get(letter, f"SUM({letter}2:{letter}{last - 1})"))
    if corner:
        cells[(last, 5)] = Cell("0", corner)
    return filled(cells)


ROWS = [("Nord", 10.0, 20.0, 30.0), ("Ost", 1.0, 2.0, 3.0), ("Süd", 100.0, 200.0, 300.0), ("West", 5.0, 6.0, 7.0)]


def found(cells):
    return {f.key: f for f in analyse(cells, layout(Sheet("T", cells)), name="T").findings}


@pytest.mark.parametrize("total, what, got", [
    ("SUM(E2:E4)", "E5 (Umsatz West, 444) nicht mit", "693,5"),
    ("E2+E3+E4+E5+E3", "E3 (Umsatz Ost, 227,5) doppelt", "1365"),
    ("E2+E3-E4+E5", "E4 (Umsatz Süd, 311) abgezogen statt dazu", "515,5"),
    ("SUM(E2:E5)+B3", "B3 (Umsatz Ost) dazu, das nicht zu den Teilen gehört", "1337,5"),
    ("SUM(E2:E5)+100", "die feste Zahl 100 dazu", "1237,5"),
])
def test_a_total_that_does_not_add_up_its_parts(total, what, got):
    hint = found(overview(REGIONS, total))["parts:E6"]
    assert hint.text == (f"„Umsatz gesamt“ in E6 soll die 4 Teile E2:E5 zusammenzählen, zählt aber {what}. "
                         f"So ergibt E6 {got} statt 1137,5.")
    assert hint.tip == "Korrigiert: =SUMME(E2:E5), nachgerechnet 1137,5."
    assert hint.level == "fehler"


def test_parts_named_like_the_groups_of_the_table():
    """„Nord“, „Ost“, „Süd“ über „Total“: Werte derselben Spalte, also Teile eines Ganzen."""
    lines = [(name.split()[1], f) for name, f in REGIONS]
    hint = found(overview(lines, "SUM(E2:E4)", label="Total"))["parts:E6"]
    assert hint.text.startswith("„Total“ in E6 soll die 4 Teile E2:E5 zusammenzählen, zählt aber E5 (West, 444) nicht mit.")


def test_a_typed_total():
    hint = found(overview(REGIONS, "1100"))["parts:E6"]
    assert hint.text == "„Umsatz gesamt“ in E6 ist eine feste Zahl 1100, die 4 Teile E2:E5 ergeben zusammen 1137,5."
    assert hint.tip == "Korrigiert: =SUMME(E2:E5), nachgerechnet 1137,5."


def test_a_row_total_and_a_column_total_in_an_overview():
    hint = found(grid(ROWS, total_col={3: "SUM(B3:C3)"}))
    assert hint["parts:E3"].text == ("Die Summe der Zeile „Ost“ in E3 soll die 3 Teile B3:D3 zusammenzählen, zählt aber D3 "
                                     "(Monitor, 3) nicht mit. So ergibt E3 3 statt 6.")
    assert hint["parts:E3"].tip == "Korrigiert: =SUMME(B3:D3), nachgerechnet 6."
    hint = found(grid(ROWS, total_row={"C": "SUM(C2:C4)"}))
    assert hint["parts:C6"].text == ("Die Summe der Spalte „Dock“ in C6 soll die 4 Teile C2:C5 zusammenzählen, zählt aber C5 "
                                     "(West, 6) nicht mit. So ergibt C6 222 statt 228.")


def test_the_grand_total_in_the_corner():
    hint = found(grid(ROWS, corner="SUM(E2:E4)"))
    assert hint["parts:E6"].text == ("Die Gesamtsumme in E6 soll die 4 Teile E2:E5 zusammenzählen, zählt aber E5 (West, 18) "
                                     "nicht mit. So ergibt E6 666 statt 684.")
    # Die Zeilensummen oder die Spaltensummen: beides ist richtig.
    assert not [k for k in found(grid(ROWS, corner="SUM(B6:D6)")) if k.startswith("parts")]


def test_a_missing_part_that_is_zero_for_now():
    rows = [*ROWS[:3], ("West", 0.0, 0.0, 0.0)]
    hint = found(grid(rows, total_row={"C": "SUM(C2:C4)"}))["parts:C6"]
    assert hint.level == "pruefen"
    assert hint.text == ("Die Summe der Spalte „Dock“ in C6 soll die 4 Teile C2:C5 zusammenzählen, zählt aber C5 (West, 0) "
                         "nicht mit. Weil dort gerade 0 steht, stimmt C6 noch. Sobald sich der Wert ändert, ist die Summe "
                         "falsch.")
    assert hint.tip == "Korrigiert: =SUMME(C2:C5), nachgerechnet 222."


def test_a_group_without_its_own_line():
    hint = found(overview(REGIONS[:3], "SUM(E2:E4)"))["groups:E5"]
    assert hint.text == ("„Umsatz gesamt“ in E5 zählt die 3 Teile E2:E4 zusammen, aber die Teile erfassen nicht alle Zeilen "
                         "der Tabelle: Region „West“ hat keine eigene Zeile (3 Zeilen, zusammen 444). Darum ergibt E5 693,5 "
                         "statt 1137,5.")
    assert hint.tip == ("Füg unter D4 eine Zeile „Umsatz West“ mit =SUMMEWENN(A:A;\"West\";B:B) ein (nachgerechnet 444) und "
                        "nimm sie in die Summe E5 auf. Oder rechne E5 direkt aus der Tabelle: =SUMME(B:B) (nachgerechnet "
                        "1137,5).")


def test_a_typo_in_the_table():
    sales = [*SALES[:4], ("Nrod", 50.0), *SALES[5:]]
    hint = found(overview(REGIONS, "SUM(E2:E5)", sales=sales))["groups:E6"]
    assert hint.text == ("„Umsatz gesamt“ in E6 zählt die 4 Teile E2:E5 zusammen, aber die Teile erfassen nicht alle Zeilen "
                         "der Tabelle: in A6 steht „Nrod“, wohl vertippt für „Nord“ (50). Darum ergibt E6 1087,5 statt "
                         "1137,5.")
    assert hint.tip == "Schreib in A6 „Nord“ statt „Nrod“. Dann ergibt E6 1137,5 (nachgerechnet)."


def test_an_empty_group_in_the_table():
    sales = [*SALES[:4], ("", 50.0), *SALES[5:]]
    hint = found(overview(REGIONS, "SUM(E2:E5)", sales=sales))["groups:E6"]
    assert hint.text.endswith("in A6 ist „Region“ leer (50), und diese Zeile zählt in keinem Teil. Darum ergibt E6 1087,5 "
                              "statt 1137,5.")
    assert hint.tip == ("Trag in A6 die passende „Region“ ein. Oder rechne E6 direkt aus der Tabelle: =SUMME(B:B) "
                        "(nachgerechnet 1137,5).")


@pytest.mark.parametrize("lines, total, label", [
    (REGIONS, "SUM(E2:E5)", "Umsatz gesamt"),  # richtig
    (REGIONS, "ROUND(SUM(E2:E5),0)", "Umsatz gesamt"),  # gerundet
    (REGIONS, "AVERAGE(E2:E5)", "Durchschnitt"),  # keine Summe
    ([("Einnahmen", "1200"), ("Ausgaben", "180")], "E2-E3", "Summe"),  # Einnahmen minus Ausgaben
    ([("Preis", "12"), ("Menge", "3")], "E2*E3", "Summe"),  # Preis mal Menge
    ([("Anzahl", "COUNT(B:B)"), *REGIONS], "SUM(E3:E6)", "Umsatz gesamt"),  # die Anzahl darüber gehört nicht dazu
])
def test_quiet_when_the_total_is_right_or_something_else(lines, total, label):
    hints = found(overview(lines, total, label=label))
    assert not [k for k in hints if k.split(":")[0] in ("parts", "groups")], hints


def test_an_overview_next_to_the_table_is_not_its_header():
    """Die Übersicht fängt in Zeile 3 an, neben den Daten: Die Kopfzeile der Tabelle bleibt Zeile 1."""
    cells = overview(REGIONS, "SUM(E2:E5)")
    moved = {(r + 1, c) if c >= 4 else (r, c): cell for (r, c), cell in cells.items()}
    moved[(2, 4)], moved[(2, 5)] = Cell("Umsatz"), Cell("Betrag")
    assert layout(Sheet("T", moved)).header_row == 1


def test_a_share_column_in_the_total_row_is_no_pattern_break():
    cells = grid(ROWS)
    cells[(1, 6)] = Cell("Anteil")
    for r in range(2, 7):
        cells[(r, 6)] = Cell("0", f"E{r}/E$6")
    assert not [k for k in found(filled(cells)) if k.startswith("pattern")]


def test_a_long_table_stays_fast():
    """Tausende Zeilen mit Zeilensumme: Jede Zelle sucht ihren Kopf nur einmal, nicht die ganze Spalte hinauf."""
    import time

    from voiceagent import summen

    cells = {(1, 1): Cell("Name"), (1, 2): Cell("Jan"), (1, 3): Cell("Feb"), (1, 4): Cell("Gesamt")}
    for r in range(2, 4002):
        cells[(r, 1)], cells[(r, 2)], cells[(r, 3)] = Cell(f"P{r}"), Cell(str(r % 7)), Cell(str(r % 5))
        cells[(r, 4)] = Cell(str(r % 7 + r % 5), f"SUM(B{r}:C{r})")
    start = time.monotonic()
    summen.checks(Checked(cells, layout(Sheet("T", cells)), "T", {}))
    assert time.monotonic() - start < 3
