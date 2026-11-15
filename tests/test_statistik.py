"""Statistik über Codes, Nullen und leere Zellen, die falsche Streuung und Anteile von der falschen Summe."""

import re
import statistics

import pytest

from voiceagent import formel
from voiceagent.analyse import Sheet as Checked
from voiceagent.analyse import analyse
from voiceagent.workbook import FUNCTIONS, layout
from voiceagent.xlsx import Cell, Sheet

# Erfundenes Alter, 999 heißt „unbekannt“.
AGES = [54, 61, 999, 47, 72, 999, 38, 66, 59, 999, 44, 70]
CLEAN = [a for a in AGES if a != 999]


def filled(cells):
    """Was Excel in den Formelzellen zeigt."""
    for _ in range(4):
        calc = Checked(cells, layout(Sheet("T", cells)), "T", {}).evaluator()[0]
        for pos, cell in list(cells.items()):
            if cell.formula:
                value = calc.run(formel.parse(cell.formula))
                text = value.code if isinstance(value, formel.Err) else repr(value) if isinstance(value, float) else str(value)
                cells[pos] = Cell(text, cell.formula)
    return cells


def table(values, figures, head="Alter"):
    """Fälle in A:B, rechts daneben die Kennzahlen in D:E."""
    cells = {(1, 1): Cell("Fall"), (1, 2): Cell(head)}
    for r, v in enumerate(values, start=2):
        cells[(r, 1)] = Cell(f"F{r:03d}")
        cells[(r, 2)] = Cell(v if isinstance(v, str) else repr(float(v)))
    for r, (name, f) in enumerate(figures, start=2):
        cells[(r, 4)] = Cell(name)
        cells[(r, 5)] = Cell("0", f)
    return filled(cells)


def found(cells):
    return {f.key: f for f in analyse(cells, layout(Sheet("T", cells)), name="T").findings}


def run(cells, german):
    """Eine vorgeschlagene Formel, wie sie in deutsches Excel eingefügt wird, auf dem Blatt ausrechnen."""
    back = {name: english for english, name in FUNCTIONS.items()}
    out = []
    for i, part in enumerate(re.split(r'("(?:[^"]|"")*")', german)):
        if i % 2 == 0:
            part = re.sub(r"(?<![A-Za-z0-9_.])([A-ZÄÖÜ][A-ZÄÖÜ0-9.]*)\(", lambda m: back.get(m.group(1), m.group(1)) + "(", part)
            part = re.sub(r"(\d),(\d)", r"\1.\2", part).replace(";", ",")
        out.append(part)
    calc = Checked(cells, layout(Sheet("T", cells)), "T", {}).evaluator()[0]
    return calc.run(formel.parse("".join(out)))


def suggested(hint):
    return hint.tip.split("Korrigiert: =", 1)[1].split(", nachgerechnet", 1)[0]


@pytest.mark.parametrize("name, formula, fixed, got, want", [
    ("Mittelwert Alter", "AVERAGE(B2:B13)", 'MITTELWERTWENN(B2:B13;"<>999")', "292,3333", "56,7778"),
    ("Median Alter", "MEDIAN(B2:B13)", "MEDIAN(FILTER(B2:B13;B2:B13<>999))", "63,5", "59"),
    ("Maximum", "MAX(B2:B13)", 'MAXWENNS(B2:B13;B2:B13;"<>999")', "999", "72"),
    ("Anzahl", "COUNT(B2:B13)", "ANZAHL(B2:B13)-ZÄHLENWENN(B2:B13;999)", "12", "9"),
])
def test_statistics_over_a_code_for_missing(name, formula, fixed, got, want):
    cells = table(AGES, [(name, formula)])
    hint = found(cells)["statcode:E2"]
    assert hint.level == "fehler"
    assert "darin steht 3-mal 999 (B4, B7 und B11), wohl ein Code für „fehlt“" in hint.text
    assert hint.text.endswith(f"So ergibt E2 {got} statt {want}.")
    assert suggested(hint) == fixed
    # Die Korrektur ergibt in Excel genau den Wert ohne die Codes.
    exact = {"Mittelwert Alter": statistics.mean(CLEAN), "Median Alter": statistics.median(CLEAN),
             "Maximum": max(CLEAN), "Anzahl": len(CLEAN)}[name]
    assert run(cells, fixed) == pytest.approx(exact)


def test_a_range_that_also_stops_too_early():
    """Codes und ein Bereich, der vor dem Ende der Spalte aufhört: Die Korrektur nimmt beides zurecht."""
    hint = found(table(AGES, [("Mittelwert Alter", "AVERAGE(B2:B10)")]))["statcode:E2"]
    assert "Dazu hört der Bereich bei B10 auf, in Alter stehen Zahlen aber bis B13." in hint.text
    assert suggested(hint) == 'MITTELWERTWENN(B2:B13;"<>999")'


def test_statistics_without_codes_stay_quiet():
    cells = table(CLEAN, [("Mittelwert Alter", "AVERAGE(B2:B10)"), ("Median Alter", "MEDIAN(B2:B10)"),
                          ("SD (Stichprobe)", "STDEV.S(B2:B10)")])
    assert not [k for k in found(cells) if k.split(":")[0] in ("statcode", "statblank", "spread", "spreadmix")]


def test_a_code_already_left_out_stays_quiet():
    cells = table(AGES, [("Mittelwert Alter", 'AVERAGEIF(B2:B13,"<>999")')])
    assert "statcode:E2" not in found(cells)


@pytest.mark.parametrize("formula, how", [
    ("SUM(B2:B13)/COUNTA(A2:A13)", "teilt die Summe von B2:B13 durch ANZAHL2(A2:A13) = 12"),
    ("SUM(B2:B13)/12", "teilt die Summe von B2:B13 durch die feste Zahl 12"),
])
def test_a_mean_that_counts_empty_cells_as_zero(formula, how):
    ages = [str(a) if a != 999 else "" for a in AGES]
    hint = found(table(ages, [("Mittelwert Alter", formula)]))["statblank:E2"]
    assert hint.text == (f"E2 (Mittelwert Alter) {how}, aber nur 9 Zellen darin enthalten Zahlen. Damit zählen leere "
                         "Zellen (B4, B7 und B11) als 0 mit. So ergibt E2 42,5833 statt 56,7778.")
    assert hint.level == "pruefen"
    assert suggested(hint) == "MITTELWERT(B2:B13)"


@pytest.mark.parametrize("formula, how", [
    ("SUM(B2:B13)/COUNTA(B2:B13)", "teilt die Summe von B2:B13 durch ANZAHL2(B2:B13) = 12, aber nur 9 Zellen darin "
                                   "enthalten Zahlen. Damit zählt Text"),
    ("AVERAGEA(B2:B13)", "rechnet MITTELWERTA über B2:B13. MITTELWERTA zählt Text"),
])
def test_a_mean_that_counts_text_as_zero(formula, how):
    ages = [str(a) if a != 999 else "k.A." for a in AGES]
    hint = found(table(ages, [("Mittelwert Alter", formula)]))["statblank:E2"]
    assert hint.text.startswith(f"E2 (Mittelwert Alter) {how} wie „k.A.“ (B4, B7 und B11) als 0 mit.")
    assert hint.level == "fehler"
    assert suggested(hint) == "MITTELWERT(B2:B13)"


@pytest.mark.parametrize("name, formula, said, fixed, got, want", [
    ("SD (Stichprobe)", "STDEV.P(B2:B10)", "rechnet STABW.N, das ist die Streuung der Grundgesamtheit", "STABW.S(B2:B10)",
     "11,2228", "11,9035"),
    ("Varianz Alter", "STDEV.S(B2:B10)", "rechnet STABW.S, das ist die Standardabweichung, nicht die Varianz",
     "VAR.S(B2:B10)", "11,9035", "141,6944"),
    ("Standardfehler", "STDEV.S(B2:B10)", "rechnet nur die Standardabweichung", "STABW.S(B2:B10)/WURZEL(ANZAHL(B2:B10))",
     "11,9035", "3,9678"),
    ("Variationskoeffizient", "AVERAGE(B2:B10)/STDEV.S(B2:B10)", "teilt den Mittelwert durch die Standardabweichung",
     "STABW.S(B2:B10)/MITTELWERT(B2:B10)", "4,7698", "0,209652"),
    ("Standardabweichung", "ROUND(VAR.S(B2:B10),2)", "rechnet VAR.S, das ist die Varianz", "RUNDEN(STABW.S(B2:B10);2)",
     "141,69", "11,9"),
    ("Standardabweichung Grundgesamtheit Punkte", "STDEV.S(B2:B10)", "rechnet STABW.S", "STABW.N(B2:B10)",
     "11,9035", "11,2228"),
])
def test_the_wrong_spread(name, formula, said, fixed, got, want):
    cells = table(CLEAN, [(name, formula)])
    hint = found(cells)["spread:E2"]
    assert hint.text.startswith(f"E2 ({name}) {said}")
    assert hint.text.endswith(f"So zeigt E2 {got} statt {want}.")
    assert suggested(hint) == fixed
    assert hint.level == "fehler"


def test_one_spread_unlike_its_neighbours():
    cells = {(1, c): Cell(f"Messung {c}") for c in range(1, 5)}
    for r in range(2, 12):
        for c in range(1, 5):
            cells[(r, c)] = Cell(repr(float((r * 7 + c * 13) % 23 + c)))
    cells[(13, 5)] = Cell("Standardabweichung")
    for c, letter in enumerate("ABCD", start=1):
        cells[(13, c)] = Cell("0", f"STDEV.{'P' if letter == 'C' else 'S'}({letter}2:{letter}11)")
    hint = found(filled(cells))["spreadmix:C13"]
    assert "STABW.N" in hint.text
    assert suggested(hint) == "STABW.S(C2:C11)"


def test_zeros_from_iferror_counted_as_values():
    cells = {(1, 1): Cell("Fall"), (1, 2): Cell("Dosis"), (1, 3): Cell("Gewicht"), (1, 4): Cell("Dosis pro kg")}
    for r, (dose, weight) in enumerate([(100, 50), (120, 60), (90, ""), (80, 40), (110, ""), (130, 65)], start=2):
        cells[(r, 1)], cells[(r, 2)], cells[(r, 3)] = Cell(f"F{r}"), Cell(str(dose)), Cell(str(weight))
        cells[(r, 4)] = Cell("0", f"IFERROR(B{r}/C{r},0)")
    cells[(2, 6)], cells[(2, 7)] = Cell("Mittelwert"), Cell("0", "AVERAGE(D2:D7)")
    hints = found(filled(cells))
    hint = hints["statzero:G2"]
    assert hint.text == ("G2 (Mittelwert) rechnet MITTELWERT über D2:D7. 2 Zellen darin (D4 und D6) zeigen 0, weil die "
                         "Formel dort #DIV/0! ausrechnet und WENNFEHLER das durch 0 ersetzt (D4: =WENNFEHLER(B4/C4;0)). "
                         "MITTELWERT zählt diese Nullen als echte Werte. So ergibt G2 1,3333 statt 2.")
    assert hint.tip == ('Schreib in D2 =WENNFEHLER(B2/C2;"") ("" statt 0) und zieh es bis D7. '
                        "Dann ergibt G2 2 (nachgerechnet).")
    # „Dosis pro kg“ ist kein Gewicht: kein unmögliches Gewicht von 0 kg.
    assert not [k for k in hints if k.startswith("impossible:")]


def shares(formula, other=True):
    cells = {(1, 1): Cell("Region"), (1, 2): Cell("Umsatz"), (1, 3): Cell("Anteil")}
    for r, (group, amount) in enumerate([("Nord", 100), ("Ost", 300), ("Süd", 600)], start=2):
        cells[(r, 1)], cells[(r, 2)] = Cell(group), Cell(str(amount))
        cells[(r, 3)] = Cell("0", formula.format(r=r))
    cells[(5, 1)], cells[(5, 2)] = Cell("Summe"), Cell("0", "SUM(B2:B4)")
    if other:
        cells[(2, 5)], cells[(2, 6)] = Cell("Vorjahr"), Cell("1200")
    return filled(cells)


def test_shares_of_the_wrong_total():
    hint = found(shares("B{r}/$F$2"))["base:C2"]
    assert hint.text == ("Die Anteile in C2:C4 teilen durch F2 (1200), aber ihre Teile B2:B4 ergeben zusammen 1000, und "
                         "diese Summe steht in B5. Darum ergeben die Anteile zusammen 83,3 % statt 100 %.")
    assert hint.tip.startswith("Korrigiert in C2: =B2/$B$5, dann bis C4 herunterziehen.")


def test_total_divided_by_the_part():
    hint = found(shares("B$5/B{r}", other=False))["base:C2"]
    assert hint.text == ("Die Anteile in C2:C4 teilen die Summe B5 durch den Teil, nicht den Teil durch die Summe. "
                         "So zeigt C2 1000 % statt 10 %.")
    assert hint.tip.startswith("Korrigiert in C2: =B2/B$5, dann bis C4 herunterziehen.")


def test_right_shares_stay_quiet():
    assert not [k for k in found(shares("B{r}/$B$5")) if k.startswith("base:")]


def test_a_total_next_to_a_longer_table_sums_only_its_block():
    """Eine kleine Übersicht neben den Fällen: Die Summe darüber nimmt den Vorjahreswert unter der Leerzeile nicht."""
    cells = {(1, 1): Cell("Fall"), (1, 2): Cell("Punkte"), (1, 4): Cell("Station"), (1, 5): Cell("Fälle")}
    for r in range(2, 12):
        cells[(r, 1)], cells[(r, 2)] = Cell(f"F{r}"), Cell(str(r * 3))
    for r, (name, count) in enumerate([("S1", 854), ("S2", 618), ("S3", 516)], start=2):
        cells[(r, 4)], cells[(r, 5)] = Cell(name), Cell(str(count))
    cells[(5, 4)], cells[(5, 5)] = Cell("Summe"), Cell("0", "SUM(E2:E4)")
    cells[(7, 4)], cells[(7, 5)] = Cell("Vorjahr"), Cell("2274")
    assert "label:E5" not in found(filled(cells))


def test_the_evaluator_knows_the_statistics_functions():
    cells = {(1, 1): Cell("Wert")}
    for r, v in enumerate(["3", "k.A.", "5", "", "999", "7"], start=2):
        cells[(r, 1)] = Cell(v)
    calc = Checked(cells, layout(Sheet("T", cells)), "T", {}).evaluator()[0]
    assert calc.run(formel.parse("AVERAGEA(A2:A7)")) == pytest.approx((3 + 0 + 5 + 999 + 7) / 5)
    # FILTER behält „k.A.“ (Text ist ungleich 999), MEDIAN lässt Text weg wie Excel: Median von 3, 5 und 7.
    assert calc.run(formel.parse("MEDIAN(FILTER(A2:A7,(A2:A7<>999)*(A2:A7<>\"\")))")) == 5.0
    assert calc.run(formel.parse("ROWS(A2:A7)")) == 6.0
    assert calc.run(formel.parse("STDEVP(A2:A4)")) == pytest.approx(1.0)
