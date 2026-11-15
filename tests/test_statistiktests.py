"""T.TEST, Konfidenzintervalle und Prozentpunkte: falsch aufgesetzt erkannt, richtig aufgesetzt still, und jede
Korrektur auf dem Blatt nachgerechnet."""

import random

import pytest
from test_einsetzen import applied
from test_statistik import filled

from voiceagent import formel
from voiceagent.analyse import Sheet as Checked
from voiceagent.analyse import analyse
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Sheet


def found(cells):
    return {f.key: f for f in analyse(cells, layout(Sheet("T", cells)), name="T", exact=True).findings}


def run(cells, english):
    return Checked(cells, layout(Sheet("T", cells)), "T", {}).evaluator()[0].run(formel.parse(english))


@pytest.mark.parametrize("english, excel", [
    # Die Beispiele aus der Hilfe von Microsoft zu jeder Funktion (wo die Hilfe rundet, genauer nachgerechnet).
    ("T.INV.2T(0.546449,60)", 0.606533), ("T.DIST.2T(1.959999998,60)", 0.054644930), ("T.DIST(60,1,TRUE)", 0.99469533),
    ("T.DIST(8,3,FALSE)", 0.000736906521), ("T.DIST.RT(1.959999998,60)", 0.027322465), ("TDIST(1.959999998,60,2)", 0.054644930),
    ("T.INV(0.75,2)", 0.8164966), ("CONFIDENCE.NORM(0.05,2.5,50)", 0.692952), ("CONFIDENCE.T(0.05,1,50)", 0.284196855),
    ("NORM.S.INV(0.908789)", 1.3333347), ("NORM.S.DIST(1.333333,TRUE)", 0.908788726),
])
def test_the_functions_compute_what_excel_shows(english, excel):
    assert run({}, english) == pytest.approx(excel, rel=1e-6)


def test_t_test_like_the_example_in_excel():
    cells = {}
    for r, (a, b) in enumerate(zip([3, 4, 5, 8, 9, 1, 2, 4, 5], [6, 19, 3, 2, 14, 4, 5, 17, 1]), start=2):
        cells[(r, 1)], cells[(r, 2)] = Cell(str(a)), Cell(str(b))
    assert run(cells, "T.TEST(A2:A10,B2:B10,2,1)") == pytest.approx(0.196016, rel=1e-5)
    # Welch mit ungerundeten Freiheitsgraden, wie T.TEST in Excel (die Analyse-Funktionen runden sie).
    assert run(cells, "T.TEST(A2:A10,B2:B10,2,3)") == pytest.approx(0.2022939, rel=1e-6)
    assert run(cells, "T.TEST(A2:A10,B2:B10,1,2)") == pytest.approx(0.0959979, rel=1e-5)
    assert isinstance(run(cells, "T.TEST(A2:A10,B2:B9,2,1)"), formel.Err)


def rows(n, seed, pairs):
    rnd = random.Random(seed)
    cells = {(1, 1): Cell("Patient"), (1, 2): Cell(pairs[0]), (1, 3): Cell(pairs[1])}
    for r in range(2, n + 2):
        before = round(rnd.uniform(130, 160), 1)
        cells[(r, 1)], cells[(r, 2)] = Cell(f"P{r:03d}"), Cell(repr(before))
        cells[(r, 3)] = Cell(repr(round(before - rnd.uniform(0, 10), 1)))
    return cells


def test_before_and_after_of_the_same_cases_belong_paired():
    cells = rows(12, 3, ("Blutdruck vorher", "Blutdruck nachher"))
    cells[(2, 5)], cells[(2, 6)] = Cell("p-Wert"), Cell("0", "T.TEST(B2:B13,C2:C13,2,2)")
    cells = filled(cells)
    hint = found(cells)["ttest:F2"]
    assert hint.level == "pruefen" and "„Blutdruck vorher“ und „Blutdruck nachher“" in hint.text
    assert hint.fix["cells"] == {"F2": "=T.TEST(B2:B13,C2:C13,2,1)"} and hint.fix["german"] == "=T.TEST(B2:B13;C2:C13;2;1)"
    assert "ttest:F2" not in found(applied(cells, hint.fix))


@pytest.mark.parametrize("heads", [("Messung 1", "Messung 2"), ("Umsatz 2024", "Umsatz 2025"), ("Auge links", "Auge rechts")])
def test_other_ways_to_say_the_same_case_twice(heads):
    cells = rows(12, 4, heads)
    cells[(2, 5)], cells[(2, 6)] = Cell("p"), Cell("0", "T.TEST(B2:B13,C2:C13,2,3)")
    assert "ttest:F2" in found(filled(cells))


def test_two_groups_side_by_side_are_not_pairs():
    cells = rows(12, 5, ("Gruppe A", "Gruppe B"))
    cells[(2, 5)], cells[(2, 6)] = Cell("p"), Cell("0", "T.TEST(B2:B13,C2:C13,2,3)")
    assert "ttest:F2" not in found(filled(cells))


def groups(seed, sizes=(10, 10), spread=(5, 5)):
    rnd = random.Random(seed)
    cells = {(1, 1): Cell("Gruppe"), (1, 2): Cell("Wert")}
    r = 2
    for name, n, sd in zip("AB", sizes, spread):
        for _ in range(n):
            cells[(r, 1)], cells[(r, 2)] = Cell(name), Cell(repr(round(rnd.gauss(50, sd), 1)))
            r += 1
    return cells


def test_paired_over_two_groups_in_one_column_is_wrong():
    cells = groups(6)
    cells[(2, 4)], cells[(2, 5)] = Cell("p A gegen B"), Cell("0", "T.TEST(B2:B11,B12:B21,2,1)")
    cells = filled(cells)
    hint = found(cells)["ttest:E2"]
    assert hint.level == "fehler" and "derselben Spalte untereinander" in hint.text
    assert hint.fix["cells"] == {"E2": "=T.TEST(B2:B11,B12:B21,2,3)"}
    assert hint.tip.endswith(f"nachgerechnet p = {formel_p(run(cells, 'T.TEST(B2:B11,B12:B21,2,3)'))}.")


def formel_p(value):
    from voiceagent.statistiktests import _p

    return _p(value)


def test_equal_variances_assumed_where_they_are_far_apart():
    cells = groups(9, sizes=(8, 22), spread=(2, 12))
    cells[(2, 4)], cells[(2, 5)] = Cell("p"), Cell("0", "T.TEST(B2:B9,B10:B31,2,2)")
    cells = filled(cells)
    hint = found(cells)["ttest:E2"]
    assert "Welch (Typ 3)" in hint.text and hint.fix["cells"] == {"E2": "=T.TEST(B2:B9,B10:B31,2,3)"}
    assert "p = 0,031 statt 0,169" in hint.text and "signifikant" in hint.text
    # Weit weg von 0,05 und fast derselbe p-Wert (0,84 statt 0,90): dann lohnt kein Hinweis.
    cells = groups(7, sizes=(8, 22), spread=(2, 12))
    cells[(2, 4)], cells[(2, 5)] = Cell("p"), Cell("0", "T.TEST(B2:B9,B10:B31,2,2)")
    assert "ttest:E2" not in found(filled(cells))
    # Gleich große Gruppen mit ähnlicher Streuung: Typ 2 ist dann in Ordnung.
    cells = groups(7)
    cells[(2, 4)], cells[(2, 5)] = Cell("p"), Cell("0", "T.TEST(B2:B11,B12:B21,2,2)")
    assert "ttest:E2" not in found(filled(cells))


@pytest.mark.parametrize("label, level", [("p", "pruefen"), ("p zweiseitig", "fehler"), ("p einseitig", None)])
def test_one_sided_without_saying_so(label, level):
    cells = groups(8)
    cells[(2, 4)], cells[(2, 5)] = Cell(label), Cell("0", "T.TEST(B2:B11,B12:B21,1,3)")
    hint = found(filled(cells)).get("ttest:E2")
    assert (hint.level if hint else None) == level
    if hint:
        assert hint.fix["cells"] == {"E2": "=T.TEST(B2:B11,B12:B21,2,3)"}


def ages(figures, n=20):
    rnd = random.Random(9)
    cells = {(1, 1): Cell("Fall"), (1, 2): Cell("Alter")}
    for r in range(2, n + 2):
        cells[(r, 1)], cells[(r, 2)] = Cell(f"F{r}"), Cell(repr(float(rnd.randint(30, 80))))
    for r, (name, f) in enumerate(figures, start=2):
        cells[(r, 4)], cells[(r, 5)] = Cell(name), Cell("0", f)
    return filled(cells)


BASE = [("Mittelwert", "AVERAGE(B2:B21)"), ("Standardabweichung", "STDEV.S(B2:B21)"),
        ("Standardfehler", "E3/SQRT(COUNT(B2:B21))")]


def test_confidence_with_the_level_and_the_standard_error():
    cells = ages([*BASE, ("KI halbe Breite", "CONFIDENCE.T(0.95,E4,20)")])
    hint = found(cells)["confidence:E5"]
    assert hint.level == "fehler" and "Alpha 0,95" in hint.text and "Standardfehler" in hint.text
    assert hint.fix["cells"] == {"E5": "=CONFIDENCE.T(0.05,E3,20)"}
    assert hint.fix["german"] == "=KONFIDENZ.T(0,05;E3;20)"
    assert not [k for k in found(applied(cells, hint.fix)) if k.startswith("confidence:")]


def test_confidence_with_the_wrong_count_and_the_normal_distribution_for_few_values():
    cells = ages([*BASE, ("KI halbe Breite", "CONFIDENCE.NORM(0.05,E3,COUNTA(A1:A21))")])
    hint = found(cells)["confidence:E5"]
    assert "n = 21" in hint.text and "KONFIDENZ.T" in hint.text
    assert hint.fix["cells"] == {"E5": "=CONFIDENCE.T(0.05,E3,COUNT(B2:B21))"}
    # Mit 40 Werten passt die Normalverteilung.
    cells = ages([("Mittelwert", "AVERAGE(B2:B41)"), ("Standardabweichung", "STDEV.S(B2:B41)"),
                  ("KI halbe Breite", "CONFIDENCE.NORM(0.05,E3,40)")], n=40)
    assert "confidence:E4" not in found(cells)


def test_a_hand_made_interval_without_the_root_of_n():
    cells = ages([*BASE, ("95%-KI untere Grenze", "E2-1.96*E3")])
    hint = found(cells)["interval:E5"]
    assert "der einzelnen Werte" in hint.text and hint.level == "fehler"
    assert hint.fix["cells"] == {"E5": "=E2-T.INV.2T(0.05,COUNT(B2:B21)-1)*(E3/SQRT(COUNT(B2:B21)))"}
    after = applied(cells, hint.fix)
    assert "interval:E5" not in found(after)
    assert run(after, "E5") == pytest.approx(run(cells, "E2-T.INV.2T(0.05,19)*E3/SQRT(20)"), rel=1e-12)


def test_a_hand_made_interval_with_n_and_with_the_wrong_z():
    big = [("Mittelwert", "AVERAGE(B2:B41)"), ("Standardabweichung", "STDEV.S(B2:B41)")]
    cells = ages([*big, ("KI oben", "E2+1.96*E3/40"), ("99 % KI oben", "E2+1.96*E3/SQRT(40)")], n=40)
    hints = found(cells)
    assert "durch n (40)" in hints["interval:E4"].text and hints["interval:E4"].fix["cells"] == {"E4": "=E2+1.96*E3/SQRT(40)"}
    assert "99 %" in hints["interval:E5"].text and hints["interval:E5"].fix["cells"] == {"E5": "=E2+2.576*E3/SQRT(40)"}
    for hint in hints.values():
        assert hint.key not in found(applied(cells, hint.fix))


@pytest.mark.parametrize("label", ["Referenzbereich oben", "Mittelwert plus 2 SD"])
def test_the_range_of_single_values_is_not_an_interval(label):
    assert not [k for k in found(ages([*BASE, (label, "E2+1.96*E3")])) if k.startswith("interval:")]


def test_a_right_interval_is_quiet():
    cells = ages([*BASE, ("95%-KI unten", "E2-T.INV.2T(0.05,19)*E4"), ("KI halbe Breite", "CONFIDENCE.T(0.05,E3,20)")])
    assert not [k for k in found(cells) if k.startswith(("interval:", "confidence:"))]


def rates(head, formula):
    cells = {(1, 1): Cell("Station"), (1, 2): Cell("Quote 2024"), (1, 3): Cell("Quote 2025"), (1, 4): Cell(head)}
    for r, (a, b) in enumerate([(20, 25), (12.5, 10), (30, 33), (8, 12)], start=2):
        cells[(r, 1)], cells[(r, 2)], cells[(r, 3)] = Cell(f"S{r}"), Cell(f"{a} %".replace(".", ",")), Cell(f"{b} %")
        cells[(r, 4)] = Cell("0", formula.format(r=r))
    return filled(cells)


def test_percentage_points_called_percent():
    cells = rates("Veränderung %", "C{r}-B{r}")
    hint = found(cells)["points:D2"]
    assert hint.level == "pruefen" and "5 Prozentpunkte" in hint.text and "25 %" in hint.text
    assert hint.fix["range"] == "D2:D5" and hint.fix["cells"]["D5"] == "=(C5-B5)/B5"
    assert not [k for k in found(applied(cells, hint.fix)) if k.startswith("points:")]


def test_percentage_points_that_compute_the_relative_change():
    hint = found(rates("Veränderung in Prozentpunkten", "(C{r}-B{r})/B{r}"))["points:D2"]
    assert hint.level == "fehler" and hint.fix["cells"]["D2"] == "=C2-B2"


@pytest.mark.parametrize("head, formula", [("Veränderung in Prozentpunkten", "C{r}-B{r}"), ("Veränderung %", "(C{r}-B{r})/B{r}"),
                                           ("Differenz", "C{r}-B{r}")])
def test_percentages_said_right_are_quiet(head, formula):
    assert not [k for k in found(rates(head, formula)) if k.startswith("points:")]


def test_percentage_points_with_no_change_in_the_first_row():
    cells = {(1, 1): Cell("Station"), (1, 2): Cell("Quote 2024"), (1, 3): Cell("Quote 2025"), (1, 4): Cell("Wachstum %")}
    for r, (a, b) in enumerate([(14.5, 14.5), (10, 12), (30, 27)], start=2):
        cells[(r, 1)], cells[(r, 2)], cells[(r, 3)] = Cell(f"S{r}"), Cell(f"{a} %".replace(".", ",")), Cell(f"{b} %")
        cells[(r, 4)] = Cell("0", f"C{r}-B{r}")
    hint = found(filled(cells))["points:D2"]
    assert hint.fix is not None and hint.fix["cells"]["D3"] == "=(C3-B3)/B3"


def test_two_numbered_groups_side_by_side_are_not_pairs():
    for heads in [("Gruppe 1", "Gruppe 2"), ("Station 1", "Station 2"), ("Arm 1", "Arm 2")]:
        cells = rows(12, 5, heads)
        cells[(2, 5)], cells[(2, 6)] = Cell("p"), Cell("0", "T.TEST(B2:B13,C2:C13,2,3)")
        assert "ttest:F2" not in found(filled(cells)), heads


@pytest.mark.parametrize("label", ["p (one-sided)", "p 1-seitig", "p, einseitig getestet"])
def test_one_sided_said_in_other_words(label):
    cells = groups(8)
    cells[(2, 4)], cells[(2, 5)] = Cell(label), Cell("0", "T.TEST(B2:B11,B12:B21,1,3)")
    assert "ttest:E2" not in found(filled(cells))


def test_paired_over_groups_of_different_size_shows_n_a():
    cells = groups(6, sizes=(10, 14))
    cells[(2, 4)], cells[(2, 5)] = Cell("p"), Cell("0", "T.TEST(B2:B11,B12:B25,2,1)")
    cells = filled(cells)
    assert cells[(2, 5)].value == "#N/A"
    hint = found(cells)["ttest:E2"]
    assert hint.level == "fehler" and "verschieden viele Werte (10 und 14)" in hint.text and "statt #NV" in hint.text
    assert hint.fix["cells"] == {"E2": "=T.TEST(B2:B11,B12:B25,2,3)"}


def test_a_tiny_p_value_from_excel_stays_a_number():
    from voiceagent.addin import _cell_text
    from voiceagent.analyse import value_of
    from voiceagent.xlsx import number_text

    assert number_text(1.3467932460218719e-05) == "0,00001346793246" and number_text(12345678901.5) == "12345678901,5"
    assert value_of(Cell(_cell_text(1.3467932460218719e-05, None))) == pytest.approx(1.346793246e-05)


def test_the_settings_of_a_test_are_no_fixed_numbers():
    # Seiten, Typ und Alpha sagen, wie getestet wird; die 20 ist eine Anzahl, die besser ANZAHL(...) wäre.
    assert formel.literals(formel.parse("T.TEST(A1:A5,B1:B5,2,3)+CONFIDENCE.T(0.05,B1,20)+T.INV.2T(0.05,19)")) == [20.0, 19.0]
