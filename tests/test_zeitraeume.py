"""Alter und Zeiträume: von Hand gerechnet erkannt, genau gerechnet still, und jede Korrektur für jede Zeile
nachgerechnet."""

from datetime import date, timedelta

import pytest

from voiceagent import formel
from voiceagent.analyse import Sheet as Checked
from voiceagent.analyse import analyse
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Sheet
from voiceagent.zeitraeume import _freeze

KEYS = ("age:", "months:", "today:", "days:")


def filled(cells):
    """Was Excel zeigt, mit HEUTE() als heutigem Datum."""
    for _ in range(3):
        calc = Checked(cells, layout(Sheet("T", cells)), "T", {}).evaluator()[0]
        for pos, cell in list(cells.items()):
            if cell.formula:
                value = calc.run(_freeze(formel.parse(cell.formula)))
                text = value.code if isinstance(value, formel.Err) else repr(value) if isinstance(value, float) else str(value)
                cells[pos] = Cell(text, cell.formula)
    return cells


def found(cells):
    return {f.key: f for f in analyse(cells, layout(Sheet("T", cells)), name="T", exact=True).findings if f.key.startswith(KEYS)}


def applied(cells, fix):
    cells = dict(cells)
    for where, formula in fix["cells"].items():
        cells[formel.position(where)] = Cell("0", formula[1:])
    return filled(cells)


def shown(cells, where):
    return float(cells[formel.position(where)].value)


def d(day: date) -> str:
    return day.strftime("%d.%m.%Y")


# Geburtsdatum, Eintritt, Austritt. Fall 3 hat zwei Tage nach dem Eintritt Geburtstag, Fall 5 am 29. Februar.
CASES = [(date(1961, 5, 3), date(2025, 4, 28), date(2025, 5, 6)), (date(1980, 1, 15), date(2025, 6, 2), date(2025, 6, 2)),
         (date(1985, 3, 12), date(2025, 3, 10), date(2025, 3, 21)), (date(1999, 11, 30), date(2025, 9, 1), date(2025, 9, 15)),
         (date(2000, 2, 29), date(2026, 2, 28), date(2026, 3, 4)), (date(1944, 7, 7), date(2025, 7, 7), date(2025, 7, 30))]
AGES = [63, 45, 39, 25, 25, 81]


def cases(head, formula, rows=CASES, start=True, end=True):
    cells = {(1, 1): Cell("Fall"), (1, 2): Cell("Geburtsdatum")}
    heads = [h for h, keep in (("Eintritt", start), ("Austritt", end)) if keep]
    for c, h in enumerate(heads, start=3):
        cells[(1, c)] = Cell(h)
    col = 3 + len(heads)
    cells[(1, col)] = Cell(head)
    for r, (b, e, a) in enumerate(rows, start=2):
        cells[(r, 1)], cells[(r, 2)] = Cell(f"F{r}"), Cell(d(b))
        for c, day in zip(range(3, col), [x for x, keep in ((e, start), (a, end)) if keep]):
            cells[(r, c)] = Cell(d(day))
        cells[(r, col)] = Cell("0", formula.format(r=r))
    return filled(cells)


def test_age_over_365_days_is_a_year_too_old_before_the_birthday():
    cells = cases("Alter", "INT((C{r}-B{r})/365)")
    hint = found(cells)["age:E2"]
    assert hint.level == "fehler" and "durch 365" in hint.text
    assert "In 3 von 6 Zeilen" in hint.text and "Zeilen 2, 4, 6; in Zeile 2: 64 statt 63 Jahre, der 64. Geburtstag ist erst am " \
        "03.05.2025" in hint.text
    assert hint.fix["range"] == "E2:E7" and hint.fix["cells"]["E7"] == '=DATEDIF(B7,C7,"Y")'
    assert hint.tip.startswith('Korrigiert in E2: =DATEDIF(B2;C2;"Y"), bis E7 herunterziehen.')
    after = applied(cells, hint.fix)
    assert [shown(after, f"E{r}") for r in range(2, 8)] == AGES and not found(after)


def test_the_29th_of_february_counts_like_datedif():
    # Excel: Am 28.02.2026 ist der 29.02.2000 noch nicht 26, am 01.03.2026 schon.
    rows = [(date(2000, 2, 29), date(2026, 2, 28), date(2026, 3, 4)), (date(2000, 2, 29), date(2026, 3, 1), date(2026, 3, 4))]
    cells = cases("Alter", "ROUND((C{r}-B{r})/365.25,0)", rows)
    hint = found(cells)["age:E2"]
    assert [shown(applied(cells, hint.fix), f"E{r}") for r in (2, 3)] == [25, 26]


def test_a_hand_made_age_that_is_right_today_is_still_worth_a_look():
    rows = [(date(1961, 5, 3), date(2025, 4, 28), date(2025, 5, 6)), (date(1980, 1, 15), date(2025, 6, 2), date(2025, 6, 2))]
    hint = found(cases("Alter", "ROUNDDOWN((C{r}-B{r})/365.25,0)", rows))["age:E2"]
    assert hint.level == "pruefen" and "Im Moment stimmt es in allen 2 Zeilen" in hint.text


def test_years_minus_years_only_for_an_age():
    hint = found(cases("Alter", "YEAR(C{r})-YEAR(B{r})"))["age:E2"]
    assert "zieht nur die Jahreszahlen ab" in hint.text and hint.level == "fehler"
    # Kalenderjahre zwischen Eintritt und Austritt: Das kann so gemeint sein.
    assert not found(cases("Jahre", "YEAR(D{r})-YEAR(C{r})"))


@pytest.mark.parametrize("formula", ['DATEDIF(B{r},C{r},"Y")', "(C{r}-B{r})/365.25", 'DATEDIF($B{r},$C{r},"Y")'])
def test_exact_or_decimal_ages_are_quiet(formula):
    assert not found(cases("Alter", formula))


def test_whole_months_over_30_days():
    rows = [(date(2025, 1, 31), date(2025, 3, 1), date(2025, 7, 30)), (date(2024, 12, 5), date(2025, 1, 2), date(2025, 6, 3))]
    cells = cases("Alter in Monaten", "INT((D{r}-B{r})/30)", rows)
    hint = found(cells)["months:E2"]
    assert hint.fix["cells"]["E2"] == '=DATEDIF(B2,D2,"M")'
    assert [shown(applied(cells, hint.fix), f"E{r}") for r in (2, 3)] == [5, 5]


def test_age_until_today_in_a_table_of_cases():
    cells = cases("Alter", 'DATEDIF(B{r},TODAY(),"Y")')
    hint = found(cells)["today:E2"]
    assert hint.level == "pruefen" and "Alter bei Eintritt gemeint" in hint.text
    assert hint.tip.startswith('Ist das Alter bei Eintritt gemeint? Dann in E2: =DATEDIF(B2;C2;"Y")')
    assert [shown(applied(cells, hint.fix), f"E{r}") for r in range(2, 8)] == AGES
    # Steht „bei Eintritt“ schon im Namen, ist HEUTE() ein Fehler.
    hint = found(cases("Alter bei Eintritt", 'DATEDIF(B{r},TODAY(),"Y")'))["today:E2"]
    assert hint.level == "fehler" and hint.tip.startswith("Korrigiert")


def test_age_until_today_in_a_list_of_people_is_meant():
    cells = cases("Alter", 'DATEDIF(B{r},TODAY(),"Y")', start=False, end=False)
    assert not found(cells)
    # „Alter heute“ sagt selbst, dass HEUTE() gemeint ist.
    assert not found(cases("Alter heute", 'DATEDIF(B{r},TODAY(),"Y")'))
    # Von Hand gerechnet: Die Korrektur behält HEUTE() und stimmt am heutigen Tag für jede Zeile.
    cells = cases("Alter", "INT((TODAY()-B{r})/365)", start=False, end=False)
    hint = found(cells)["age:C2"]
    assert hint.fix["cells"]["C2"] == '=DATEDIF(B2,TODAY(),"Y")'
    today = date.today()
    want = [today.year - b.year - ((today.month, today.day) < (b.month, b.day)) for b, _, _ in CASES]
    assert [shown(applied(cells, hint.fix), f"C{r}") for r in range(2, 8)] == want


def test_a_stay_until_today_although_every_case_has_left():
    cells = cases("Liegedauer", "TODAY()-C{r}")
    hint = found(cells)["today:E2"]
    assert hint.fix["cells"]["E2"] == "=D2-C2" and "obwohl in jeder Zeile „Austritt“ steht" in hint.text
    # Ein Fall ist noch da: Dann rechnet HEUTE() mit Absicht.
    del cells[(4, 4)]
    assert not found(filled(cells))


@pytest.mark.parametrize("head, formula, want", [
    ("Liegedauer (Tage inkl.)", "D{r}-C{r}", "=D2-C2+1"), ("Nächte", "D{r}-C{r}+1", "=D2-C2"),
    ("Arbeitstage", "D{r}-C{r}", "=NETWORKDAYS(C2,D2)"), ("Werktage", "1+D{r}-C{r}", "=NETWORKDAYS(C2,D2)"),
])
def test_days_with_or_without_the_first_day(head, formula, want):
    cells = cases(head, formula)
    hint = found(cells)["days:E2"]
    assert hint.level == "fehler" and hint.fix["cells"]["E2"] == want
    after = applied(cells, hint.fix)
    assert not found(after)
    if "NETWORKDAYS" in want:
        assert "NETTOARBEITSTAGE(C2;D2)" in hint.tip
        # 28.04.2025 (Montag) bis 06.05.2025 (Dienstag): 7 Arbeitstage.
        assert shown(after, "E2") == 7


@pytest.mark.parametrize("head, formula", [("Liegedauer (Tage inkl.)", "D{r}-C{r}+1"), ("Nächte", "D{r}-C{r}"),
                                           ("Arbeitstage", "NETWORKDAYS(C{r},D{r})"), ("Liegedauer", "D{r}-C{r}"),
                                           ("Liegedauer", "D{r}-C{r}+1")])
def test_days_said_right_are_quiet(head, formula):
    assert not found(cases(head, formula))


def run(english, cells=None):
    cells = cells or {}
    return Checked(cells, layout(Sheet("T", cells)), "T", {}).evaluator()[0].run(formel.parse(english))


def test_the_new_functions_compute_like_excel():
    holidays = {(1, 1): Cell("01.05.2025"), (2, 1): Cell("29.05.2025"), (3, 1): Cell("01.05.2025")}
    assert run("NETWORKDAYS(DATE(2025,4,28),DATE(2025,5,6))") == 7
    assert run("NETWORKDAYS(DATE(2025,4,28),DATE(2025,6,6),A1:A3)", holidays) == 28
    assert run("NETWORKDAYS(DATE(2025,5,6),DATE(2025,4,28))") == -7
    assert run("QUOTIENT(-7,2)") == -3 and run("FLOOR(7.5,2)") == 6 and run("FLOOR(-7.5,-2)") == -6
    assert run("FLOOR(-2.5,2)") == -4
    assert run("FLOOR.MATH(-7.5)") == -8 and run("FLOOR.MATH(-7.5,1,1)") == -7
    assert isinstance(run("FLOOR(7.5,-2)"), formel.Err)
    start = date(2025, 4, 28)
    assert sum((start + timedelta(i)).weekday() < 5 for i in range(9)) == 7
