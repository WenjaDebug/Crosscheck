"""Große Spalten: Formeln in jeder Zeile über ganze Spalten (Duplikate, erstes Vorkommen, Laufsummen, SUMMEWENN je
Station) sind schnell und brauchen wenig Speicher, und sie rechnen genau so wie Zelle für Zelle."""

import math
import random
import time

import pytest

from voiceagent import analyse, formel, pruefer
from voiceagent.analyse import Columns, Readers, Refs
from voiceagent.formel import Err, Evaluator, parse
from voiceagent.summen import _counted, _weights
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Sheet, column_letters

STATIONS = ["K1", "K2", "C1", "M1"]


def _text(value):
    if isinstance(value, bool):
        return "WAHR" if value else "FALSCH"
    if isinstance(value, float):
        return repr(round(value, 6)).replace(".", ",").removesuffix(",0")
    return str(value)


def clinic(n, seed=1):
    """Fälle mit Duplikat-Spalte, erstem Vorkommen, Laufsumme, SUMMEWENN und ZÄHLENWENNS je Zeile, die Werte so,
    wie Excel sie zeigt."""
    rng = random.Random(seed)
    last = n + 1
    cases = [f"F{100000 + i}" for i in range(n)]
    cases[n // 2] = cases[3]  # ein doppelter Fall
    patients = [f"P{rng.randint(1, n // 3)}" for _ in range(n)]
    stations = [rng.choice(STATIONS) for _ in range(n)]
    costs = [round(rng.uniform(500, 9000), 2) for _ in range(n)]
    cells = {(1, c): Cell(h) for c, h in enumerate(["Fall-Nr", "Patient", "Station", "Kosten", "Doppelt", "Erstes",
                                                     "Laufsumme", "Kosten Station", "Fälle Patient Station"], 1)}
    running = 0.0
    for i in range(n):
        r = i + 2
        running += costs[i]
        values = [cases[i], patients[i], stations[i], costs[i]]
        for c, v in enumerate(values, 1):
            cells[(r, c)] = Cell(_text(v))
        formulas = [
            (f"COUNTIF($A$2:$A${last},A{r})>1", cases.count(cases[i]) > 1),
            (f"COUNTIF($B$2:B{r},B{r})=1", patients[: i + 1].count(patients[i]) == 1),
            (f"SUM($D$2:D{r})", running),
            (f"SUMIF($C$2:$C${last},C{r},$D$2:$D${last})", math.fsum(k for s, k in zip(stations, costs) if s == stations[i])),
            (f"COUNTIFS($B$2:$B${last},B{r},$C$2:$C${last},C{r})",
             float(sum(p == patients[i] and s == stations[i] for p, s in zip(patients, stations)))),
        ]
        for c, (f, v) in enumerate(formulas, 5):
            cells[(r, c)] = Cell(_text(v), f)
    return cells


def test_thousands_of_rows_with_formulas_over_the_whole_column_are_checked_quickly():
    # Vorher: 4000 Zeilen ZÄHLENWENN($A$2:$A$4001;A2) brauchten fast eine Minute und 2,6 GB.
    cells = clinic(4000)
    start = time.monotonic()
    done = pruefer.sheet_result("B", "Fälle", {"Fälle": cells})
    assert time.monotonic() - start < 30
    keys = {h["key"] for h in done["hints"]}
    # Nichts falsch nachgerechnet, und das erste Vorkommen ZÄHLENWENN($B$2:B3;B3) ist kein zu kurzer Bereich.
    assert not [k for k in keys if k.startswith(("recalc:", "partial:", "range:", "group:"))], keys
    sheet = analyse.Sheet(cells, layout(Sheet("Fälle", cells)), "Fälle")
    assert isinstance(sheet.precedents[(2, 5)], Refs) and len(sheet.precedents[(2, 5)]) == 4000


def test_first_occurrence_with_a_growing_range_is_no_short_range():
    cells = {(1, 1): Cell("Fall-Nr"), (1, 2): Cell("Patient"), (1, 3): Cell("Erstes")}
    patients = ["P1", "P2", "P1", "P3", "P2", "P4", "P5", "P1", "P6", "P7", "P3", "P8"]
    for i, p in enumerate(patients):
        r = i + 2
        cells[(r, 1)], cells[(r, 2)] = Cell(f"F{i}"), Cell(p)
        cells[(r, 3)] = Cell(_text(patients[: i + 1].count(p) == 1), f"COUNTIF($B$2:B{r},B{r})=1")
    done = pruefer.sheet_result("B", "T", {"T": cells})
    assert not [h for h in done["hints"] if h["key"].startswith(("partial:", "range:", "recalc:"))], done["hints"]


@pytest.mark.parametrize("seed", range(6))
def test_lookup_indexes_and_running_sums_compute_like_cell_by_cell(seed, monkeypatch):
    rng = random.Random(seed)
    n = 60
    values = {}
    for r in range(1, n + 1):
        values[(r, 1)] = rng.choice(["a", "A", "b", " b", "c", 1.0, 2.0, True, False, None, "", "1", Err("#N/A")])
        values[(r, 2)] = rng.choice(["x", "y", "X", 3.0, None])
        values[(r, 3)] = rng.choice([1.5, -2.25, 0.1, 0.2, 0.3, -0.0, 1e16, 7.0, None, "t", True, Err("#DIV/0!")])
    formulas = []
    for _ in range(80):
        r1 = rng.choice([1, 1, rng.randint(1, n)])
        r2 = rng.randint(r1, n)
        a, b, c = (f"{col}{r1}:{col}{r2}" for col in "ABC")
        key = rng.choice(['"a"', '"=b"', "1", "TRUE", '"1"', '""', "A5", "B7", '"x"', '"a*"', '">1"'])
        formulas += [f"COUNTIF({a},{key})", f"SUMIF({a},{key},{c})", f"AVERAGEIF({a},{key},{c})",
                     f'COUNTIFS({a},{key},{b},"x")', f"SUMIFS({c},{a},{key},{b},B3)", f"AVERAGEIFS({c},{b},{key})",
                     f"SUM({c})", f"AVERAGE({c})", f"COUNT({c})", f"SUM({a})", f"MINIFS({c},{a},{key})",
                     f'MAXIFS({c},{b},"x",{a},{key})']

    def run(index):
        monkeypatch.setattr(formel, "MATCH_INDEX", index)
        calc = Evaluator(lambda p: values.get(p), static=bool(seed % 2))
        out = []
        for f in formulas:
            try:
                out.append(repr(calc.run(parse(f))))
            except formel.Unknown as err:
                out.append(f"Unknown {err}")
        return out

    assert run(0) == run(10**9)


def test_prefix_sums_round_exactly_like_fsum():
    values = {(r, 1): v for r, v in enumerate([0.1, 0.2, 0.3, -0.6, 1e16, 1.0, -1e16, 5e-324, -0.0, 2.5] * 5, 1)}
    calc = Evaluator(lambda p: values.get(p))
    for r2 in range(34, 51):
        for r1 in (1, 2):
            want = math.fsum(v for (r, _), v in values.items() if r1 <= r <= r2)
            assert repr(calc.run(parse(f"SUM(A{r1}:A{r2})"))) == repr(want)
    # Ein Fehler im Bereich kommt durch, wie in Excel, und nur dort.
    values[(40, 1)] = Err("#N/A")
    calc = Evaluator(lambda p: values.get(p))
    assert calc.run(parse("SUM(A1:A50)")) == Err("#N/A") and calc.run(parse("AVERAGE(A2:A45)")) == Err("#N/A")
    assert repr(calc.run(parse("SUM(A1:A39)"))) == repr(math.fsum(v for (r, _), v in values.items() if r <= 39))
    assert calc.run(parse("COUNT(A1:A39)")) == 39


def test_wide_references_read_the_same_cells_as_one_by_one(monkeypatch):
    rng = random.Random(4)
    cells = {(1, 1): Cell("Wert"), (1, 2): Cell("Summe")}
    for r in range(2, 40):
        if rng.random() < 0.8:
            cells[(r, 1)] = Cell(str(r))
        cells[(r, 3)] = Cell("0", "SUM($A$2:A{0})+A{0}+B1+$A$3".format(r) if r > 2 else "SUM(A2:A2)+B2")
        if r % 7 == 0:
            cells[(r, 2)] = Cell("Kennzahl")  # ihr Wort links nennt eine Kennzahl, die andere Zeilen liest
        if r % 5 == 0:
            cells[(r, 4)] = Cell("0", f"SUM(A2:C{r},A{r}:A39,Z99)")
    plan = layout(Sheet("T", cells))

    def build(wide):
        monkeypatch.setattr(analyse, "WIDE", wide)
        return analyse.Sheet(cells, plan, "T")

    new, old = build(0), build(10**12)
    assert isinstance(new.dependents, Readers) and new.columns is not None
    for pos in new.nodes:
        refs = new.precedents[pos]
        assert isinstance(refs, Refs)
        listed = list(refs)
        assert len(listed) == len(set(listed)) == len(refs) and set(listed) == set(old.precedents[pos])
        assert all(p in refs for p in old.precedents[pos])
        assert new.figure_label(pos) == old.figure_label(pos) and analyse._chain(new, pos) == analyse._chain(old, pos)
        assert ((99, 26) in refs) == ("Z99" in cells[pos].formula)
        assert set(new.formula_inputs(pos)) == {p for p in old.precedents[pos] if p in old.nodes}
    for pos in {*cells, (1, 2), (99, 26), (50, 1)}:
        assert set(new.dependents.get(pos, ())) == set(old.dependents.get(pos, ()))
    assert set(new.read_cells()) == {p for before in old.precedents.values() for p in before}


def test_columns_find_the_cells_inside_an_area():
    found = Columns([(2, 1), (5, 1), (3, 2), (9, 2), (4, 4)])
    assert list(found.inside((3, 1, 9, 2))) == [(5, 1), (3, 2), (9, 2)] and found.count((1, 1, 100, 100)) == 5
    assert found.count((6, 3, 8, 3)) == 0


@pytest.mark.parametrize("formula", [
    "SUM(A1:A6)", "SUM(A1:A6)-SUM(A2:A6)+A4", "ROUND(SUM(A1:B4),2)", "-A1+A1+A2", "SUM(A1:A3,A2,7)-3",
    "A1+B2-C3", "SUM(A1:C6)-A1-A1", "SUM(B:B)",
])
def test_lazy_sum_weights_answer_like_the_full_count(formula):
    cells = {(r, c): Cell(str(r * c)) for r in range(1, 7) for c in range(1, 4)}
    cells[(3, 2)] = Cell("Text")
    sheet = analyse.Sheet(cells, layout(Sheet("T", cells)), "T")
    node = parse(formula)
    lazy, full = _weights(sheet, node), _counted(sheet, node)
    assert (lazy is None) == (full is None)
    for pos in [*cells, (9, 9)]:
        assert lazy.get(pos) == full.get(pos) and lazy.get(pos, 0) == full.get(pos, 0)
    assert dict(lazy.items()) == full and list(lazy) == list(full) and len(lazy) == len(full)


def running(n, broken=None, end=None, first=2):
    """Kosten in B, Laufsumme =SUMME($B$2:B{r}) in C. broken: diese Zeile endet in Zeile end."""
    cells = {(1, 1): Cell("Fall"), (1, 2): Cell("Kosten"), (1, 3): Cell("Laufsumme")}
    costs = {r: 100 + 3 * r for r in range(2, n + 2)}
    for r in range(2, n + 2):
        cells[(r, 1)], cells[(r, 2)] = Cell(f"F{r}"), Cell(str(costs[r]))
        stop = end if r == broken else r
        start = first if r >= first else 2
        cells[(r, 3)] = Cell(str(sum(costs[k] for k in range(start, stop + 1))), f"SUM($B${start}:B{stop})")
    return cells


def hints_of(cells):
    return {h["key"]: h for h in pruefer.sheet_result("B", "T", {"T": cells})["hints"]}


def test_a_running_sum_that_skips_its_own_row_is_a_copy_error():
    # Vorher galt jede SUMME in einer Spalte als Zwischensumme, auch eine Laufsumme, die ihre eigene Zeile vergisst.
    hint = hints_of(running(400, broken=200, end=199))["pattern:C200"]
    assert hint["level"] == "fehler" and "übrigen 399 Formeln in C2:C401" in hint["text"], hint
    # Die echten Formeln: „Kosten F2 bis F199“ las man als Zellen F2 bis F199, die Fall-Nr F200 als Zelle F200.
    assert hint["text"].startswith("C200 (Laufsumme, Fall F200) rechnet anders")
    assert ("=SUMME($B$2:B199) statt =SUMME($B$2:B200). So fehlt B200 (Kosten, Fall F200) mit 700. "
            "C200 zeigt darum 79497 (so viel wie C199) statt 80197.") in hint["text"]
    want = sum(100 + 3 * r for r in range(2, 201))
    assert hint["fix"]["cells"] == {"C200": "=SUM($B$2:B200)"} and f"nachgerechnet {want}." in hint["tip"]
    fixed = running(400)
    assert not [k for k in hints_of(fixed) if k.startswith("pattern:")]
    # Zu weit: Sie zählt die nächste Zeile schon mit.
    hint = hints_of(running(40, broken=20, end=22))["pattern:C20"]
    assert "So zählt sie B21, B22 zu viel." in hint["text"]
    # Kurz vor dem Ende: Sie soll in ihrer Zeile enden. Vorher schlug Einsetzen vor, sie bis ans Ende zu ziehen.
    hints = hints_of(running(40, broken=39, end=40))
    assert hints["pattern:C39"]["fix"]["cells"] == {"C39": "=SUM($B$2:B39)"} and "range:C39" not in hints
    hints = hints_of(running(40, broken=30, end=29))
    assert hints["pattern:C30"]["fix"]["cells"] == {"C30": "=SUM($B$2:B30)"} and "range:C30" not in hints
    # Die letzte Zeile, die eine Zeile zu früh endet: Sie soll bis zur letzten Zahl reichen.
    hints = hints_of(running(40, broken=41, end=40))
    assert hints["range:C41"]["fix"]["cells"] == {"C41": "=SUM($B$2:B41)"}, hints


def test_subtotals_and_a_running_sum_that_starts_anew_stay_quiet():
    # Ein neues Jahr fängt die Laufsumme neu an: Das ist kein Kopierfehler.
    assert not [k for k in hints_of(running(16, first=14)) if k.startswith("pattern:")]
    # Fängt sie früh neu an, rechnen die ersten Zeilen anders als die meisten: Auch sie sind kein Fehler.
    assert not [k for k in hints_of(running(30, first=5)) if k.startswith("pattern:")]
    # Zeilensummen =SUMME(B:D) mit einer Zwischensumme und der Summe unten, beide über die eigene Spalte.
    cells = {(1, c): Cell(h) for c, h in enumerate(["Monat", "A", "B", "C", "Total"], 1)}
    for r in range(2, 16):
        cells[(r, 1)] = Cell(f"M{r}")
        for c in (2, 3, 4):
            cells[(r, c)] = Cell(str(r * c))
        cells[(r, 5)] = Cell(str(r * 9), f"SUM(B{r}:D{r})")
    cells[(8, 1)], cells[(8, 5)] = Cell("Halbjahr"), Cell(str(sum(r * 9 for r in range(2, 8))), "SUM(E2:E7)")
    cells[(16, 1)], cells[(16, 5)] = Cell("Total"), Cell(str(sum(r * 9 for r in range(2, 16) if r != 8)), "SUM(E2:E7,E9:E15)")
    for c in (2, 3, 4):
        cells[(8, c)] = Cell(str(sum(r * c for r in range(2, 8))), f"SUM({column_letters(c)}2:{column_letters(c)}7)")
    assert not [k for k in hints_of(cells) if k.startswith("pattern:")]
    # Eine Zeilensumme, der eine Spalte fehlt, fällt dagegen auf (als Summe ihrer Teile, die einen auslässt).
    cells[(11, 5)], cells[(16, 5)] = Cell(str(11 * 5), "SUM(B11:C11)"), Cell("955", "SUM(E2:E7,E9:E15)")
    found = [h for k, h in hints_of(cells).items() if k.endswith(":E11")]
    assert len(found) == 1 and found[0]["level"] == "fehler" and "D11" in found[0]["text"], found
