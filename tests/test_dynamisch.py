"""BEREICH.VERSCHIEBEN, INDIREKT und ADRESSE: Bezüge, die erst beim Rechnen feststehen. Der Prüfer rechnet sie wie Excel
nach (gegen LibreOffice an 1700 Zufallsformeln geprüft), erklärt #BEZUG! daraus und hält Kennzahlen unter einer Summe,
die nur einen Wert aus der Spalte holen, nicht für fehlende Zeilen."""

from test_workbook import xlsx

from voiceagent import formel, pruefer


def _calc():
    grid = {(r, 2): float(r * 10) for r in range(1, 20)}
    grid.update({(1, 6): 3.0, (1, 1): "B5"})

    def other(sheet, pos):
        if sheet.casefold() == "mein satz":
            return 1000.0 + pos[0]
        raise formel.Unknown("Blatt")

    return formel.Evaluator(grid.get, other=other)


def _run(text: str):
    try:
        return _calc().run(formel.parse(text))
    except formel.Unknown:
        return "unbekannt"


def test_offset_indirect_and_address_compute_like_excel():
    assert _run("SUM(OFFSET(B2,0,0,3,1))") == 90.0
    assert _run("SUM(OFFSET(B10,1-$F$1,0,$F$1,1))") == 270.0
    assert _run("OFFSET(B2,1.9,0)") == 30.0
    assert _run("OFFSET(OFFSET(B2,1,0),1,0)") == 40.0
    assert _run("VLOOKUP(30,OFFSET(B1,0,0,10,1),1,FALSE)") == 30.0
    assert _run("ROWS(OFFSET(B2,0,0,4,2))+COLUMNS(OFFSET(B2,0,0,4,2))") == 6.0
    assert _run('INDIRECT("B"&5)') == 50.0 and _run("INDIRECT(A1)") == 50.0
    assert _run('SUM(INDIRECT("b2:b4"))') == 90.0
    assert _run("INDIRECT(\"'Mein Satz'!B7\")") == 1007.0
    assert _run("INDIRECT(ADDRESS(5,2))") == 50.0
    assert _run('ADDRESS(1,1,1,TRUE,"Mein Satz")') == "'Mein Satz'!$A$1" and _run("ADDRESS(5,28,2)") == "AB$5"
    # Über den Rand des Blatts: #BEZUG! wie in Excel.
    assert _run("OFFSET(B2,-5,0)") == formel.Err("#REF!") and _run('INDIRECT("")') == formel.Err("#REF!")
    # Was Excel und LibreOffice verschieden rechnen oder ein Name der Mappe ist, rechnet er nicht nach.
    assert _run("OFFSET(B2,0,0,0,1)") == _run('INDIRECT("R1C1",FALSE)') == _run('INDIRECT("MwSt")') == "unbekannt"


def _umsatz(**changes) -> dict:
    cells = {"A1": "Monat", "B1": "Umsatz", "C1": "Rollierend 3", "F1": 3.0, "G1": "Fenster"}
    values = [100, 120, 90, 130, 150, 110, 170, 160, 140, 180, 200, 190]
    for i, v in enumerate(values):
        r = i + 2
        cells.update({f"A{r}": f"M{i + 1}", f"B{r}": float(v)})
        if i >= 2:
            cells[f"C{r}"] = ("=", f"SUM(OFFSET(B{r},1-$F$1,0,$F$1,1))", float(sum(values[i - 2:i + 1])))
    cells.update({"A15": "Summe", "B15": ("=", "SUM(B2:B13)", float(sum(values)))})
    cells.update(changes)
    return cells


def _hints(cells: dict) -> dict:
    done = pruefer.Checker(None).check_file("umsatz.xlsx", xlsx(("Umsatz", cells)))
    assert done["state"] == "ok", done
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


def test_a_rolling_sum_with_offset_is_recomputed():
    assert _hints(_umsatz()) == {}
    hint = _hints(_umsatz(C9=("=", "SUM(OFFSET(B9,1-$F$1,0,$F$1,1))", 999.0)))["recalc:C9"]
    assert "nachgerechnet ergibt" in hint["text"] and "440" in hint["text"]


def test_a_shorter_window_in_one_row_is_a_pattern_break_with_a_fix():
    """Eine Summe, die dieselben Zellen liest wie ihre Nachbarn, nur anders verrechnet, ist keine Zwischensumme."""
    hint = _hints(_umsatz(C9=("=", "SUM(OFFSET(B9,1-$F$1,0,$F$1-1,1))", 280.0)))["pattern:C9"]
    assert "rechnet anders als die übrigen 9 Formeln in C4:C13" in hint["text"]
    assert hint["fix"]["cells"] == {"C9": "=SUM(OFFSET(B9,1-$F$1,0,$F$1,1))"}
    # Eine Zwischensumme mitten in der Spalte zählt andere Zellen: Sie bleibt still.
    found = _hints(_umsatz(C9=("=", "SUM(C4:C8)", 2000.0)))
    assert "pattern:C9" not in found


def test_figures_below_the_total_that_only_pick_a_value_are_no_missing_rows():
    for formula, value in (("B13", 190.0), ("INDEX(B2:B13,12)", 190.0), ('INDIRECT("B"&13)', 190.0),
                           ("OFFSET(B2,11,0)", 190.0)):
        assert _hints(_umsatz(A16="Letzter", B16=("=", formula, value))) == {}, formula


def test_a_ref_error_from_offset_or_indirect_names_its_cause():
    found = _hints(_umsatz(A17="Vor 3 Monaten", B17=("=", "OFFSET(B3,-3,0)", "#REF!"),
                           A18="Plan", B18=("=", "INDIRECT(\"'Plan 2025'!B2\")", "#REF!"),
                           A19="Kaputt", B19=("=", 'INDIRECT("B"&"x")', "#REF!")))
    assert "BEREICH.VERSCHIEBEN(B3;-3;0) zeigt 3 Zeilen über B3, über den oberen Rand des Blatts hinaus" in found["error:B17"]["text"]
    assert "zeigt auf das Blatt „Plan 2025“, und das gibt es in dieser Mappe nicht" in found["error:B18"]["text"]
    assert "bekommt „Bx“, und das ist keine Zelle und kein Bereich" in found["error:B19"]["text"]


def test_a_figure_name_is_not_its_own_group():
    """Kennzahlen gleich neben der Liste, ohne eigenen Kopf: „Kennzahl 2“ in D2 meint nicht die Zeilen, in denen
    „Kennzahl 2“ steht, denn das ist nur die eigene Zeile."""
    cells = {"A1": "Fall", "B1": "Wert", "C1": "Menge"}
    for r in range(2, 16):
        cells.update({f"A{r}": f"F{r:03d}", f"B{r}": float(10 * r + r % 3), f"C{r}": float(r % 7 + 1)})
    calc = formel.Evaluator(lambda p: float(10 * p[0] + p[0] % 3) if p[1] == 2 and 2 <= p[0] <= 15 else None)
    formulas = ["SUM(B2:B6)", "B9", "B12*2", "AVERAGE(B2:B15)", "B3", "SUM(B4:B5)", "B12-B2", "MIN(B2:B15)",
                "B7", "COUNT(B2:B15)"]
    for r, formula in enumerate(formulas, start=1):
        cells.update({f"D{r}": f"Kennzahl {r}", f"E{r}": ("=", formula, calc.run(formel.parse(formula)))})
    assert not any(k.startswith("label:") for k in _hints(cells))
