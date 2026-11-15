"""Bewertungsmodelle, wie eine KI sie als Prüfungstest baut: DCF mit Sensitivität, Investitionsrechnung, Tilgungsplan."""

from test_workbook import xlsx
from test_zeilenmodell import _computed

from voiceagent import formel, pruefer
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Sheet

YEARS = "CDEFG"
STEPS = ["-0.01", "-0.005", "", "+0.005", "+0.01"]


def _dcf(**changes) -> dict:
    """Das DCF einer Testmappe (Tabelle2): Annahmen, fünf Jahre Projektion, Terminalwert, Bewertung und eine Sensitivität
    Wert je Aktie nach WACC (Zeilen) und terminalem Wachstum (Spalten). changes: Zelle → neuer Inhalt."""
    cells: dict = {"A1": "Mehrjähriges DCF-Modell mit Sensitivitätsanalyse", "A3": "Annahmen"}
    rows = [("Umsatz Basisjahr (Mio. €)", 1000), ("Umsatzwachstum Jahr 1", .08), ("Umsatzwachstum Jahr 2", .07),
            ("Umsatzwachstum Jahr 3", .06), ("Umsatzwachstum Jahr 4", .05), ("Umsatzwachstum Jahr 5", .04),
            ("EBIT-Marge", .15), ("Steuersatz", .30), ("Abschreibungen in % vom Umsatz", .04),
            ("Investitionen in % vom Umsatz", .05), ("NWC-Bindung in % der Umsatzänderung", .10), ("WACC", .09),
            ("Terminales Wachstum g", .02), ("Nettoverschuldung (Mio. €)", 500), ("Anzahl Aktien (Mio.)", 100)]
    for i, (name, value) in enumerate(rows):
        cells[f"A{4 + i}"], cells[f"B{4 + i}"] = name, value
    cells["A21"], cells["B21"] = "Projektion (Mio. €)", "Basisjahr"
    names = {22: "Umsatzwachstum (%)", 23: "Umsatz", 24: "EBIT", 25: "Steuern auf EBIT", 26: "NOPAT", 27: "Abschreibungen",
             28: "Investitionen", 29: "Veränderung NWC", 30: "Free Cash Flow", 31: "Periode", 32: "Barwert FCF"}
    for r, name in names.items():
        cells[f"A{r}"] = name
    cells["B23"] = ("=", "B4")
    for i, c in enumerate(YEARS):
        p = "BCDEFG"[i]
        cells[f"{c}21"] = f"Jahr {i + 1}"
        for r, f in ((22, f"$B${5 + i}"), (23, f"{p}23*(1+{c}22)"), (24, f"{c}23*$B$10"), (25, f"-{c}24*$B$11"),
                     (26, f"{c}24+{c}25"), (27, f"{c}23*$B$12"), (28, f"-{c}23*$B$13"), (29, f"-({c}23-{p}23)*$B$14"),
                     (30, f"{c}26+{c}27+{c}28+{c}29"), (32, f"{c}30/(1+$B$15)^{c}31")):
            cells[f"{c}{r}"] = ("=", f)
        cells[f"{c}31"] = i + 1
    for r, name, f in ((36, "FCF Jahr 6", "G30*(1+B16)"), (37, "Terminalwert (Ende Jahr 5)", "B36/(B15-B16)"),
                       (38, "Barwert Terminalwert", "B37/(1+B15)^G31"), (41, "Summe Barwerte FCF", "SUM(C32:G32)"),
                       (42, "Barwert Terminalwert", "B38"), (43, "Enterprise Value (Mio. €)", "B41+B42"),
                       (44, "Equity Value (Mio. €)", "B43-B17"), (45, "Wert je Aktie (€)", "B44/B18"),
                       (46, "Anteil Terminalwert am EV", "B42/B43")):
        cells[f"A{r}"], cells[f"B{r}"] = name, ("=", f)
    cells["A35"], cells["A40"] = "Terminalwert", "Bewertung"
    cells["A49"] = "Sensitivität: Wert je Aktie (€) – WACC (Zeilen) vs. terminales Wachstum (Spalten)"
    cells["A50"] = "WACC \\ g"
    for c, step in zip("BCDEF", STEPS):
        cells[f"{c}50"] = ("=", f"$B$16{step}")
    for r, step in zip(range(51, 56), STEPS):
        cells[f"A{r}"] = ("=", f"$B$15{step}")
        for c in "BCDEF":
            cells[f"{c}{r}"] = ("=", f"(SUMPRODUCT($C$30:$G$30/(1+$A{r})^$C$31:$G$31)+$G$30*(1+{c}$50)/($A{r}-{c}$50)"
                                     f"/(1+$A{r})^$G$31-$B$17)/$B$18")
    for address, value in changes.items():
        if value is None:
            del cells[address]
        else:
            cells[address] = ("=", value[1:]) if isinstance(value, str) and value.startswith("=") else value
    return _computed(cells)


def _hints(cells: dict, name: str = "Plan") -> dict:
    done = pruefer.Checker(None).check_file("modell.xlsx", xlsx((name, cells)))
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


def test_a_clean_dcf_with_its_sensitivity_has_nothing_to_say():
    assert _hints(_dcf()) == {}


def test_the_three_planted_errors_are_found_with_cause_and_button():
    """Testmappe, Tabelle2: B41 vergisst Jahr 5, B38 zinst mit Periode 4 ab, E27 ist die gerundete Zahl 49."""
    found = _hints(_dcf(B41="=SUM(C32:F32)", B38="=B37/(1+B15)^F31", E27=49))
    line, year, typed = found["line:B41"], found["year:B38"], found["typed:E27"]
    assert line["level"] == year["level"] == typed["level"] == "fehler"
    assert "So fehlt G32 (Barwert FCF Jahr 5)" in line["text"] and line["fix"]["german"] == "=SUMME(C32:G32)"
    assert "rechnet mit F31 (Periode Jahr 4)" in year["text"] and "Richtig wäre G31" in year["text"]
    assert year["fix"]["german"] == "=B37/(1+B15)^G31"
    assert "nur gerundet" in typed["text"] and typed["fix"]["german"] == "=E23*$B$12"
    # Die Sensitivität sagt, dass genau diese drei Fehler sie vom Modell trennen.
    sensitivity = found["sensitivity:D53"]
    assert sensitivity["level"] == "pruefen" and "Mit den Korrekturen in" in sensitivity["text"]
    assert all(cell in sensitivity["text"] for cell in ("E27", "B38", "B41"))
    # Die Schritte der Sensitivität (0,005 und 0,01) sind keine verdächtigen festen Zahlen.
    assert not any(key.startswith("constant") for key in found)


def test_a_sum_that_starts_a_year_late_misses_the_first_year():
    found = _hints(_dcf(B41="=SUM(D32:G32)"))
    assert "fängt aber schon mit Jahr 1 an" in found["line:B41"]["text"]
    assert found["line:B41"]["fix"]["german"] == "=SUMME(C32:G32)"


def test_a_terminal_value_from_the_wrong_year_is_found_where_it_starts():
    found = _hints(_dcf(B36="=F30*(1+B16)"))
    assert "rechnet mit F30 (Free Cash Flow Jahr 4)" in found["year:B36"]["text"]
    assert found["year:B36"]["fix"]["german"] == "=G30*(1+B16)"


def test_a_growth_rate_that_reads_the_list_twice():
    found = _hints(_dcf(E22="=$B$6"))
    assert "liest B6 (Umsatzwachstum Jahr 2), wie schon D22" in found["walk:E22"]["text"]
    assert found["walk:E22"]["fix"]["german"] == "=$B$7"


def test_a_relative_wacc_dragged_right_reads_empty_cells():
    found = _hints(_dcf(**{f"{c}32": f"={c}30/(1+{p}15)^{c}31" for c, p in zip("CDEFG", "BCDEF")}))
    drift = found["drift:C32"]
    assert "ohne $" in drift["text"] and "D32 liest C15" in drift["text"]
    assert set(drift["fix"]["cells"]) == {f"{c}32" for c in YEARS}
    # Die leeren C15 bis F15 sind keine eigenen Hinweise mehr.
    assert not any(key.startswith("empty") for key in found)


def test_a_typed_present_value_in_the_last_year():
    found = _hints(_dcf(G32=79.2))
    assert found["typed:G32"]["level"] == "fehler" and found["typed:G32"]["fix"]["german"] == "=G30/(1+$B$15)^G31"


def test_a_sensitivity_head_that_reads_the_wrong_assumption():
    found = _hints(_dcf(A53="=$B$14"))
    axis = found["sensaxis:A53"]
    assert "liest B14" in axis["text"] and "lesen B15 (WACC)" in axis["text"] and axis["fix"]["german"] == "=$B$15"


def test_when_model_and_sensitivity_differ_it_names_the_cell_where_they_part():
    """Das Modell zählt die Nettoverschuldung dazu, die Sensitivität zieht sie ab: Der Unterschied liegt in B44."""
    sensitivity = _hints(_dcf(B44="=B43+B17"))["sensitivity:D53"]
    assert sensitivity["level"] == "fehler" and "trennen sich in B44" in sensitivity["text"]
    assert sensitivity["fix"]["german"] == "=B43-$B$17"


def test_a_change_from_the_year_before_last_breaks_the_pattern():
    found = _hints(_dcf(G29="=-(G23-E23)*$B$14"))
    assert found["pattern:G29"]["fix"]["german"] == "=-(G23-F23)*$B$14"


def test_typed_sensitivity_heads_are_no_false_alarm():
    heads = {f"A{51 + i}": 0.08 + 0.005 * i for i in range(5)} | {f"{c}50": 0.01 + 0.005 * i for i, c in enumerate("BCDEF")}
    assert _hints(_dcf(**heads)) == {}


def _invest(**changes) -> dict:
    """Investitionsrechnung: Jahre 0 bis 5 nur als Zahlen im Kopf, Barwerte, Kapitalwert, NBW und IKV darunter."""
    cells: dict = {"A1": "Investitionsrechnung neue Produktionslinie"}
    for r, name, value in ((3, "Investition (€)", 250000), (4, "Kalkulationszins", 0.08), (5, "Absatz Jahr 1 (Stück)", 12000),
                           (6, "Absatzwachstum p.a.", 0.05), (7, "Preis je Stück (€)", 18), (8, "Variable Kosten je Stück (€)", 11),
                           (9, "Fixkosten p.a. (€)", 20000), (10, "Restwert am Ende (€)", 30000)):
        cells[f"A{r}"], cells[f"B{r}"] = name, value
    cells["A12"] = "Jahr"
    names = {13: "Absatz (Stück)", 14: "Umsatz", 15: "Variable Kosten", 16: "Fixkosten", 17: "Investition / Restwert",
             18: "Cashflow", 19: "Diskontfaktor", 20: "Barwert Cashflow", 21: "Kumulierter Cashflow"}
    for r, name in names.items():
        cells[f"A{r}"] = name
    cells["B17"] = ("=", "-B3")
    for i, c in enumerate("BCDEFG"):
        p = "ABCDEFG"[i]
        cells[f"{c}12"] = i
        if i:
            cells[f"{c}13"] = ("=", "$B$5" if i == 1 else f"{p}13*(1+$B$6)")
            cells[f"{c}14"], cells[f"{c}15"], cells[f"{c}16"] = ("=", f"{c}13*$B$7"), ("=", f"-{c}13*$B$8"), ("=", "-$B$9")
        if i == 5:
            cells[f"{c}17"] = ("=", "$B$10")
        cells[f"{c}18"], cells[f"{c}19"] = ("=", f"SUM({c}14:{c}17)"), ("=", f"1/(1+$B$4)^{c}12")
        cells[f"{c}20"], cells[f"{c}21"] = ("=", f"{c}18*{c}19"), ("=", "B18" if i == 0 else f"{p}21+{c}18")
    for r, name, f in ((23, "Kapitalwert (NPV)", "SUM(B20:G20)"), (24, "Kapitalwert mit NBW", "B18+NPV(B4,C18:G18)"),
                       (25, "Interner Zinsfuß (IRR)", "IRR(B18:G18)"), (26, "Amortisation (Jahre)", 'COUNTIF(B21:G21,"<0")')):
        cells[f"A{r}"], cells[f"B{r}"] = name, ("=", f)
    for address, value in changes.items():
        cells[address] = ("=", value[1:]) if isinstance(value, str) and value.startswith("=") else value
    return _computed(cells)


def test_numbers_0_to_5_under_a_year_word_are_the_years_of_a_row_model():
    cells = {formel.position(a): Cell(str(v[2] if isinstance(v, tuple) else v), v[1] if isinstance(v, tuple) else None)
             for a, v in _invest().items()}
    plan = layout(Sheet("Plan", cells))
    assert plan.across and plan.headers[2] == "Jahr 0" and plan.headers[7] == "Jahr 5"
    assert _hints(_invest()) == {}


def test_npv_over_year_0_discounts_the_investment_once_too_often():
    npv = _hints(_invest(B24="=NPV(B4,B18:G18)"))["npv:B24"]
    assert "Jahr 0 fällt aber heute an" in npv["text"] and npv["fix"]["german"] == "=B18+NBW(B4;C18:G18)"


def test_an_irr_without_year_0_and_a_short_total():
    found = _hints(_invest(B25="=IRR(C18:G18)", B23="=SUM(B20:F20)"))
    assert "rechnet nur mit" in found["line:B25"]["text"] and found["line:B25"]["fix"]["german"] == "=IKV(B18:G18)"
    assert found["line:B23"]["fix"]["german"] == "=SUMME(B20:G20)"


def test_a_typed_rate_in_a_repayment_plan():
    """Tilgungsplan: Eine Rate als gerundete feste Zahl zwischen Formeln."""
    cells: dict = {"A1": "Tilgungsplan Annuitätendarlehen", "A3": "Darlehensbetrag (€)", "B3": 500000, "A4": "Zinssatz p.a.",
                   "B4": 0.04, "A5": "Laufzeit (Jahre)", "B5": 10, "A6": "Jährliche Rate (€)", "B6": ("=", "PMT(B4,B5,-B3)")}
    for c, head in zip("ABCDEF", ["Jahr", "Anfangsbestand", "Zinsen", "Tilgung", "Rate", "Endbestand"]):
        cells[f"{c}9"] = head
    for i in range(10):
        r = 10 + i
        cells[f"A{r}"] = i + 1
        cells[f"B{r}"] = ("=", "$B$3" if i == 0 else f"F{r - 1}")
        for c, f in (("C", f"B{r}*$B$4"), ("E", "$B$6"), ("D", f"E{r}-C{r}"), ("F", f"B{r}-D{r}")):
            cells[f"{c}{r}"] = ("=", f)
    assert _hints(_computed(cells), "Kredit") == {}
    cells["E14"] = 61645.5
    typed = _hints(_computed(cells), "Kredit")["typedcell:E14"]
    assert typed["level"] == "fehler" and typed["fix"]["german"] == "=$B$6" and "61645,5" in typed["text"]


def test_financial_functions_compute_like_excel():
    # A1:C1 Zahlungen, A2:C2 ihre Perioden, A3:D3 Zahlungen mit der Investition vorne.
    values = {(1, 1): 100.0, (1, 2): 200.0, (1, 3): 300.0, (2, 1): 1.0, (2, 2): 2.0, (2, 3): 3.0,
              (3, 1): -1000.0, (3, 2): 300.0, (3, 3): 400.0, (3, 4): 500.0}
    calc = formel.Evaluator(values.get)

    def run(text):
        return calc.run(formel.parse(text))

    assert round(run("PMT(0.01,12,-12000)"), 2) == 1066.19
    assert round(run("IPMT(0.01,1,12,-12000)"), 2) == 120.0
    assert round(run("PV(0.05,10,-100)"), 2) == 772.17
    assert round(run("FV(0.05,10,0,-1000)"), 2) == 1628.89
    assert round(run("NPER(0.05,-100,500)"), 4) == 5.8963
    assert round(run("NPV(0.1,A1:C1)"), 4) == round(100 / 1.1 + 200 / 1.21 + 300 / 1.331, 4)
    assert round(run("IRR(A3:D3)"), 4) == 0.0890
    # SUMMENPRODUKT rechnet Bereiche Stück für Stück: Barwerte ohne Hilfszeile.
    assert round(run("SUMPRODUCT(A1:C1/(1+0.1)^A2:C2)"), 6) == round(100 / 1.1 + 200 / 1.21 + 300 / 1.331, 6)
