"""Finanzpläne und Budgets: Zeilen sind Größen (Umsatz, EBIT), Spalten sind Jahre."""

from test_workbook import xlsx

from voiceagent import formel, pruefer
from voiceagent.analyse import Sheet as Checked
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Sheet, column_letters

YEARS = "CDEFG"


def _computed(cells: dict) -> dict:
    """Die Zellen mit den Ergebnissen ihrer Formeln, wie Excel sie speichert."""
    positions = {formel.position(a): v for a, v in cells.items()}
    plain = {p: Cell(str(v[1]) if isinstance(v, tuple) else str(v), v[1] if isinstance(v, tuple) else None)
             for p, v in positions.items()}
    plan = layout(Sheet("Plan", plain))
    sheet = Checked(plain, plan, "Plan")
    _, value = sheet.evaluator(chain={p for p, c in plain.items() if c.formula})
    out = {}
    for address, v in cells.items():
        if isinstance(v, tuple):
            result = value(formel.position(address))
            out[address] = ("=", v[1], result)
        else:
            out[address] = v
    return out


def _model(**changes) -> dict:
    """Ein Finanzmodell, wie eine KI es baut: Annahmen je Jahr, GuV, Cashflow. changes: Zelle → neuer Inhalt
    (eine Zahl oder "=Formel")."""
    cells: dict = {"A1": "Finanzmodell (TEUR)", "A4": "Position", "B4": "2025A"}
    for i, c in enumerate(YEARS):
        cells[f"{c}4"] = f"{2026 + i}E"
    cells["A5"] = "Annahmen"
    assumptions = [
        (6, "Umsatzwachstum (%)", [0.08, 0.07, 0.06, 0.05, 0.04]),
        (7, "Bruttomarge (%)", [0.42, 0.425, 0.43, 0.435, 0.44]),
        (8, "Kosten (% Umsatz)", [0.15, 0.15, 0.16, 0.16, 0.17]),
        (9, "Working Capital (% Umsatz)", [0.12] * 5),
        (10, "Steuersatz (%)", [0.3] * 5),
    ]
    for r, name, values in assumptions:
        cells[f"A{r}"] = name
        for c, v in zip(YEARS, values):
            cells[f"{c}{r}"] = v
    cells["B9"] = 0.12
    cells["A11"] = "Gewinn- und Verlustrechnung"
    lines = {
        12: ("Umsatz", "{p}12*(1+{c}6)"),
        13: ("Herstellungskosten", "-{c}12*(1-{c}7)"),
        14: ("Bruttogewinn", "{c}12+{c}13"),
        15: ("Kosten", "-{c}12*{c}8"),
        16: ("EBIT", "SUM({c}14:{c}15)"),
        17: ("Steuern", "-MAX(0,{c}16*{c}10)"),
        18: ("Jahresüberschuss", "{c}16+{c}17"),
        21: ("Jahresüberschuss", "{c}18"),
        22: ("Veränderung Working Capital", "-({c}12*{c}9-{p}12*{p}9)"),
        23: ("Operativer Cashflow", "SUM({c}21:{c}22)"),
        24: ("Liquide Mittel Ende", "{p}24+{c}23"),
    }
    cells["A20"] = "Cashflow"
    for r, (name, formula) in lines.items():
        cells[f"A{r}"] = name
        for i, c in enumerate(YEARS):
            cells[f"{c}{r}"] = ("=", formula.format(c=c, p="B" if i == 0 else YEARS[i - 1]))
    cells["B12"], cells["B24"] = 10000, 2000
    for address, value in changes.items():
        if value == "":
            del cells[address]
        else:
            cells[address] = ("=", value[1:]) if isinstance(value, str) and value.startswith("=") else value
    return _computed(cells)


def _hints(cells: dict) -> dict:
    done = pruefer.Checker(None).check_file("plan.xlsx", xlsx(("Plan", cells)))
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


def test_a_financial_plan_is_a_row_model_without_false_alarms():
    cells = {formel.position(a): Cell(str(v[2]) if isinstance(v, tuple) else str(v), v[1] if isinstance(v, tuple) else None)
             for a, v in _model().items()}
    assert layout(Sheet("Plan", cells)).across
    # Keine Lücke in der Summe des Cashflows, kein doppelter „Jahresüberschuss“, keine Ausreißer je Jahr.
    assert _hints(_model()) == {}


def test_a_typed_number_in_one_year_gets_the_formula_of_the_other_years():
    hint = _hints(_model(E12=12000))["typed:E12"]
    assert hint["level"] == "fehler"
    assert hint["text"].startswith("E12 (Umsatz 2028E) ist eine feste Zahl, 12000. Die übrigen Jahre in C12:G12 rechnen mit "
                                   "einer Formel, hier wäre das =Umsatz 2027E * (1 + Umsatzwachstum (%)) mit ")
    assert hint["fix"]["cells"] == {"E12": "=D12*(1+E6)"}
    assert hint["fix"]["was"] == {"E12": ""}


def test_a_year_pasted_as_values_is_one_hint_with_one_button():
    computed = _model()
    pasted = {f"E{r}": computed[f"E{r}"][2] for r in (12, 13, 14, 15, 16, 17, 18, 21, 22, 23, 24)}
    hints = _hints(_model(**pasted))
    assert list(hints) == ["typed:E12"]
    hint = hints["typed:E12"]
    assert hint["level"] == "pruefen"
    assert "In 2028E (E12:E24) stehen 11 feste Zahlen" in hint["text"]
    assert "Die festen Zahlen rechnen aber nicht mit" in hint["text"]
    assert hint["fix"]["cells"]["E23"] == "=SUM(E21:E22)"
    assert len(hint["fix"]["cells"]) == 11


def test_an_assumption_held_with_dollars_reads_its_own_year():
    frozen = {f"{c}12": f"={'B' if i == 0 else YEARS[i - 1]}12*(1+$C$6)" for i, c in enumerate(YEARS)}
    hint = _hints(_model(**frozen))["frozen:D12"]
    assert hint["level"] == "fehler"
    assert hint["text"] == ("C12:G12 (Umsatz) rechnen in jedem Jahr mit Umsatzwachstum (%) 2026E ($C$6). Die Werte der "
                            "übrigen Jahre (D6:G6) liest keine Formel.")
    # Auch 2026E selbst mit C$6, sonst rechnete es danach anders als die übrigen Jahre.
    assert hint["fix"]["cells"] == {"C12": "=B12*(1+C$6)", "D12": "=C12*(1+D$6)", "E12": "=D12*(1+E$6)",
                                    "F12": "=E12*(1+F$6)", "G12": "=F12*(1+G$6)"}


def test_one_assumption_for_all_years_is_no_frozen_assumption():
    # Nur 2026E hat einen Wert: Dann ist $C$6 die eine Annahme für alle Jahre.
    changes = {f"{c}12": f"={'B' if i == 0 else YEARS[i - 1]}12*(1+$C$6)" for i, c in enumerate(YEARS)}
    changes |= {f"{c}6": "" for c in YEARS[1:]}
    assert not any(k.startswith("frozen:") for k in _hints(_model(**changes)))


def test_a_sum_that_misses_a_row_in_every_year_is_one_hint():
    short = {f"{c}23": f"=SUM({c}21:{c}21)" for c in YEARS}
    hint = _hints(_model(**short))["gap:C23"]
    assert hint["level"] == "fehler"
    assert hint["text"] == ("Die Summen in C23:G23 (Operativer Cashflow) zählen Veränderung Working Capital nicht mit: "
                            "=SUMME(Jahresüberschuss bis Jahresüberschuss) statt =SUMME(Jahresüberschuss bis Veränderung "
                            "Working Capital). C23 zeigt darum 2041,2 statt 1945,2.")
    assert hint["fix"]["cells"] == {f"{c}23": f"=SUM({c}21:{c}22)" for c in YEARS}


def test_a_sum_that_misses_a_row_in_one_year_is_a_pattern_break():
    hints = _hints(_model(E23="=SUM(E21:E21)"))
    assert list(hints) == ["pattern:E23"]
    assert "So fehlt E22 (Veränderung Working Capital 2028E) mit" in hints["pattern:E23"]["text"]
    assert hints["pattern:E23"]["fix"]["cells"] == {"E23": "=SUM(E21:E22)"}


def test_most_years_short_do_not_make_the_full_sum_the_error():
    short = {f"{c}23": f"=SUM({c}21:{c}21)" for c in "CDFG"}
    hints = _hints(_model(**short))
    assert list(hints) == ["gap:C23"]
    assert hints["gap:C23"]["text"].startswith("Die Summen in C23, D23, F23 und G23 (Operativer Cashflow)")


def test_a_wrong_assumption_row_in_one_year_is_a_pattern_break():
    hint = _hints(_model(F15="=-F12*F7"))["pattern:F15"]
    assert hint["text"].startswith("F15 (Kosten 2029E) rechnet anders als die übrigen 4 Formeln in C15:G15: "
                                   "=-Umsatz * Bruttomarge (%) statt =-Umsatz * Kosten (% Umsatz).")
    assert hint["fix"]["cells"] == {"F15": "=-F12*F8"}


def test_a_total_column_for_the_forecast_years_and_its_growth_rate():
    cells = _model()
    cells["H4"], cells["I4"] = "Summe 2026-2030", "CAGR"
    for r in (12, 13, 14, 15, 16, 17, 18):
        cells[f"H{r}"] = ("=", f"SUM(C{r}:G{r})")
    cells["I12"] = ("=", "(G12/C12)^(1/4)-1")
    cells = _computed({a: (v[0], v[1]) if isinstance(v, tuple) else v for a, v in cells.items()})
    # Die Summe zählt 2025A nicht mit, das sagt ihr Name. CAGR rechnet anders als die Jahre, das gehört so.
    assert _hints(cells) == {}


def _budget() -> dict:
    cells: dict = {"A2": "Position"}
    for i, c in enumerate("BCDE"):
        cells[f"{c}2"] = f"Budget {2026 + i}"
    rows = [("Erlöse", [1200, 1260, 1320, 1390]), ("Raumkosten", [100, 100, 105, 105]),
            ("Personal", [600, 620, 640, 660]), ("Sachkosten", [150, 155, 160, 165])]
    for k, (name, values) in enumerate(rows):
        cells[f"A{3 + k}"] = name
        for c, v in zip("BCDE", values):
            cells[f"{c}{3 + k}"] = v
    cells["A7"], cells["A8"] = "Summe Kosten", "Ergebnis"
    for c in "BCDE":
        cells[f"{c}7"], cells[f"{c}8"] = ("=", f"SUM({c}4:{c}6)"), ("=", f"{c}3-{c}7")
    return _computed(cells)


def test_costs_below_revenue_in_a_budget_are_not_its_parts():
    assert _hints(_budget()) == {}


def test_a_cost_sum_that_misses_a_row_in_every_year():
    cells = _budget()
    for c in "BCDE":
        cells[f"{c}7"] = ("=", f"SUM({c}4:{c}5)")
    hint = _hints(_computed({a: (v[0], v[1]) if isinstance(v, tuple) else v for a, v in cells.items()}))["gap:B7"]
    assert "zählen Sachkosten nicht mit" in hint["text"]
    assert hint["fix"]["cells"] == {f"{c}7": f"=SUM({c}4:{c}6)" for c in "BCDE"}


def test_an_amount_like_a_year_inside_a_monthly_budget_is_an_amount():
    months = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun"]
    cells: dict = {"A3": "Position", "H3": "Summe"}
    for i, m in enumerate(months):
        cells[f"{column_letters(2 + i)}3"] = m
    values = {4: [3430, 7060, 6570, 2330, 4780, 7180], 5: [5370, 5040, 1990, 6900, 5550, 2370],
              6: [3640, 7880, 5460, 8970, 7410, 4080]}
    for r, row in values.items():
        cells[f"A{r}"] = ["Miete", "IT", "Reisen"][r - 4]
        for i, v in enumerate(row):
            cells[f"{column_letters(2 + i)}{r}"] = v
        cells[f"H{r}"] = ("=", f"SUM(B{r}:G{r})")
    cells["A7"] = "Total"
    for i in range(7):
        c = column_letters(2 + i)
        cells[f"{c}7"] = ("=", f"SUM({c}4:{c}6)")
    # 1990 Franken im März sind ein Betrag, kein Kopf „1990“, der die Zeile in zwei Teile schneidet.
    assert _hints(_computed(cells)) == {}


def test_three_errors_say_what_the_cell_shows_and_what_is_right():
    # Wie in einem Test: die Steuern 2028E vom Ergebnis des Vorjahrs, eine eingetippte Zahl, eine Summe zu kurz.
    hints = _hints(_model(E17="=-MAX(0,D16*E10)", F15=-2000, G23="=SUM(G21:G21)"))
    assert sorted(hints) == ["pattern:E17", "pattern:G23", "typed:F15"]
    steuern = hints["pattern:E17"]["text"]
    assert "=-MAX(0; EBIT 2027E * Steuersatz (%)) statt =-MAX(0; EBIT * Steuersatz (%))." in steuern
    assert "E17 zeigt darum " in steuern and "(so viel wie D17)" in steuern
    summe = hints["pattern:G23"]["text"]
    assert "So fehlt G22 (Veränderung Working Capital 2030E) mit " in summe
    assert "G23 zeigt darum " in summe
    assert hints["pattern:G23"]["fix"]["cells"] == {"G23": "=SUM(G21:G22)"}
    assert hints["typed:F15"]["fix"]["cells"] == {"F15": "=-F12*F8"}
