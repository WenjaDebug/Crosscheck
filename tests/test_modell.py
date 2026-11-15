"""Modelle mit mehreren Tabellen untereinander, wie eine KI sie als Prüfungstest baut: eine Rating-Migration mit
Markov-Kette (Testmappe, Tabelle3). Übergangsmatrix, Startverteilung, Verteilung je Jahr und Verlust je Jahr stehen
auf einem Blatt, jede mit eigener Kopfzeile."""

from test_dcf import _hints
from test_zeilenmodell import _computed

from voiceagent import ansicht, formel
from voiceagent.workbook import layout, others
from voiceagent.xlsx import Cell, Sheet

MATRIX = [("A", [0.90, 0.08, 0.02, 0.0]), ("B", [0.05, 0.85, 0.08, 0.02]), ("C", [0.01, 0.10, 0.80, 0.09]),
          ("D (Ausfall)", [0.0, 0.0, 0.0, 1.0])]


def _markov(**changes) -> dict:
    cells: dict = {"A1": "Rating-Migration mit Markov-Kette – Kreditportfolio (Prüfungstest, Modell 3)",
                   "A2": "Blau = Eingabe, Schwarz = Formel. Beträge in TEUR. Wahrscheinlichkeiten je Jahr.",
                   "A4": "Annahmen", "A5": "Portfoliowert Basisjahr (TEUR)", "B5": 50000.0,
                   "A6": "Verlustquote bei Ausfall (LGD)", "B6": 0.45, "A7": "Diskontsatz p.a.", "B7": 0.05,
                   "A9": "Übergangsmatrix (Zeile = Ausgangsrating, Spalte = Zielrating)"}
    for c, h in zip("ABCDEF", ["Von / Nach", "A", "B", "C", "D (Ausfall)", "Zeilensumme"]):
        cells[f"{c}10"] = h
    for i, (name, row) in enumerate(MATRIX):
        r = 11 + i
        cells[f"A{r}"] = name
        for c, v in zip("BCDE", row):
            cells[f"{c}{r}"] = v
        cells[f"F{r}"] = ("=", f"SUM(B{r}:E{r})")
    cells["A16"] = "Startverteilung"
    for c, h in zip("ABCDEF", ["Rating", "A", "B", "C", "D (Ausfall)", "Summe"]):
        cells[f"{c}17"] = h
    cells["A18"] = "Anteil Basisjahr"
    for c, v in zip("BCDE", [0.30, 0.45, 0.20, 0.05]):
        cells[f"{c}18"] = v
    cells["F18"] = ("=", "SUM(B18:E18)")
    cells["A19"] = "Portfolioverteilung nach Jahren (Potenzen der Übergangsmatrix)"
    for c, h in zip("ABCDEF", ["Jahr", "A", "B", "C", "D (Ausfall)", "Summe"]):
        cells[f"{c}20"] = h
    cells["A21"] = 0.0
    for c in "BCDE":
        cells[f"{c}21"] = ("=", f"{c}18")
    cells["F21"] = ("=", "SUM(B21:E21)")
    for y in range(1, 6):
        r = 21 + y
        cells[f"A{r}"] = float(y)
        for c in "BCDE":
            cells[f"{c}{r}"] = ("=", f"MMULT($B{r - 1}:$E{r - 1},{c}$11:{c}$14)")
        cells[f"F{r}"] = ("=", f"SUM(B{r}:E{r})")
    cells["A27"] = "Erwarteter Verlust"
    for c, h in zip("ABCD", ["Jahr", "Neu-Ausfallanteil", "Erwarteter Verlust (TEUR)", "Barwert (TEUR)"]):
        cells[f"{c}28"] = h
    for y in range(1, 6):
        r = 28 + y
        cells[f"A{r}"] = float(y)
        cells[f"B{r}"] = ("=", f"E{21 + y}-E{20 + y}")
        cells[f"C{r}"] = ("=", f"B{r}*$B$5*$B$6")
        cells[f"D{r}"] = ("=", f"C{r}/(1+$B$7)^A{r}")
    cells["A34"] = "Summe"
    for c in "BCD":
        cells[f"{c}34"] = ("=", f"SUM({c}29:{c}33)")
    cells["A36"] = "Ergebnisübersicht"
    for r, name, f in ((37, "Kumulierte Ausfallquote nach 5 Jahren", "E26"), (38, "Erwarteter Verlust gesamt (TEUR)", "C34"),
                       (39, "Barwert Verlust (TEUR)", "D34"), (40, "Verlustquote am Portfolio", "B38/B5"),
                       (41, "Prüfung: Verteilung Jahr 5 = 100 %", 'IF(ABS(F26-1)<0.0001,"OK","Fehler")')):
        cells[f"A{r}"], cells[f"B{r}"] = name, ("=", f)
    for address, value in changes.items():
        cells[address] = ("=", value[1:]) if isinstance(value, str) and value.startswith("=") else value
    return _computed(cells)


def test_mmult_of_a_row_and_a_column_computes_like_excel():
    values = {(1, 1): 0.3, (1, 2): 0.7, (2, 3): 0.5, (3, 3): 0.25}
    calc = formel.Evaluator(values.get)
    assert abs(calc.run(formel.parse("MMULT(A1:B1,C2:C3)")) - 0.325) < 1e-12


def test_tables_stacked_on_one_sheet_are_found_one_by_one():
    cells = {formel.position(a): Cell(str(v[2] if isinstance(v, tuple) else v), v[1] if isinstance(v, tuple) else None)
             for a, v in _markov().items()}
    sheet = Sheet("Tabelle3", cells)
    plan = layout(sheet)
    found = [(p.header_row, p.last_row) for p in [plan, *others(sheet, plan)]]
    assert (20, 26) in found and (10, 14) in found and (28, 34) in found
    # Die Kopfzeile der nächsten Tabelle („Jahr | Neu-Ausfallanteil …“) gehört nicht zur Verteilung darüber.
    assert all(not (top <= 28 <= bottom) for top, bottom in found if top != 28)


def test_a_clean_markov_model_has_nothing_to_say():
    assert _hints(_markov(), "Tabelle3") == {}


def test_the_four_planted_errors_are_found_with_their_cause():
    """Testmappe, Tabelle3: C13 getippt 0,12 statt 0,10, C23 multipliziert mit der Spalte B, B32 nimmt E23 statt E24,
    D29:D33 zinsen mit der LGD statt mit dem Diskontsatz ab."""
    found = _hints(_markov(C13=0.12, C23="=MMULT($B22:$E22,B$11:B$14)", B32="=E25-E23",
                           **{f"D{r}": f"=C{r}/(1+$B$6)^A{r}" for r in range(29, 34)}), "Tabelle3")
    assert set(found) == {"pattern:C23", "pattern:B32", "hundred:F13", "discount:D29"}
    assert all(h["level"] == "fehler" for h in found.values())
    assert found["pattern:C23"]["fix"]["german"] == "=MMULT($B22:$E22;C$11:C$14)"
    assert found["pattern:B32"]["fix"]["german"] == "=E25-E24"
    hundred = found["hundred:F13"]["text"]
    assert "Zeile „C“ ergibt in F13 102 % statt 100 %" in hundred and "2 Prozentpunkte zu viel" in hundred
    assert "„A“, „B“ und „D (Ausfall)“ ergeben genau 100 %" in hundred
    discount = found["discount:D29"]
    assert "zinsen mit B6 (Verlustquote bei Ausfall (LGD), 45 %) ab statt mit B7 (Diskontsatz p.a., 5 %)" in discount["text"]
    assert discount["fix"]["german"] == "=C29/(1+$B$7)^A29" and set(discount["fix"]["cells"]) == {f"D{r}" for r in range(29, 34)}
    # Die eigene Prüfung in B41 zeigt „Fehler“: Das sagen die beiden Fehler, von denen es kommt.
    assert "Prüfung des Modells in B41" in found["pattern:C23"]["text"] and "B41" in hundred


def test_a_sum_next_to_a_year_counter_does_not_miss_the_year():
    """=SUMME(B21:E21) neben „Jahr 0, 1, 2 …“ in Spalte A lässt kein Jahr aus: A ist der Name der Zeile."""
    assert not any(k.startswith("parts:") for k in _hints(_markov(), "Tabelle3"))


def test_a_failing_model_check_without_a_known_cause_is_reported():
    found = _hints(_markov(B18=0.32), "Tabelle3")
    assert found["selfcheck:B41"]["level"] == "fehler"
    assert "zeigt „Fehler“" in found["selfcheck:B41"]["text"] and "F26" in found["selfcheck:B41"]["text"]


def test_the_pane_view_splits_title_place_and_the_two_formulas():
    hint = {"key": "pattern:B32", "cell": "B32", "level": "fehler",
            "text": "B32 (Neu-Ausfallanteil, Jahr 4) rechnet anders als die übrigen 4 Formeln in B29:B33: "
                    "=E25 - E23 statt =E25 - E24. B32 zeigt darum 0,0213 statt 0,022989."}
    view = ansicht.view(hint)
    assert view["title"] == "Formel weicht vom Muster ab" and view["where"] == "Neu-Ausfallanteil · Jahr 4"
    assert view["lead"] == "Rechnet anders als die übrigen 4 Formeln in B29:B33."
    compare = view["compare"]
    assert (compare["now"], compare["want"]) == ("=E25-E23", "=E25-E24")
    assert compare["values"]["now"] == "0,0213" and compare["values"]["want"] == "0,022989"
    # Mit Namen vorgelesene Formeln behalten ihre Leerzeichen.
    named = ansicht.view({"key": "pattern:C37", "cell": "C37", "level": "fehler",
                          "text": "C37 (end 13,5) rechnet anders als die übrigen 13 Formeln in C24:C37: "
                                  "=end 12,5 + 0,5 statt =end 12,5 + 1."})
    assert named["compare"]["now"] == "=end 12,5 + 0,5"
