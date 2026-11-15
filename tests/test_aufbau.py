"""Aufbauten, die in sauberen Mappen falschen Alarm gaben: ein zweizeiliger Kopf (verbunden), zwei Tabellen
nebeneinander, Annahmen rechts neben der Tabelle, eine leere Trennzeile und Berichte mit Zwischensummen. Dazu die
Fehler, die gerade in solchen Berichten typisch sind: Die Gesamtsumme lässt eine Gruppe aus oder zählt alles doppelt."""

from test_workbook import xlsx

from voiceagent import formel, pruefer

STYLES = '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs>'
REGIONS = ["Nord", "Süd", "Ost", "West"]


def _hints(cells: dict, name: str = "Daten") -> dict:
    done = pruefer.Checker(None).check_file("mappe.xlsx", xlsx((name, cells), styles=STYLES))
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


def _two_headers(**changes) -> dict:
    """„Umsatz“ über 2025 und 2026 (in Excel verbunden: nur die erste Zelle hat den Text), ebenso Kosten und Gewinn."""
    cells = {"A1": "Filiale", "B1": "Umsatz", "D1": "Kosten", "F1": "Gewinn"}
    for col, year in zip("BCDEFG", ["2025", "2026"] * 3):
        cells[f"{col}2"] = float(year)
    names = ["Basel", "Bern", "Chur", "Genf", "Luzern", "Lugano", "Sion", "Thun", "Zug", "Zürich", "Aarau", "Biel"]
    values = {}
    for i, n in enumerate(names):
        r = i + 3
        b, c, d, e = 1000 + 37 * i, 1100 + 41 * i, 600 + 13 * i, 640 + 17 * i
        values[r] = (b, c, d, e)
        cells.update({f"A{r}": n, f"B{r}": b, f"C{r}": c, f"D{r}": d, f"E{r}": e,
                      f"F{r}": ("=", f"B{r}-D{r}", b - d), f"G{r}": ("=", f"C{r}-E{r}", c - e)})
    for k, col in enumerate("BCDE"):
        cells[f"{col}15"] = ("=", f"SUM({col}3:{col}14)", sum(v[k] for v in values.values()))
    cells["A15"] = "Total"
    cells["F15"] = ("=", "SUM(F3:F14)", sum(v[0] - v[2] for v in values.values()))
    cells["G15"] = ("=", "SUM(G3:G14)", sum(v[1] - v[3] for v in values.values()))
    cells.update(changes)
    return cells


def test_a_two_row_header_is_one_header_and_the_columns_get_both_names():
    assert _hints(_two_headers()) == {}
    found = _hints(_two_headers(G9=("=", "C9-D9", 1346 - 678)))
    hint = found["pattern:G9"]
    assert "=Umsatz 2026 - Kosten 2025 statt =Umsatz 2026 - Kosten 2026" in hint["text"]
    assert hint["fix"]["cells"] == {"G9": "=C9-E9"}


def _side_by_side(**changes) -> dict:
    cells = {"A1": "Monat", "B1": "Umsatz", "C1": "Kosten", "D1": "Gewinn",
             "F1": "Filiale", "G1": "Mitarbeitende", "H1": "Umsatz", "I1": "Umsatz je MA"}
    months = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November",
              "Dezember"]
    for i, m in enumerate(months):
        r = i + 2
        b, c = 5000 + 211 * i, 3000 + 97 * i
        cells.update({f"A{r}": m, f"B{r}": b, f"C{r}": c, f"D{r}": ("=", f"B{r}-C{r}", b - c)})
    for i, (town, staff) in enumerate(zip(["Basel", "Bern", "Chur", "Genf", "Luzern", "Zug", "Zürich", "Thun"],
                                          [4, 5, 6, 7, 8, 4, 5, 6])):
        r = i + 2
        cells.update({f"F{r}": town, f"G{r}": staff, f"H{r}": 20000 + 1500 * i,
                      f"I{r}": ("=", f"H{r}/G{r}", (20000 + 1500 * i) / staff)})
    cells.update(changes)
    return cells


def test_two_tables_side_by_side_are_checked_each_for_itself():
    """Die rechte Tabelle ist kürzer: Ihre leeren Zeilen sind keine fehlenden Werte, und ihre Zeilen heißen nach
    der Filiale, nicht nach dem Monat links."""
    assert _hints(_side_by_side()) == {}
    found = _hints(_side_by_side(I6=("=", "H6/G5", 26000 / 7)))
    hint = found["pattern:I6"]
    assert "I6 (Umsatz je MA Luzern)" in hint["text"] and hint["fix"]["cells"] == {"I6": "=H6/G6"}


def test_assumptions_next_to_the_table_are_no_column_of_it():
    """„MwSt | 0,081“ und „Rabatt | 0,05“ zwei Spalten rechts neben der Tabelle: Annahmen, keine Spalte „MwSt“."""
    cells = {"A1": "Artikel", "B1": "Netto", "C1": "Rabatt", "D1": "MwSt-Betrag", "E1": "Brutto",
             "H1": "MwSt", "I1": 0.081, "H2": "Rabatt", "I2": 0.05}
    for i in range(14):
        r = i + 2
        net = 100 + 13 * i
        off = net * 0.05
        tax = (net - off) * 0.081
        cells.update({f"A{r}": f"Artikel {i + 1}", f"B{r}": net, f"C{r}": ("=", f"B{r}*$I$2", off),
                      f"D{r}": ("=", f"(B{r}-C{r})*$I$1", tax), f"E{r}": ("=", f"B{r}-C{r}+D{r}", net - off + tax)})
    assert _hints(cells) == {}


def test_an_empty_row_between_two_parts_of_a_table_is_no_missing_value():
    cells = {"A1": "Datum", "B1": "Region", "C1": "Menge", "D1": "Preis", "E1": "Umsatz"}
    total, r = 0.0, 2
    for i in range(16):
        if i == 8:
            r += 1
        q, p = 10 + (i * 7) % 23, round(12.5 + (i * 3.1) % 9, 2)
        total += q * p
        cells.update({f"A{r}": (46027 + 7 * i, 1), f"B{r}": REGIONS[i % 4], f"C{r}": q, f"D{r}": p,
                      f"E{r}": ("=", f"C{r}*D{r}", q * p)})
        r += 1
    cells.update({f"D{r}": "Summe", f"E{r}": ("=", f"SUM(E2:E{r - 1})", total)})
    assert _hints(cells) == {}


def _report(subtotal: str = "SUM", total: str | None = None, **changes) -> dict:
    """Region | Produkt | Umsatz, je Region drei Zeilen und eine Zwischensumme, ganz unten Gesamt."""
    cells = {"A1": "Region", "B1": "Produkt", "C1": "Umsatz"}
    r, subs, values = 2, [], {}
    for reg in REGIONS:
        first = r
        for prod in ["Alpha", "Beta", "Gamma"]:
            values[r] = 200 + 37 * r % 300
            cells.update({f"A{r}": reg, f"B{r}": prod, f"C{r}": values[r]})
            r += 1
        amount = sum(values[x] for x in range(first, r))
        inner = f"C{first}:C{r - 1}"
        cells.update({f"A{r}": f"Summe {reg}",
                      f"C{r}": ("=", f"SUM({inner})" if subtotal == "SUM" else f"SUBTOTAL(9,{inner})", amount)})
        values[r] = amount
        subs.append(r)
        r += 1
    formula = total or "+".join(f"C{s}" for s in subs)
    shown = sum(values[s] for s in subs)
    if total == "SUM(C2:C17)":
        shown *= 2
    if total == "C5+C9+C13":
        shown = sum(values[s] for s in subs[:3])
    cells.update({f"A{r}": "Gesamt", f"C{r}": ("=", formula, shown)})
    cells.update(changes)
    return cells


def test_a_report_with_subtotals_has_nothing_to_say():
    """Jede Zwischensumme zählt nur ihre Gruppe: Das ist keine Summe, die Zeilen auslässt."""
    assert _hints(_report()) == {}
    # So baut Excel es selbst (Daten → Teilergebnis): TEILERGEBNIS über alles lässt die inneren weg.
    assert _hints(_report("SUBTOTAL", "SUBTOTAL(9,C2:C17)")) == {}


def test_a_grand_total_that_leaves_out_a_group():
    hint = _hints(_report(total="C5+C9+C13"))["subtotal:C18"]
    assert hint["level"] == "fehler"
    assert "Gesamt (C18) zählt nur 3 der 4 Zwischensummen zusammen: Summe West (C17) fehlt." in hint["text"]
    assert hint["fix"]["cells"] == {"C18": "=C5+C9+C13+C17"}


def test_a_grand_total_that_counts_every_value_twice():
    found = _hints(_report(total="SUM(C2:C17)"))
    hint = found["subtotal:C18"]
    assert "summiert C2:C17 und zählt damit die Zeilen und ihre 4 Zwischensummen" in hint["text"]
    assert "Jeder Wert zählt doppelt" in hint["text"]
    assert hint["fix"]["cells"] == {"C18": "=SUM(C5,C9,C13,C17)"}
    assert not [k for k in found if k.startswith("double:")]


def test_subtotal_leaves_out_other_subtotals_in_its_range():
    values = {(1, 1): 5.0, (2, 1): 7.0, (3, 1): 12.0, (4, 1): 3.0, (5, 1): 15.0, (6, 1): "x"}
    calc = formel.Evaluator(lambda pos: values.get(pos))
    calc.nested = lambda pos: pos in {(3, 1), (5, 1)}
    assert calc.run(formel.parse("SUBTOTAL(9,A1:A5)")) == 15.0
    assert calc.run(formel.parse("SUBTOTAL(109,A1:A5)")) == 15.0
    assert calc.run(formel.parse("SUBTOTAL(1,A1:A4)")) == 5.0
    assert calc.run(formel.parse("SUBTOTAL(3,A1:A6)")) == 4.0
    assert calc.run(formel.parse("SUBTOTAL(4,A1:A4)")) == 7.0
    assert isinstance(calc.run(formel.parse("SUBTOTAL(12,A1:A4)")), formel.Err)


def test_a_repeated_name_is_only_asked_about_when_it_looks_like_the_same_entry():
    """„Food“ mit zwei Konten ist eine Kategorie, 700 ist kein Name, und „Pork“ unter der Tabelle mit SVERWEIS
    daneben ist eine Abfrage. Ein Name zweimal mit sonst gleichen Angaben bleibt eine Frage."""
    cells = {"A1": "Category", "B1": "Account", "C1": "Opening", "D1": "Ending"}
    for r, (cat, acc, a, b) in enumerate([("Food", "City Bank", 1000, 800), ("Food", "J.P Morgan", 500, 400),
                                          ("Gas", "J.P Morgan", 400, 200), ("Clothes", "City Bank", 800, 700),
                                          ("Payroll", "City Bank", 700, 1700), ("Business", "J.P Morgan", 700, 1700)], 2):
        cells.update({f"A{r}": cat, f"B{r}": acc, f"C{r}": a, f"D{r}": b})
    assert not [k for k in _hints(cells) if k.startswith("dupe:")]
    lookup = {"A1": "Meat", "B1": "Amount", "A2": "Beef", "B2": 50, "A3": "Chicken", "B3": 30, "A4": "Pork", "B4": 10,
              "A5": "Fish", "B5": 50, "A7": "Pork", "B7": ("=", "VLOOKUP(A7,A2:B5,2,FALSE)", 10)}
    assert not [k for k in _hints(lookup) if k.startswith("dupe:")]
    names = {"A1": "Name", "B1": "Ort", "C1": "Betrag"}
    for r, (n, o, b) in enumerate([("Meier", "Bern", 120), ("Huber", "Basel", 80), ("Keller", "Chur", 95),
                                   ("Meier", "Bern", 120), ("Graf", "Zug", 60), ("Frei", "Thun", 70)], 2):
        names.update({f"A{r}": n, f"B{r}": o, f"C{r}": b})
    assert "dupe:A5" in _hints(names)
