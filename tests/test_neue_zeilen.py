"""Neue Fälle unten angehängt: Formeln, die in den neuen Zeilen fehlen, und feste Bereiche ($A$2:$A$21), die vor
ihnen aufhören. Je ein Hinweis für eine ganze Spalte, mit dem Knopf, der sie auf einmal richtig macht."""

from test_haertetest import _hints

from voiceagent import pivotdiagramm
from voiceagent.xlsx import Cell, Pivot, PivotValue, Sheet

STATIONS = ["K1", "C1", "M1"]


def cases(n: int, new: int = 0, filled: bool = False, end: int | None = None) -> dict:
    """n Fälle mit festen Bereichen bis Zeile end (sonst n + 1) in D (Anzahl der Fall-Nr) und E (Kosten der Station),
    darunter new neue Fälle, mit den Formeln (filled) oder ohne."""
    rows = [(f"F{100 + i}", STATIONS[i % 3], 100.0 + 10 * i) for i in range(n + new)]
    end = end or n + 1
    cells = {"A1": "Fall-Nr", "B1": "Station", "C1": "Kosten", "D1": "Anzahl Nr", "E1": "Kosten Station"}
    counted = rows[:end - 1]
    for i, (case, station, cost) in enumerate(rows):
        r = i + 2
        cells[f"A{r}"], cells[f"B{r}"], cells[f"C{r}"] = case, station, cost
        if i < n or filled:
            cells[f"D{r}"] = ("=", f"COUNTIF($A$2:$A${end},A{r})", sum(1 for c, _, _ in counted if c == case))
            cells[f"E{r}"] = ("=", f"SUMIF($B$2:$B${end},B{r},$C$2:$C${end})",
                              sum(cost for _, s, cost in counted if s == station))
    return cells


def on(found: dict, sheet: str = "Fälle") -> dict:
    return {k: h for (s, k), h in found.items() if s == sheet}


def test_a_clean_table_is_quiet():
    assert not [k for k in on(_hints(("Fälle", cases(30)))) if k.split(":")[0] in ("unfilled", "partial", "criterion")]


def test_new_rows_without_their_formulas():
    here = on(_hints(("Fälle", cases(30, new=2))))
    hint = here["unfilled:D32"]
    assert "In den Zeilen 32 bis 33 fehlen die Formeln von Anzahl Nr und Kosten Station (D bis E)" in hint["text"]
    assert "Darüber stehen sie in jeder Zeile bis 31" in hint["text"]
    assert hint["fix"]["range"] == "D32:E33"
    assert hint["fix"]["cells"] == {"D32": "=COUNTIF($A$2:$A$31,A32)", "D33": "=COUNTIF($A$2:$A$31,A33)",
                                    "E32": "=SUMIF($B$2:$B$31,B32,$C$2:$C$31)", "E33": "=SUMIF($B$2:$B$31,B33,$C$2:$C$31)"}
    # Die leeren Zellen sagt der Hinweis schon: nicht noch einmal unter „Fehlende Werte“.
    assert not [k for k in here if k.startswith("missing:")]


def test_a_formula_that_reads_the_next_row_is_missing_in_the_last_row_on_purpose():
    cells = {"A1": "Monat", "B1": "Patienten", "C1": "Zunahme"}
    for i in range(20):
        cells[f"A{i + 2}"], cells[f"B{i + 2}"] = f"M{i + 1}", 100.0 + 3 * i
        if i < 19:
            cells[f"C{i + 2}"] = ("=", f"B{i + 3}-B{i + 2}", 3.0)
    assert not [k for k in on(_hints(("Fälle", cells))) if k.startswith("unfilled:")]


def test_fixed_ranges_that_stop_before_the_new_rows_are_one_hint_per_column():
    here = on(_hints(("Fälle", cases(30, new=1, filled=True))))
    partial = {k: h for k, h in here.items() if k.startswith("partial:")}
    assert len(partial) == 2, partial
    sums = next(h for h in partial.values() if "SUMMEWENN" in h["text"])
    assert "SUMMEWENN in E2 bis E32 (Kosten Station) prüft in jeder Zeile nur B2:B31, die Tabelle geht aber bis Zeile 32" in sums["text"]
    assert len(sums["fix"]["cells"]) == 31 and sums["fix"]["cells"]["E2"] == "=SUMIF($B$2:$B$32,B2,$C$2:$C$32)"
    counts = here["partial:D32"]
    assert counts["fix"]["cells"]["D32"] == "=COUNTIF($A$2:$A$32,A32)" and len(counts["fix"]["cells"]) == 31
    # Nicht „Gemeint ist wohl F129“: Die Fall-Nr steht ja in A32, nur nicht im Bereich.
    assert not [k for k in here if k.startswith("criterion:")]


def test_ranges_on_another_sheet_one_hint_per_column_for_all_ranges_of_a_formula():
    report = {"A1": "Station", "B1": "Fälle", "C1": "Kosten"}
    for r, station in enumerate(STATIONS, start=2):
        end = 29 if r == 3 else 31
        report[f"A{r}"] = station
        report[f"B{r}"] = ("=", f"COUNTIF(Fälle!$B$2:$B${end},A{r})", 0)
        report[f"C{r}"] = ("=", f"SUMIF(Fälle!$B$2:$B$31,A{r},Fälle!$C$2:$C$31)", 0)
    report["A5"] = "Neurologie"
    report["B5"] = ("=", "COUNTIF(Fälle!$B$2:$B$31,A5)", 0)
    here = on(_hints(("Fälle", cases(30, new=2, filled=True, end=31)), ("Auswertung", report)), "Auswertung")
    ranges = {k: h for k, h in here.items() if k.startswith("range:")}
    # B2, B4 und B5 zusammen (B3 hört noch früher auf), C2 bis C4 mit beiden Bereichen in einem Hinweis.
    assert set(ranges) == {"range:B2", "range:B3", "range:C2"}, ranges
    assert "und 2 gleichen Formeln darunter nehmen Fälle!B2:B31" in ranges["range:B2"]["text"]
    assert set(ranges["range:B2"]["fix"]["cells"]) == {"B2", "B4", "B5"}
    assert ranges["range:B3"]["fix"]["cells"] == {"B3": "=COUNTIF(Fälle!$B$2:$B$33,A3)"}
    sums = ranges["range:C2"]
    assert "nehmen Fälle!B2:B31 und Fälle!C2:C31, auf „Fälle“ gehen Station und Kosten aber bis Zeile 33" in sums["text"]
    assert sums["fix"]["cells"]["C4"] == "=SUMIF(Fälle!$B$2:$B$33,A4,Fälle!$C$2:$C$33)"


def test_a_pivot_source_with_a_running_sum_beside_the_data():
    """Eine Laufsumme in der neuen Zeile (=SUMME($C$2:C12)) ist keine Summenzeile: Die Zeile fehlt der Pivot-Tabelle."""
    cells = {(1, 1): Cell("Station"), (1, 2): Cell("Fälle"), (1, 3): Cell("Kosten"), (1, 4): Cell("Laufsumme")}
    total = 0.0
    for r in range(2, 13):
        total += 10.0 * r
        cells.update({(r, 1): Cell(STATIONS[r % 3]), (r, 2): Cell("1"), (r, 3): Cell(str(10 * r)),
                      (r, 4): Cell(str(total), f"SUM($C$2:C{r})")})
    described = Pivot("PivotTable1", (1, 7, 5, 8), ("Station",), (), ("Summe von Kosten",),
                      value_fields=(PivotValue("Summe von Kosten", "Kosten", "sum"),), body=(2, 8, 5, 8),
                      source_sheet="Daten", source_area=(1, 1, 11, 4))
    sheet = Sheet("Daten", cells, pivots=[described])
    found = {f.key.split(":")[0]: f for f in pivotdiagramm.checks("Daten", sheet, {"Daten": cells})}
    assert "Zeile 12 fehlt in der Pivot-Tabelle" in found["pivotsource"].text
    # Eine echte Summenzeile darunter bleibt still.
    cells[(12, 1)], cells[(12, 3)], cells[(12, 4)] = Cell("Summe"), Cell("770", "SUM(C2:C11)"), Cell("")
    cells[(12, 2)] = Cell("10", "SUM(B2:B11)")
    found = {f.key.split(":")[0]: f for f in pivotdiagramm.checks("Daten", Sheet("Daten", cells, pivots=[described]),
                                                                  {"Daten": cells})}
    assert "pivotsource" not in found


def test_a_longer_lookup_table_beside_the_cases_is_no_new_rows():
    """Eine Nachschlagetabelle neben den Fällen ist länger als sie: Ihre Zeilen darunter sind keine neuen Fälle, auch
    wenn sie ohne leere Spalte dazwischen anschließt."""
    for left in (6, 4):
        code, price = column(left), column(left + 1)
        cells = {"A1": "Code", "B1": "Menge", "C1": "Preis", f"{code}1": "Code", f"{price}1": "Preis"}
        for i in range(15):
            cells[f"{code}{i + 2}"], cells[f"{price}{i + 2}"] = f"P{i}", 10.0 + i
        for i in range(12):
            r = i + 2
            cells[f"A{r}"], cells[f"B{r}"] = f"P{(3 * i) % 15}", 1.0 + i % 4
            cells[f"C{r}"] = ("=", f"VLOOKUP(A{r},${code}$2:${price}$16,2,FALSE)", 10.0 + (3 * i) % 15)
        found = on(_hints(("Fälle", cells)))
        assert not [k for k in found if k.startswith("unfilled:")], (left, found)


def column(n: int) -> str:
    return chr(ord("A") + n - 1)
