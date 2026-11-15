"""Pivot-Tabellen und Diagramme: aktuell und richtig still, alt, zu kurz oder verschoben erkannt, mit dem, was jetzt
herauskäme, und im Seitenfenster mit einem Knopf."""

import math
from statistics import mean, pstdev, stdev

import pytest
from fastapi.testclient import TestClient
from test_excel_structure import MAIN, _rels, mappe

from voiceagent import pivotdiagramm, pruefer
from voiceagent.addin import sheets_from
from voiceagent.app import create_app
from voiceagent.config import Settings
from voiceagent.xlsx import Cell, Chart, ExcelTable, Pivot, PivotValue, Series, Sheet, area_text, column_letters, number_text, read_book

ROWS = [("Nord", "Maus", 120.5), ("Ost", "Maus", 80.0), ("Nord", "Monitor", 310.0), ("Süd", "Tastatur", 45.25),
        ("Ost", "Monitor", 260.0), ("Nord", "Maus", 99.0), ("Süd", "Maus", 30.0), ("Ost", "Tastatur", 15.0)]
HEADS = ("Region", "Produkt", "Umsatz")
FUNCTIONS = {"sum": math.fsum, "count": len, "average": mean, "max": max, "min": min, "stdDev": stdev,
             "stdDevp": pstdev, "countNums": len}
CAPTIONS = {"sum": "Summe von", "count": "Anzahl von", "average": "Mittelwert von", "max": "Maximum von",
            "min": "Minimum von", "stdDev": "StdAbw von", "stdDevp": "StdAbwN von", "countNums": "Anzahl2 von"}


def text(value) -> str:
    return number_text(value) if isinstance(value, float) else str(value)


def data(rows=ROWS, heads=HEADS, top=1, left=1) -> dict:
    cells = {(top, left + j): Cell(h) for j, h in enumerate(heads)}
    for i, row in enumerate(rows, start=top + 1):
        for j, value in enumerate(row):
            if value is not None and value != "":
                cells[(i, left + j)] = Cell(text(value))
    return cells


def pivot(rows=ROWS, by=0, across=None, field=2, functions=("sum",), at=(1, 6), skipped=(), grand=(True, True),
          heads=HEADS, source=None, name="PivotTable1", filtered=False):
    """Eine Pivot-Tabelle so, wie Excel sie aus rows zeigt: Einträge sortiert, Gesamtergebnisse, leere Kreuzungen leer.
    Gibt ihre Zellen und ihre Beschreibung."""
    hidden = {heads[f]: {v.casefold() for v in items} for f, items in skipped}
    kept = [r for r in rows if not any(str(r[heads.index(f)]).casefold() in items for f, items in hidden.items())]
    row_items = sorted({r[by] for r in kept}, key=lambda v: str(v).casefold())
    col_items = sorted({r[across] for r in kept}, key=lambda v: str(v).casefold()) if across is not None else []
    top, left = at
    cells = {}
    values = [PivotValue(f"{CAPTIONS[f]} {heads[field]}", heads[field], f) for f in functions]

    def number(f, group):
        nums = [r[field] for r in group if isinstance(r[field], float)]
        if f in ("count",):
            return float(len([r for r in group if r[field] not in (None, "")]))
        if f == "countNums":
            return float(len(nums))
        return float(FUNCTIONS[f](nums)) if nums and not (f == "stdDev" and len(nums) < 2) else None

    if across is None:
        cells[(top, left)] = Cell("Zeilenbeschriftungen")
        for j, v in enumerate(values):
            cells[(top, left + 1 + j)] = Cell(v.caption)
        body_top = top + 1
        lines = [(item, [r for r in kept if r[by] == item]) for item in row_items]
        if grand[0]:
            lines.append(("Gesamtergebnis", kept))
        for i, (label, group) in enumerate(lines, start=body_top):
            cells[(i, left)] = Cell(text(label))
            for j, f in enumerate(functions):
                value = number(f, group)
                if value is not None:
                    cells[(i, left + 1 + j)] = Cell(text(value))
        bottom, right = body_top + len(lines) - 1, left + len(functions)
    else:
        f = functions[0]
        cells[(top, left)] = Cell(values[0].caption)
        cells[(top, left + 1)] = Cell("Spaltenbeschriftungen")
        cells[(top + 1, left)] = Cell("Zeilenbeschriftungen")
        heads_across = col_items + (["Gesamtergebnis"] if grand[1] else [])
        for j, item in enumerate(heads_across):
            cells[(top + 1, left + 1 + j)] = Cell(text(item))
        body_top = top + 2
        lines = [(item, [r for r in kept if r[by] == item]) for item in row_items]
        if grand[0]:
            lines.append(("Gesamtergebnis", kept))
        for i, (label, group) in enumerate(lines, start=body_top):
            cells[(i, left)] = Cell(text(label))
            for j, item in enumerate(heads_across):
                part = group if item == "Gesamtergebnis" and grand[1] and j == len(heads_across) - 1 else \
                    [r for r in group if r[across] == item]
                value = number(f, part) if part else None
                if value is not None:
                    cells[(i, left + 1 + j)] = Cell(text(value))
        bottom, right = body_top + len(lines) - 1, left + len(heads_across)
    area = (top, left, bottom, right)
    body = (body_top, left + 1, bottom, right)
    src = source or (1, 1, len(rows) + 1, len(heads))
    described = Pivot(name, area, (heads[by],), (heads[across],) if across is not None else (), tuple(v.caption for v in values),
                      value_fields=tuple(values), body=body, source_sheet="Daten", source_area=src,
                      skipped=tuple((heads[f], tuple(items)) for f, items in skipped), grand=grand, filtered=filtered)
    return cells, described


def hints(cells, pivots=(), charts=(), others=None, tables=None, live=True):
    sheet = Sheet("Daten", cells, pivots=list(pivots), charts=list(charts))
    return {f.key.split(":")[0]: f for f in pivotdiagramm.checks("Daten", sheet, {"Daten": cells, **(others or {})}, tables, live)}


def fresh(**kw):
    rows = kw.pop("rows", ROWS)
    cells = data(rows)
    block, described = pivot(rows, **kw)
    return {**cells, **block}, described


@pytest.mark.parametrize("kw", [
    {}, {"functions": ("sum", "count", "average")}, {"functions": ("max", "min", "stdDev", "stdDevp")}, {"across": 1},
    {"across": 1, "grand": (True, False)}, {"grand": (False, True)}, {"by": 1}, {"skipped": ((0, ("Süd",)),)},
    {"skipped": ((1, ("Maus",)),)}, {"across": 1, "skipped": ((1, ("Tastatur",)),)},
])
def test_a_fresh_pivot_is_quiet(kw):
    cells, described = fresh(**kw)
    assert hints(cells, [described]) == {}


def test_a_pivot_that_was_not_refreshed():
    cells, described = fresh()
    # Danach geändert, ohne zu aktualisieren: Ost/Maus 80 → 95.
    cells[(3, 3)] = Cell("95")
    hint = hints(cells, [described])["pivot"]
    assert hint.level == "fehler" and hint.key == "pivot:G3"
    assert "G3 „Ost“ zeigt 355, die Daten ergeben jetzt 370" in hint.text
    assert "das Gesamtergebnis zeigt 959,75, die Daten ergeben jetzt 974,75" in hint.text
    assert "nicht aktualisiert" in hint.text and "Aktualisieren" in hint.tip
    assert hint.action == {"kind": "refresh", "pivot": "PivotTable1"}
    # Aus einer Datei: kein Knopf.
    assert hints(cells, [described], live=False)["pivot"].action is None


def test_new_and_renamed_entries_since_the_last_refresh():
    rows = ROWS + [("West", "Maus", 70.0)]
    cells, described = fresh()
    cells.update(data(rows))
    hint = hints(cells, [replace_source(described, len(rows) + 1)])["pivot"]
    assert "„West“ fehlt ganz, obwohl 1 Zeilen der Daten dazugehören" in hint.text
    cells, described = fresh()
    cells[(5, 1)] = Cell("Süden")  # Süd → Süden, Zeile 5
    cells[(8, 1)] = Cell("Süden")
    hint = hints(cells, [described])["pivot"]
    assert "„Süd“ kommt in den Daten nicht mehr vor" in hint.text and "„Süden“ fehlt ganz" in hint.text


def replace_source(described: Pivot, bottom: int) -> Pivot:
    from dataclasses import replace
    return replace(described, source_area=(1, 1, bottom, 3))


def test_cross_table_with_a_column_field():
    cells, described = fresh(across=1)
    cells[(9, 3)] = Cell("25")  # Ost/Tastatur 15 → 25
    hint = hints(cells, [described])["pivot"]
    assert "„Ost“ / „Tastatur“ zeigt 15, die Daten ergeben jetzt 25" in hint.text
    assert "das Gesamtergebnis für „Tastatur“" in hint.text and "„Ost“ im Gesamtergebnis" in hint.text


def test_unknown_filters_count_only_what_is_shown():
    # Ein Wertfilter (Top 2) blendet „Süd“ aus: Weiß man nicht, ob gefiltert ist, zählt das Gesamtergebnis nur die
    # gezeigten Einträge und fehlende Einträge sind kein Fehler.
    cells = data()
    block, described = pivot([r for r in ROWS if r[0] != "Süd"])
    from dataclasses import replace
    described = replace(described, source_area=(1, 1, len(ROWS) + 1, 3), filtered=None)
    assert hints({**cells, **block}, [described]) == {}
    # Weiß man, dass nichts gefiltert ist, fehlt „Süd“.
    assert "„Süd“ fehlt ganz" in hints({**cells, **block}, [replace(described, filtered=False)])["pivot"].text


def test_hidden_items_that_do_not_match_the_data_are_not_guessed():
    cells, described = fresh(skipped=((0, ("Süd",)),))
    from dataclasses import replace
    odd = replace(described, skipped=(("Region", ("Sued",)),))
    cells[(3, 3)] = Cell("95")
    assert hints(cells, [odd]) == {}


def test_pivots_it_does_not_understand_stay_quiet():
    from dataclasses import replace
    cells, described = fresh()
    cells[(3, 3)] = Cell("95")
    for odd in (replace(described, plain=False), replace(described, rows=("Region", "Produkt")),
                replace(described, rows=("Gebiet",)), replace(described, value_fields=(PivotValue("Summe von Umsatz", "Umsatz", ""),)),
                replace(described, value_fields=(PivotValue("Summe von Umsatz", "Umsatz", "sum", True),)),
                replace(described, body=None)):
        assert "pivot" not in hints(cells, [odd])


def test_a_fixed_source_that_ends_before_the_data():
    rows = ROWS + [("Nord", "Maus", 40.0), ("West", "Monitor", 210.0)]
    cells, described = fresh()
    cells.update(data(rows))
    hint = hints(cells, [described])["pivotsource"]
    assert hint.level == "fehler" and hint.key == "pivotsource:A10"
    assert "fest aus A1:C9" in hint.text and "die 2 Zeilen 10 bis 11 fehlen" in hint.text and "auch nach dem Aktualisieren" in hint.text
    assert "Mit ihnen wäre „Nord“ 569,5 statt 529,5 und es käme „West“ (210) dazu." in hint.text
    assert "Datenquelle ändern → A1:C11" in hint.tip and "Befehl-T" in hint.tip


@pytest.mark.parametrize("below", [
    {(11, 1): Cell("Summe"), (11, 3): Cell("959,75", "SUM(C2:C9)")},
    {(10, 3): Cell("959,75", "SUM(C2:C9)")},
    {(11, 1): Cell("Nord"), (11, 3): Cell("5")},  # nach einer leeren Zeile: etwas anderes
    {(10, 1): Cell("Region"), (10, 2): Cell("Produkt"), (10, 3): Cell("Umsatz")},
])
def test_below_the_source_but_not_more_data(below):
    cells, described = fresh()
    assert "pivotsource" not in hints({**cells, **below}, [described])


def test_a_table_as_source_grows_by_itself():
    from dataclasses import replace
    rows = ROWS + [("Nord", "Maus", 40.0)]
    cells, described = fresh(rows=rows)
    table = ExcelTable("Verkauf", (1, 1, len(rows) + 1, 3), HEADS)
    described = replace(described, source_area=None, source_sheet="", source_table="Verkauf")
    assert hints(cells, [described], tables={"Verkauf": ("Daten", table)}) == {}
    cells[(3, 3)] = Cell("95")
    assert "der Tabelle „Verkauf“" in hints(cells, [described], tables={"Verkauf": ("Daten", table)})["pivot"].text


def test_count_instead_of_sum_because_of_a_blank_cell():
    rows = ROWS[:5] + [("Nord", "Maus", None)] + ROWS[6:]
    cells, described = fresh(rows=rows, functions=("count",))
    hint = hints(cells, [described])["pivotcount"]
    assert hint.level == "pruefen" and hint.key == "pivotcount:G2"
    assert "„Anzahl von Umsatz“ in der Pivot-Tabelle „PivotTable1“ zählt" in hint.text
    assert "(7), statt die Werte zu summieren (860,75)" in hint.text and "C7 ist leer" in hint.text
    assert hint.action == {"kind": "sum", "pivot": "PivotTable1", "field": "Anzahl von Umsatz"}
    # Text statt einer Zahl.
    rows = ROWS[:5] + [("Nord", "Maus", "k.A.")] + ROWS[6:]
    cells, described = fresh(rows=rows, functions=("count",))
    assert "in C7 steht „k.A.“" in hints(cells, [described])["pivotcount"].text


def test_count_meant_is_quiet():
    from dataclasses import replace
    # Lauter Zahlen: Excel hätte summiert, Anzahl ist also gewählt.
    cells, described = fresh(functions=("count",))
    assert hints(cells, [described]) == {}
    rows = ROWS[:5] + [("Nord", "Maus", None)] + ROWS[6:]
    cells, described = fresh(rows=rows, functions=("count",))
    # Selbst benannt („Fälle“) oder eine Spalte mit Nummern.
    assert hints(cells, [replace(described, value_fields=(PivotValue("Fälle", "Umsatz", "count"),))]) == {}
    ids = [(r[0], r[1], float(100 + i)) for i, r in enumerate(rows)]
    ids[5] = ("Nord", "Maus", None)
    cells, described = fresh(rows=ids, functions=("count",))
    assert "pivotcount" not in hints(cells, [described])


# --- Diagramme ---

MONTHS = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November",
          "Dezember"]


def months(total=True) -> dict:
    cells = {(1, 1): Cell("Monat"), (1, 2): Cell("Umsatz"), (1, 3): Cell("Kosten")}
    for r, m in enumerate(MONTHS, start=2):
        cells[(r, 1)], cells[(r, 2)], cells[(r, 3)] = Cell(m), Cell(str(1000 + 37 * r)), Cell(str(400 + 11 * r))
    if total:
        cells[(14, 1)] = Cell("Summe")
        cells[(14, 2)] = Cell("17106", "SUM(B2:B13)")
        cells[(14, 3)] = Cell("6534", "SUM(C2:C13)")
    return cells


def chart(*series, title="Umsatz und Kosten"):
    return Chart("Diagramm 1", "Säulendiagramm", title, tuple(Series(*s) for s in series))


def test_charts_over_the_whole_table_are_quiet():
    cells = months()
    assert hints(cells, charts=[chart(("Umsatz", "Daten!$B$1", "Daten!$B$2:$B$13", "Daten!$A$2:$A$13"),
                                      ("Kosten", "Daten!$C$1", "Daten!$C$2:$C$13", "Daten!$A$2:$A$13"))]) == {}
    # Ohne Beschriftungen, ohne Blatt im Bezug, waagrecht.
    assert hints(cells, charts=[chart(("Umsatz", "", "$B$2:$B$13", ""))]) == {}
    flat = {(1, c): Cell(m) for c, m in enumerate(["Jahr", "2021", "2022", "2023"], start=1)}
    flat.update({(2, 1): Cell("Fälle"), (2, 2): Cell("12"), (2, 3): Cell("15"), (2, 4): Cell("19")})
    assert hints(flat, charts=[chart(("Fälle", "Daten!$A$2", "Daten!$B$2:$D$2", "Daten!$B$1:$D$1"))]) == {}


def test_a_chart_that_ends_before_the_table():
    cells = months()
    hint = hints(cells, charts=[chart(("Umsatz", "Daten!$B$1", "Daten!$B$2:$B$10", "Daten!$A$2:$A$10"),
                                      ("Kosten", "Daten!$C$1", "Daten!$C$2:$C$10", "Daten!$A$2:$A$10"))])["chart"]
    assert hint.level == "fehler" and hint.key == "chart:B2"
    assert "Die Reihe „Umsatz“ zeigt nur B2:B10, die Tabelle geht aber bis B13: „Oktober“, „November“ und „Dezember“ " \
        "fehlen im Diagramm" in hint.text and "Die Reihe „Kosten“ zeigt nur C2:C10" in hint.text
    assert hint.action["kind"] == "series" and hint.action["chart"] == "Diagramm 1"
    assert [(c["index"], c["values"], c["categories"], c["undo"]) for c in hint.action["series"]] == [
        (0, "B2:B13", "A2:A13", {"values": "B2:B10", "categories": "A2:A10"}),
        (1, "C2:C13", "A2:A13", {"values": "C2:C10", "categories": "A2:A10"})]
    assert hint.action["series"][0]["was"] == {"values": "Daten!$B$2:$B$10", "categories": "Daten!$A$2:$A$10"}
    assert "B2:B13 mit den Beschriftungen A2:A13" in hint.tip


@pytest.mark.parametrize("title", ["Umsatz Januar bis September", "Umsatz Q1 bis Q3", "Umsatz 2025", "Top 9"])
def test_a_title_that_names_a_part_is_quiet(title):
    assert hints(months(), charts=[chart(("Umsatz", "", "Daten!$B$2:$B$10", "Daten!$A$2:$A$10"), title=title)]) == {}


def test_the_total_row_in_a_chart():
    hint = hints(months(), charts=[chart(("Kosten", "Daten!$C$1", "Daten!$C$2:$C$14", "Daten!$A$2:$A$14"))])["chart"]
    assert "nimmt mit C14 auch die Summenzeile mit" in hint.text
    assert hint.action["series"][0]["values"] == "C2:C13" and hint.action["series"][0]["categories"] == "A2:A13"
    # Eine SUMME ohne Wort davor zählt auch.
    cells = months()
    cells[(14, 1)] = Cell("")
    del cells[(14, 1)]
    assert "Summenzeile" in hints(cells, charts=[chart(("Kosten", "", "Daten!$C$2:$C$14", ""))])["chart"].text


def test_values_shifted_against_their_labels():
    hint = hints(months(), charts=[chart(("Umsatz", "", "Daten!$B$3:$B$14", "Daten!$A$2:$A$13"))])["chart"]
    assert "um eine Zeile verschoben: Bei „Januar“ steht der Wert aus B3 („Februar“), bei „Dezember“ sogar die " \
        "Summenzeile B14" in hint.text
    assert hint.action["series"][0]["values"] == "B2:B13"
    hint = hints(months(), charts=[chart(("Umsatz", "", "Daten!$B$2:$B$13", "Daten!$A$3:$A$14"))])["chart"]
    assert "Bei „Februar“ steht der Wert aus B2 („Januar“)" in hint.text
    assert hint.action["series"][0]["categories"] == "A2:A13"


def test_a_chart_that_starts_late_is_worth_a_look():
    hint = hints(months(), charts=[chart(("Umsatz", "", "Daten!$B$5:$B$13", "Daten!$A$5:$A$13"))])["chart"]
    assert hint.level == "pruefen" and "fängt erst bei B5 an, die Tabelle schon bei B2: „Januar“, „Februar“ und " \
        "„März“ fehlen" in hint.text


def test_a_series_named_after_another_column():
    hint = hints(months(), charts=[chart(("Kosten", "Daten!$C$1", "Daten!$B$2:$B$13", "Daten!$A$2:$A$13"))])["chartname"]
    assert hint.level == "pruefen" and "Die Reihe „Kosten“ im Diagramm hat ihren Namen aus C1, zeigt aber die Werte " \
        "unter „Umsatz“ (B2:B13)" in hint.text


def test_a_chart_on_another_sheet_names_the_sheet():
    cells = months()
    sheet = Sheet("Bericht", {}, charts=[chart(("Umsatz", "", "Daten!$B$2:$B$10", "Daten!$A$2:$A$10"))])
    found = pivotdiagramm.checks("Bericht", sheet, {"Bericht": {}, "Daten": cells}, live=True)
    assert len(found) == 1 and "Daten!B2:B10" in found[0].text and found[0].cell == ""
    assert found[0].action["series"][0]["sheet"] == "Daten"


def test_a_chart_over_a_pivot_table_ignores_the_grand_total():
    cells, described = fresh()
    c = chart(("Summe", "", "Daten!$G$2:$G$4", "Daten!$F$2:$F$4"))
    assert hints(cells, [described], [c]) == {}


# --- aus der Datei ---

PIVOT_TABLE = (
    f'<pivotTableDefinition xmlns="{MAIN}" name="Umsatz je Region" cacheId="1" dataCaption="Werte" colGrandTotals="0">'
    '<location ref="F1:G4" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>'
    '<pivotFields count="3"><pivotField axis="axisRow"><items count="4"><item x="0"/><item x="1"/><item x="2" h="1"/>'
    '<item t="default"/></items></pivotField><pivotField axis="axisPage"><items count="4"><item x="0"/><item x="1"/>'
    '<item x="2"/><item t="default"/></items></pivotField><pivotField dataField="1"/></pivotFields>'
    '<rowFields count="1"><field x="0"/></rowFields><pageFields count="1"><pageField fld="1" item="0"/></pageFields>'
    '<dataFields count="1"><dataField name="Mittelwert von Umsatz" fld="2" subtotal="average"/></dataFields>'
    '</pivotTableDefinition>'
)
PIVOT_CACHE = (
    f'<pivotCacheDefinition xmlns="{MAIN}"><cacheSource type="worksheet"><worksheetSource ref="A1:C9" sheet="Daten"/>'
    '</cacheSource><cacheFields count="3"><cacheField name="Region"><sharedItems><s v="Nord"/><s v="Ost"/><s v="Süd"/>'
    '</sharedItems></cacheField><cacheField name="Produkt"><sharedItems><s v="Maus"/><s v="Monitor"/><s v="Tastatur"/>'
    '</sharedItems></cacheField><cacheField name="Umsatz"><sharedItems containsNumber="1"/></cacheField></cacheFields>'
    '</pivotCacheDefinition>'
)


def test_reads_what_it_needs_from_the_file():
    cells = {"A1": "Region", "B1": "Produkt", "C1": "Umsatz", "F1": "Zeilenbeschriftungen", "G1": "Mittelwert von Umsatz",
             "F2": "Nord", "G2": 109.75, "F3": "Ost", "G3": 80, "F4": "Gesamtergebnis", "G4": 99.8333333333}
    for r, (a, b, c) in enumerate(ROWS, start=2):
        cells.update({f"A{r}": a, f"B{r}": b, f"C{r}": c})
    book = read_book(mappe(("Daten", cells, "", {"rP": ("pivotTable", "../pivotTables/pivotTable1.xml")}), files={
        "xl/pivotTables/pivotTable1.xml": PIVOT_TABLE,
        "xl/pivotTables/_rels/pivotTable1.xml.rels": _rels({"rK": ("pivotCacheDefinition", "../pivotCache/pivotCacheDefinition1.xml")}),
        "xl/pivotCache/pivotCacheDefinition1.xml": PIVOT_CACHE}))
    read = book.sheets[0].pivots[0]
    assert read.value_fields == (PivotValue("Mittelwert von Umsatz", "Umsatz", "average"),)
    assert read.body == (2, 7, 4, 7) and read.source_sheet == "Daten" and read.source_area == (1, 1, 9, 3)
    assert dict(read.skipped) == {"Region": ("Süd",), "Produkt": ("Monitor", "Tastatur")}
    assert read.grand == (True, False) and read.filtered is False and read.plain
    # Nur Maus, ohne Süd: Nord (120,5 und 99) 109,75, Ost 80, zusammen 99,83. So steht es da: still.
    sheet = book.sheets[0]
    assert pivotdiagramm.checks("Daten", sheet, {"Daten": sheet.cells}, live=False) == []
    sheet.cells[(3, 3)] = Cell("90")
    found = pivotdiagramm.checks("Daten", sheet, {"Daten": sheet.cells}, live=False)
    assert len(found) == 1 and "G3 „Ost“ zeigt 80, die Daten ergeben jetzt 90" in found[0].text


# --- im Seitenfenster ---


def office(cells: dict) -> dict:
    """Ein Blatt, wie pruefer.js es schickt."""
    rows = max(r for r, _ in cells)
    cols = max(c for _, c in cells)
    values = [[cells[(r, c)].value if (r, c) in cells else "" for c in range(1, cols + 1)] for r in range(1, rows + 1)]
    for row in values:
        for j, v in enumerate(row):
            try:
                row[j] = float(v.replace(",", "."))
            except (ValueError, AttributeError):
                pass
    formulas = [[f"={cells[(r, c)].formula}" if (r, c) in cells and cells[(r, c)].formula else None
                 for c in range(1, cols + 1)] for r in range(1, rows + 1)]
    return {"name": "Daten", "row": 1, "col": 1, "values": values, "formulas": formulas, "kinds": []}


def test_the_pane_reads_pivots_and_charts_and_gets_buttons(tmp_path):
    cells, described = fresh()
    cells[(3, 3)] = Cell("95")
    cells.update({(r + 14, c): v for (r, c), v in months(False).items()})
    sheet = office(cells)
    sheet["pivots"] = [{
        "name": "PivotTable1", "area": "Daten!F1:G5", "body": "Daten!G2:G5", "rows": ["Region"], "columns": [], "filters": [],
        "values": [{"caption": "Summe von Umsatz", "field": "Umsatz", "function": "Sum", "showAs": "None"}],
        "skipped": {}, "grand": [True, True], "filtered": False, "source": "Daten!$A$1:$C$9", "table": "",
    }]
    sheet["charts"] = [{"name": "Diagramm 1", "title": "", "series": [
        {"name": "Umsatz", "values": "=Daten!$B$16:$B$24", "categories": "=Daten!$A$16:$A$24"}]}]
    payload = {"book": "Verkauf.xlsx", "active": "Daten", "sheets": [sheet], "tables": []}
    _, _, read = sheets_from(payload)
    assert read[0].pivots[0].value_fields[0].function == "sum" and read[0].pivots[0].source_area == (1, 1, 9, 3)
    assert read[0].charts[0].series[0].values == "=Daten!$B$16:$B$24"
    app = create_app(Settings(data_dir=tmp_path))
    client = TestClient(app)
    result = client.post("/api/addin/pruefen", json=payload, headers={"X-VoiceAgent": "1"}).json()
    by_key = {h["key"].split(":")[0]: h for h in result["hints"]}
    assert by_key["pivot"]["action"] == {"kind": "refresh", "pivot": "PivotTable1"} and by_key["pivot"]["cell"] == "G3"
    assert by_key["chart"]["action"]["series"][0]["values"] == "B16:B27"
    assert "„Oktober“, „November“ und „Dezember“" in by_key["chart"]["text"]


@pytest.mark.parametrize("name", ["Werte", "Values", "Σ Werte"])
def test_the_values_field_excel_names_among_the_columns_is_not_a_field(name, tmp_path):
    # Excel für Mac nennt bei columnHierarchies auch „Werte“, schon bei einem Wertfeld.
    cells, described = fresh()
    cells[(3, 3)] = Cell("95")
    cells[(7, 3)] = Cell("")
    sheet = office({**cells, (12, 1): Cell("Nord"), (12, 2): Cell("Maus"), (12, 3): Cell("40")})
    sheet["pivots"] = [{
        "name": "PivotTable1", "area": "Daten!F1:G5", "body": "Daten!G2:G5", "rows": ["Region"], "columns": [name],
        "filters": [], "values": [{"caption": "Summe von Umsatz", "field": "Umsatz", "function": "Automatic",
                                   "showAs": "None"}],
        "skipped": {}, "grand": [True, True], "filtered": False, "source": "Daten!$A$1:$C$9", "table": "",
    }]
    _, _, read = sheets_from({"book": "Verkauf.xlsx", "active": "Daten", "sheets": [sheet], "tables": []})
    found = {f.key.split(":")[0]: f for f in pivotdiagramm.checks("Daten", read[0], {"Daten": read[0].cells}, {}, True)}
    assert "G3 „Ost“ zeigt 355, die Daten ergeben jetzt 370" in found["pivot"].text
    assert "die Daten ergeben jetzt 875,75" in found["pivot"].text


def test_a_column_of_the_data_named_like_the_values_field_stays_a_field():
    heads = ("Region", "Werte", "Umsatz")
    block, described = pivot(heads=heads, across=1)
    cells = {**data(heads=heads), **block}
    assert described.columns == ("Werte",)
    assert hints(cells, [described]) == {}
    cells[(3, 3)] = Cell("95")
    assert "„Ost“ / „Maus“" in hints(cells, [described])["pivot"].text


def test_values_stacked_in_the_rows_are_not_recomputed():
    from dataclasses import replace
    cells, described = fresh(functions=("sum", "count"))
    cells[(3, 3)] = Cell("95")
    assert "pivot" in hints(cells, [replace(described, columns=("Werte",))])
    assert "pivot" not in hints(cells, [replace(described, rows=described.rows + ("Werte",))])


def test_tables_from_the_pane_are_a_source(tmp_path):
    rows = ROWS + [("Nord", "Maus", 40.0)]
    cells, described = fresh(rows=rows)
    sheet = office(cells)
    sheet["pivots"] = [{
        "name": "PivotTable1", "area": "Daten!F1:G5", "body": "Daten!G2:G5", "rows": ["Region"], "columns": [],
        "values": [{"caption": "Summe von Umsatz", "field": "Umsatz", "function": "Sum"}], "grand": [True, True],
        "filtered": None, "source": "", "table": "Verkauf",
    }]
    payload = {"book": "Verkauf.xlsx", "active": "Daten", "sheets": [sheet],
               "tables": [{"name": "Verkauf", "sheet": "Daten", "address": "Daten!A1:C10", "totals": False}]}
    _, _, read = sheets_from(payload)
    assert read[0].tables[0].area == (1, 1, 10, 3)
    result = pruefer.sheet_result("Verkauf.xlsx", "Daten", {"Daten": read[0].cells}, None, read[0],
                                  pivotdiagramm.tables_of(read), live=True)
    assert not [h for h in result["hints"] if h["key"].startswith("pivot")]


def test_a_file_check_has_no_buttons():
    cells, described = fresh()
    cells[(3, 3)] = Cell("95")
    sheet = Sheet("Daten", cells, pivots=[described])
    result = pruefer.sheet_result("Verkauf.xlsx", "Daten", {"Daten": cells}, None, sheet, {})
    hint = next(h for h in result["hints"] if h["key"].startswith("pivot:"))
    assert "action" not in hint


def test_column_letters_in_actions_are_plain():
    assert area_text((2, 28, 13, 28)) == "AB2:AB13" and column_letters(28) == "AB"


def test_practice_workbook_has_a_stale_pivot_and_wrong_charts():
    """web/demo/pruefer-uebung.xlsx, Blatt „Pivot und Diagramm“: eine Pivot-Tabelle, die nach einer Korrektur und drei
    neuen Zeilen nicht aktualisiert wurde, eine mit Anzahl statt Summe, ein zu kurzes Diagramm und eines mit der
    Summenzeile."""
    from pathlib import Path

    data = (Path(__file__).parent.parent / "src/voiceagent/web/demo/pruefer-uebung.xlsx").read_bytes()
    sheet = next(s for s in read_book(data).sheets if s.name == "Pivot und Diagramm")
    assert [p.name for p in sheet.pivots] == ["Umsatz je Region", "Umsätze je Produkt"] and len(sheet.charts) == 2
    found = pruefer.Checker(None).check_file("pruefer-uebung.xlsx", data)
    hints = {h["key"]: h for s in found["sheets"] if s["name"] == "Pivot und Diagramm" for h in s["hints"]}
    assert {"pivot:F4", "pivotsource:A15", "pivotcount:F11", "chart:B21", "chart:B21~2", "missing:C9"} == set(hints)
    assert "F4 „Nord“ zeigt 613,4, die Daten ergeben jetzt 652,4" in hints["pivot:F4"]["text"]
    assert "Mit ihnen wäre „Ost“ 604,5 statt 349,5 und es käme „West“ (345,5) dazu." in hints["pivotsource:A15"]["text"]
    assert "„Oktober“, „November“ und „Dezember“ fehlen" in hints["chart:B21"]["text"]
    assert "Summenzeile" in hints["chart:B21~2"]["text"] and hints["chart:B21~2"]["cell"] == "B21"
    assert all("action" not in h for h in hints.values())
