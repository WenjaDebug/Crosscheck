"""Der Aufbau einer Mappe: Excel-Tabellen, Filter, Sortierung, Hervorhebungen, Auswahllisten,
Diagramme, Pivot-Tabellen, Namen und Blätter, und wie VoiceAgent danach fragt."""

import io
import zipfile
from xml.sax.saxutils import escape, quoteattr

from voiceagent.workbook import (
    WorkbookTracker,
    _as_number,
    layout,
    pattern,
    read_book,
    readable,
)
from voiceagent.xlsx import Cell, Sheet

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"


def _cell(address: str, value) -> str:
    if isinstance(value, str):
        return f'<c r="{address}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'
    if isinstance(value, tuple):
        return f'<c r="{address}"><f>{escape(value[1])}</f>{f"<v>{value[2]}</v>" if len(value) > 2 else ""}</c>'
    return f'<c r="{address}"><v>{value}</v></c>'


def _rels(links: dict[str, tuple[str, str]]) -> str:
    body = "".join(f'<Relationship Id="{i}" Type="{REL}/{kind}" Target="{target}"/>' for i, (kind, target) in links.items())
    return f'<Relationships xmlns="{PKG}">{body}</Relationships>'


def mappe(*sheets, names: dict[str, str] | None = None, files: dict[str, str] | None = None) -> bytes:
    """Eine .xlsx-Datei. Je Blatt (Name, Zellen, XML nach den Zellen, {rId: (Art, Ziel)}).

    files: weitere Teile der Datei, z. B. xl/tables/table1.xml."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        entries = "".join(f'<sheet name={quoteattr(s[0])} sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, s in enumerate(sheets))
        defined = "".join(f'<definedName name="{n}">{escape(ref)}</definedName>' for n, ref in (names or {}).items())
        defined = f"<definedNames>{defined}</definedNames>" if defined else ""
        archive.writestr("xl/workbook.xml", f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets>{entries}</sheets>{defined}</workbook>')
        archive.writestr("xl/_rels/workbook.xml.rels", _rels({f"rId{i + 1}": ("worksheet", f"worksheets/sheet{i + 1}.xml")
                                                              for i in range(len(sheets))}))
        for i, (_, cells, extra, links) in enumerate(sheets):
            rows: dict[int, list[str]] = {}
            for address, value in cells.items():
                rows.setdefault(int("".join(ch for ch in address if ch.isdigit())), []).append(_cell(address, value))
            body = "".join(f'<row r="{r}">{"".join(c)}</row>' for r, c in sorted(rows.items()))
            archive.writestr(
                f"xl/worksheets/sheet{i + 1}.xml",
                f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}" xmlns:x14="{X14}" xmlns:xm="{XM}"><sheetData>{body}</sheetData>{extra}</worksheet>',
            )
            if links:
                archive.writestr(f"xl/worksheets/_rels/sheet{i + 1}.xml.rels", _rels(links))
        for path, content in (files or {}).items():
            archive.writestr(path, content)
    return buffer.getvalue()


# Die Übungsmappe, so wie sie in Excel steht: Titelzeile, Kopfzeile in Zeile 4, eine Leerzeile.
WAHL = {
    "C2": "Wahlgang", "D2": "Sitze(577)",
    "B4": "Parteien", "C4": "Wähleranteil(%)", "D4": "Anzahl Sitze",
    "B6": "RN", "C6": 33.2, "D6": ("=", "0.332*577", 191.564),
    "B7": "NFP", "C7": 28, "D7": ("=", "0.28*577", 161.56),
    "B8": "Ensemble", "C8": 20.8, "D8": ("=", "0.208*577", 120.016),
    "B9": "Républicaine", "C9": 6.6, "D9": ("=", "0.066*577", 38.082),
    "B10": "Rest", "C10": 11.4, "D10": ("=", "0.114*577", 65.778),
    "D11": ("=", "SUM(D6:D10)", 577),
}


def sheet(cells: dict) -> Sheet:
    return read_book(mappe(("Tabelle1", cells, "", {}))).sheets[0]


# --- Tabelle erkennen ---


def test_header_below_a_title_and_a_blank_row():
    plan = layout(sheet(WAHL))
    assert plan.header_row == 4
    assert plan.headers == {2: "Parteien", 3: "Wähleranteil(%)", 4: "Anzahl Sitze"}
    assert plan.name_column == 2
    assert plan.names[6] == "RN" and plan.names[10] == "Rest"
    assert plan.last_row >= 10


def test_two_header_rows_become_one():
    plan = layout(sheet({
        "B1": "Umsatz", "C1": "Umsatz", "D1": "Kosten",
        "A2": "Filiale", "B2": "2024", "C2": "2025", "D2": "2025",
        "A3": "Bern", "B3": 10, "C3": 12, "D3": 4,
        "A4": "Basel", "B4": 8, "C4": 9, "D4": 3,
    }))
    assert plan.header_row == 2
    assert plan.headers[1] == "Filiale"
    assert "2025" in plan.headers[3] and "Umsatz" in plan.headers[3]
    assert "Kosten" in plan.headers[4]


def test_a_missing_number_in_the_first_row_keeps_the_header():
    # Die erste Datenzeile sieht ohne ihre Zahl aus wie eine Kopfzeile, aber West und Monitor stehen unten wieder.
    cells = {"A1": "Datum", "B1": "Region", "C1": "Produkt", "D1": "Umsatz"}
    for r, (region, product) in enumerate([("West", "Monitor"), ("West", "Dock"), ("Ost", "Monitor"), ("West", "Maus"),
                                           ("Ost", "Dock")], start=2):
        cells.update({f"A{r}": f"{r:02d}.01.2026", f"B{r}": region, f"C{r}": product})
        if r > 2:
            cells[f"D{r}"] = 100 * r
    plan = layout(sheet(cells))
    assert plan.header_row == 1 and plan.headers[4] == "Umsatz" and plan.last_row == 6


def test_row_labels_without_a_header():
    plan = layout(sheet({
        "B1": "Januar", "C1": "Februar", "D1": "März",
        "A2": "Miete", "B2": 1200, "C2": 1200, "D2": 1250,
        "A3": "Strom", "B3": 80, "C3": 95, "D3": 70,
    }))
    assert plan.header_row == 1
    assert plan.name_column == 1
    assert plan.names == {2: "Miete", 3: "Strom"}


# --- Excel-Tabellen und strukturierte Verweise ---

TABLE = (
    f'<table xmlns="{MAIN}" id="1" name="Rechnungen" displayName="Rechnungen" ref="A1:C4" totalsRowCount="1">'
    '<autoFilter ref="A1:C3"/><tableColumns count="3">'
    '<tableColumn id="1" name="Nummer"/><tableColumn id="2" name="Betrag"/>'
    '<tableColumn id="3" name="Mit MWST"><calculatedColumnFormula>Rechnungen[[#This Row],[Betrag]]*1.081</calculatedColumnFormula></tableColumn>'
    "</tableColumns></table>"
)


def rechnungen(extra: str = "", links: dict | None = None, files: dict | None = None, names=None) -> bytes:
    cells = {
        "A1": "Nummer", "B1": "Betrag", "C1": "Mit MWST",
        "A2": "5001", "B2": 4200, "C2": ("=", "Rechnungen[[#This Row],[Betrag]]*1.081", 4540.2),
        "A3": "5002", "B3": 6100, "C3": ("=", "Rechnungen[[#This Row],[Betrag]]*1.081", 6594.1),
        "B4": ("=", "SUBTOTAL(109,Rechnungen[Betrag])", 10300),
    }
    return mappe(
        ("Rechnungen", cells, extra + '<tableParts count="1"><tablePart r:id="rT"/></tableParts>',
         {"rT": ("table", "../tables/table1.xml"), **(links or {})}),
        names=names, files={"xl/tables/table1.xml": TABLE, **(files or {})},
    )


def test_excel_table_gives_the_columns_and_structured_references_read_like_names():
    book = read_book(rechnungen())
    table = book.sheets[0].tables[0]
    assert table.name == "Rechnungen" and table.columns == ("Nummer", "Betrag", "Mit MWST") and table.totals == 1
    plan = layout(book.sheets[0])
    assert plan.table == "Rechnungen" and plan.header_row == 1
    assert readable("Rechnungen[[#This Row],[Betrag]]*1.081", (2, 3), plan) == "=Betrag * 1,081"
    assert readable("SUBTOTAL(109,Rechnungen[Betrag])", (4, 2), plan) == "=TEILERGEBNIS(109; alle Betrag)"
    assert readable("Rechnungen[@Betrag]", (2, 3), plan) == "=Betrag"
    # Strukturierte Verweise zählen beim Herunterziehen nicht als Zelladressen.
    assert pattern("Rechnungen[@B2]+B2", (2, 3)) == "Rechnungen[@B2]+Z[0]S[-1]"


def test_overview_names_table_formulas_and_cells_outside():
    tracker = WorkbookTracker()
    tracker.update("Wahl.xlsx", [sheet(WAHL)], 0.0)
    text = tracker.overview()
    assert "Excel-Mappe „Wahl.xlsx“" in text
    assert "Kopfzeile in Zeile 4" in text and "Parteien, Wähleranteil(%), Anzahl Sitze" in text
    assert "5 Zeilen (RN, NFP, Ensemble, Républicaine, Rest)" in text
    assert "Formel Anzahl Sitze: =0,332 * 577 (bei RN; jede Zeile hat eine eigene Formel)" in text
    assert "C2 Wahlgang" in text and "D2 Sitze(577)" in text
    assert "Zeilen ohne Namen (Ergebnisse): D11 Anzahl Sitze =SUMME(Anzahl Sitze RN bis Rest)" in text


# --- Filter, Sortierung, Hervorhebung, Auswahllisten ---

FILTER = (
    '<autoFilter ref="A1:C3">'
    '<filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="5000"/></customFilters></filterColumn>'
    '<filterColumn colId="2"><filters><filter val="ja"/></filters></filterColumn>'
    '<sortState ref="A2:C3"><sortCondition descending="1" ref="B2:B3"/></sortState>'
    "</autoFilter>"
)
FORMATS = (
    '<conditionalFormatting sqref="B2:B3"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>5000</formula></cfRule></conditionalFormatting>'
    '<conditionalFormatting sqref="A2:C3"><cfRule type="expression" priority="2"><formula>$C2="nein"</formula></cfRule></conditionalFormatting>'
    '<dataValidations count="1"><dataValidation type="list" sqref="C2:C3"><formula1>"ja,nein"</formula1></dataValidation></dataValidations>'
    "<extLst><ext><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type=\"iconSet\" priority=\"3\">"
    '<x14:iconSet iconSet="3Stars"/></x14:cfRule><xm:sqref>A2:A3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>'
)


def bestellungen(extra: str = "") -> Sheet:
    cells = {
        "A1": "Nummer", "B1": "Betrag", "C1": "Bestellung",
        "A2": "5001", "B2": 4200, "C2": "ja", "A3": "5002", "B3": 6100, "C3": "nein",
    }
    return read_book(mappe(("Daten", cells, extra, {}))).sheets[0]


def parts(tracker: WorkbookTracker) -> dict[str, str]:
    return {p.label: p.text for p in (tracker._parts or {}).values()}


def test_filter_sort_formatting_and_lists_are_read_with_their_column_names():
    tracker = WorkbookTracker()
    tracker.update("Bestellungen.xlsx", [bestellungen(FILTER + FORMATS)], 0.0)
    found = parts(tracker)
    assert found["Filter Betrag"] == "> 5000"
    assert found["Filter Bestellung"] == "nur „ja“"
    assert found["Sortierung"] == "nach Betrag absteigend"
    assert found["Hervorhebung Betrag"] == "größer als 5000"
    assert found["Hervorhebung Nummer bis Bestellung"] == 'wenn Bestellung="nein"'
    assert found["Hervorhebung Nummer"] == "Symbole (Sterne)"
    assert found["Auswahlliste Bestellung"] == "ja, nein"


def test_a_new_filter_is_a_decision():
    tracker = WorkbookTracker()
    tracker.update("Bestellungen.xlsx", [bestellungen()], 0.0)
    events = tracker.update("Bestellungen.xlsx", [bestellungen(FILTER)], 1.0)
    labels = {e.label: e for e in events}
    assert set(labels) == {"Filter Betrag", "Filter Bestellung", "Sortierung"}
    event = labels["Filter Betrag"]
    assert (event.before, event.after) == ("keine", "> 5000")
    # Wieder weg: auch eine Entscheidung.
    gone = {e.label: e for e in tracker.update("Bestellungen.xlsx", [bestellungen()], 2.0)}
    assert gone["Filter Betrag"].after == "keine"


def test_list_from_cells_on_another_sheet():
    validation = (
        '<extLst><ext><x14:dataValidations count="1"><x14:dataValidation type="list">'
        "<x14:formula1><xm:f>Listen!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>C2:C3</xm:sqref>"
        "</x14:dataValidation></x14:dataValidations></ext></extLst>"
    )
    cells = {"A1": "Nummer", "B1": "Betrag", "C1": "Status", "A2": "5001", "B2": 4200, "C2": "offen"}
    book = read_book(mappe(("Daten", cells, validation, {}), ("Listen", {"A1": "offen", "A2": "bezahlt", "A3": "gemahnt"}, "", {})))
    tracker = WorkbookTracker()
    tracker.update("Mappe.xlsx", book.sheets, 0.0, names=book.names)
    assert parts(tracker)["Auswahlliste Status"] == "offen, bezahlt, gemahnt"


# --- Diagramme, Pivot-Tabellen, Namen, Blätter ---

CHART_XML = (
    f'<c:chartSpace xmlns:c="{CHART}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart>'
    "<c:title><c:tx><c:rich><a:p><a:r><a:t>Sitzverteilung</a:t></a:r></a:p></c:rich></c:tx></c:title>"
    '<c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/>'
    "<c:tx><c:strRef><c:f>Tabelle1!$D$4</c:f><c:strCache><c:pt idx=\"0\"><c:v>Anzahl Sitze</c:v></c:pt></c:strCache></c:strRef></c:tx>"
    "<c:cat><c:strRef><c:f>Tabelle1!$B$6:$B$10</c:f></c:strRef></c:cat>"
    "<c:val><c:numRef><c:f>Tabelle1!$D$6:$D$10</c:f></c:numRef></c:val>"
    "</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"
)
DRAWING_XML = (
    f'<xdr:wsDr xmlns:xdr="{DRAWING}" xmlns:c="{CHART}" xmlns:r="{REL}"><xdr:twoCellAnchor><xdr:graphicFrame>'
    '<xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Diagramm 1"/></xdr:nvGraphicFramePr>'
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData><c:chart r:id="rC"/></a:graphicData></a:graphic>'
    "</xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>"
)


def wahl_mit_diagramm() -> bytes:
    return mappe(
        ("Tabelle1", WAHL, '<drawing r:id="rD"/>', {"rD": ("drawing", "../drawings/drawing1.xml")}),
        names={"Sitze": "Tabelle1!$D$11"},
        files={
            "xl/drawings/drawing1.xml": DRAWING_XML,
            "xl/drawings/_rels/drawing1.xml.rels": _rels({"rC": ("chart", "../charts/chart1.xml")}),
            "xl/charts/chart1.xml": CHART_XML,
        },
    )


def test_chart_and_name_say_what_they_show():
    book = read_book(wahl_mit_diagramm())
    tracker = WorkbookTracker()
    tracker.update("Wahl.xlsx", book.sheets, 0.0, names=book.names)
    found = parts(tracker)
    assert found["Diagramm Sitzverteilung"] == "Säulendiagramm „Sitzverteilung“ mit Anzahl Sitze nach Parteien"
    assert found["Name Sitze"] == "Tabelle1!D11 (zurzeit 577)"


def test_new_chart_asks_what_to_see():
    tracker = WorkbookTracker()
    tracker.update("Wahl.xlsx", read_book(mappe(("Tabelle1", WAHL, "", {}))).sheets, 0.0)
    book = read_book(wahl_mit_diagramm())
    events = {e.label: e for e in tracker.update("Wahl.xlsx", book.sheets, 1.0, names=book.names)}
    assert "Diagramm Sitzverteilung" in events
    assert "Name Sitze" in events


PIVOT_TABLE = (
    f'<pivotTableDefinition xmlns="{MAIN}" name="PivotTable1" cacheId="1" dataCaption="Werte">'
    '<location ref="E1:F4" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>'
    '<pivotFields count="2"><pivotField axis="axisRow"><items count="3"><item x="0"/><item x="1" h="1"/><item t="default"/></items></pivotField>'
    '<pivotField dataField="1"/></pivotFields>'
    '<rowFields count="1"><field x="0"/></rowFields>'
    '<dataFields count="1"><dataField name="Summe von Betrag" fld="1"/></dataFields></pivotTableDefinition>'
)
PIVOT_CACHE = (
    f'<pivotCacheDefinition xmlns="{MAIN}"><cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="Daten"/></cacheSource>'
    '<cacheFields count="2"><cacheField name="Status"><sharedItems><s v="offen"/><s v="bezahlt"/></sharedItems></cacheField>'
    '<cacheField name="Betrag"><sharedItems containsNumber="1"/></cacheField></cacheFields></pivotCacheDefinition>'
)


def test_pivot_table_rows_values_hidden_items_and_its_cells_are_not_decisions():
    cells = {
        "A1": "Status", "B1": "Betrag", "A2": "offen", "B2": 4200, "A3": "bezahlt", "B3": 6100,
        "E1": "Zeilenbeschriftungen", "F1": "Summe von Betrag", "E2": "offen", "F2": 4200, "E3": "Gesamtergebnis", "F3": 4200,
    }

    def data(total: int) -> bytes:
        return mappe(
            ("Daten", {**cells, "F3": total}, "", {"rP": ("pivotTable", "../pivotTables/pivotTable1.xml")}),
            files={
                "xl/pivotTables/pivotTable1.xml": PIVOT_TABLE,
                "xl/pivotTables/_rels/pivotTable1.xml.rels": _rels({"rK": ("pivotCacheDefinition", "../pivotCache/pivotCacheDefinition1.xml")}),
                "xl/pivotCache/pivotCacheDefinition1.xml": PIVOT_CACHE,
            },
        )

    tracker = WorkbookTracker()
    tracker.update("Pivot.xlsx", read_book(data(4200)).sheets, 0.0)
    assert parts(tracker)["Pivot-Tabelle PivotTable1"] == (
        "Zeilen Status; Werte Summe von Betrag; Status ohne „bezahlt“; aus Daten A1:B3"
    )
    # Was die Pivot-Tabelle anzeigt, rechnet Excel aus: keine Entscheidung.
    assert tracker.update("Pivot.xlsx", read_book(data(9999)).sheets, 1.0) == []
    plan = layout(read_book(data(4200)).sheets[0])
    assert plan.headers == {1: "Status", 2: "Betrag"}


def test_sheets_added_renamed_and_deleted():
    tracker = WorkbookTracker()
    one = Sheet("Daten", {(1, 1): Cell("x")})
    tracker.update("M.xlsx", [one], 0.0)
    added = tracker.update("M.xlsx", [one, Sheet("Auswertung")], 1.0)
    assert [e.text for e in added] == ["Blatt Auswertung neu"]
    renamed = tracker.update("M.xlsx", [one, Sheet("Bericht")], 2.0)
    assert [e.text for e in renamed] == ["Blatt Auswertung umbenannt in Bericht"]
    deleted = tracker.update("M.xlsx", [one], 3.0)
    assert [e.text for e in deleted] == ["Blatt Bericht gelöscht"]


# --- Sortieren erkennen, Hinweise zu Formeln, Zahlen ---


def test_numbers_in_any_notation():
    assert _as_number("33.2") == 33.2
    assert _as_number("33,2") == 33.2
    assert _as_number("1.234,5") == 1234.5
    assert _as_number("1'234.5") == 1234.5
    assert _as_number("1,234.5") == 1234.5
    assert _as_number("12.345.678") == 12345678
    assert _as_number("6,6 %") == 6.6
    assert _as_number("RN") is None


