"""Excel-Datei mitlesen: Formeln und geänderte Werte direkt aus der .xlsx-Datei."""

import io
import zipfile
from xml.sax.saxutils import escape

import pytest

from voiceagent.workbook import WorkbookTracker, layout, pattern, read_xlsx, readable, shift

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def xlsx(*sheets, styles: str = "") -> bytes:
    """Eine kleine .xlsx-Datei wie aus Excel. Je Blatt (Name, {Adresse: Wert}).

    Ein Wert ist eine Zahl, ein Text, ("=", Formel), ("=", Formel, Ergebnis), ("=", Formel, Ergebnis, Stil),
    ("shared", si, Formel, Ref) für den Anfang einer heruntergezogenen Formel oder ("shared", si) für die übrigen
    Zellen, oder (Zahl, Stil) für ein Datum oder Prozent."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        entries = "".join(f'<sheet name="{escape(n)}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, (n, _) in enumerate(sheets))
        archive.writestr("xl/workbook.xml", f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets>{entries}</sheets></workbook>')
        rels = "".join(
            f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>' for i in range(len(sheets))
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>',
        )
        if styles:
            archive.writestr("xl/styles.xml", f'<styleSheet xmlns="{MAIN}">{styles}</styleSheet>')
        for i, (_, cells) in enumerate(sheets):
            rows: dict[int, list[str]] = {}
            for address, value in cells.items():
                row = int("".join(ch for ch in address if ch.isdigit()))
                rows.setdefault(row, []).append(_cell(address, value))
            body = "".join(f'<row r="{r}">{"".join(c)}</row>' for r, c in sorted(rows.items()))
            archive.writestr(f"xl/worksheets/sheet{i + 1}.xml", f'<worksheet xmlns="{MAIN}"><sheetData>{body}</sheetData></worksheet>')
    return buffer.getvalue()


def _cell(address, value) -> str:
    if isinstance(value, str):
        return f'<c r="{address}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'
    if isinstance(value, tuple) and value[0] == "=":
        result = f"<v>{value[2]}</v>" if len(value) > 2 else ""
        style = f' s="{value[3]}"' if len(value) > 3 else ""
        return f'<c r="{address}"{style}><f>{escape(value[1])}</f>{result}</c>'
    if isinstance(value, tuple) and value[0] == "shared":
        if len(value) == 2:
            return f'<c r="{address}"><f t="shared" si="{value[1]}"/></c>'
        return f'<c r="{address}"><f t="shared" si="{value[1]}" ref="{value[3]}">{escape(value[2])}</f></c>'
    if isinstance(value, tuple):
        return f'<c r="{address}" s="{value[1]}"><v>{value[0]}</v></c>'
    return f'<c r="{address}"><v>{value}</v></c>'


def wahl(**extra):
    """Übungsmappe: Sitzverteilung im französischen Parlament (577 Sitze)."""
    cells = {
        "A1": "1. Wahlgang",
        "A2": "Parteien", "B2": "Wähleranteil(%)", "C2": "Sitze(577)",
        "A3": "RN", "B3": 33.2, "A4": "NFP", "B4": 28, "A5": "Ensemble", "B5": 20.8,
        "A6": "Républicaine", "B6": 6.6, "A7": "Rest", "B7": 11.4,
    }
    cells.update(extra)
    return xlsx(("Tabelle1", cells))


FILLED = {
    "C3": ("shared", "0", "B3*577/100", "C3:C7"), "C4": ("shared", "0"), "C5": ("shared", "0"),
    "C6": ("shared", "0"), "C7": ("shared", "0"),
}


def test_reads_values_formulas_and_the_table():
    sheet = read_xlsx(wahl(**FILLED))[0]
    assert sheet.cells[(3, 2)].value == "33,2"
    assert sheet.cells[(5, 3)].formula == "B5*577/100"  # heruntergezogen
    plan = layout(sheet)
    assert (plan.header_row, plan.headers, plan.names[6], plan.last_row) == (
        2, {1: "Parteien", 2: "Wähleranteil(%)", 3: "Sitze(577)"}, "Républicaine", 7)
    assert readable("B5*577/100", (5, 3), plan) == "=Wähleranteil(%) * 577 / 100"
    assert readable("SUM(C3:C7)", (8, 3), plan) == "=SUMME(alle Sitze(577))"
    assert readable("ROUND(B3/$B$8,2)", (3, 4), plan) == "=RUNDEN(Wähleranteil(%) / $B$8; 2)"
    assert readable('IF(B3>5,"ja","nein, B3")', (3, 4), plan) == '=WENN(Wähleranteil(%)>5; "ja"; "nein, B3")'
    # Andere Blätter haben andere Spaltenköpfe.
    assert readable("B3*Gewichte!B3:B7", (3, 4), plan) == "=Wähleranteil(%) * Gewichte!B3:B7"
    assert shift("Gewichte!B3*$C$1+B3", 2, 0) == "Gewichte!B5*$C$1+B5"
    # Herunterziehen ändert die Adressen, nicht die Rechnung.
    assert pattern("B3*577/100", (3, 3)) == pattern("B7*577/100", (7, 3)) != pattern("B3*577/100", (7, 3))


def test_dates_percent_and_texts():
    styles = ('<numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts>'
              '<cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="164"/><xf numFmtId="10"/></cellXfs>')
    sheet = read_xlsx(xlsx(("Daten", {"A1": (45950, 1), "B1": (0.332, 2), "C1": "  Text  mit Lücken "}), styles=styles))[0]
    assert [sheet.cells[(1, c)].value for c in (1, 2, 3)] == ["20.10.2025", "33,2 %", "Text mit Lücken"]


def test_typing_data_is_not_a_decision_but_pasting_is_one_event():
    tracker = WorkbookTracker()
    tracker.update("wahl.xlsx", read_xlsx(xlsx(("Tabelle1", {"A1": "Parteien", "B1": "Anteil", "A2": "RN"}))), t=1.0)
    assert tracker.update("wahl.xlsx", read_xlsx(xlsx(("Tabelle1", {"A1": "Parteien", "B1": "Anteil", "A2": "RN", "B2": 33.2}))), t=2.0) == []
    tracker.update("wahl.xlsx", read_xlsx(wahl()), t=3.0)
    pasted = read_xlsx(wahl(B3=1, B4=2, B5=3, B6=4))
    assert [e.kind for e in tracker.update("wahl.xlsx", pasted, t=4.0)] == ["record_changed"]
    # Eine andere Datei fängt neu an.
    assert tracker.update("andere.xlsx", read_xlsx(wahl(**FILLED)), t=5.0) == []


def test_masks_like_the_screen():
    from voiceagent.privacy import Privacy

    privacy = Privacy(("Geburtsdatum",))
    tracker = WorkbookTracker(privacy.line)
    cells = {"A1": "Name", "B1": "Geburtsdatum", "C1": "E-Mail", "D1": "Status", "A2": "Fall 1", "B2": "01.02.1990",
             "C2": "anna@example.org", "D2": "offen"}
    tracker.update("p.xlsx", read_xlsx(xlsx(("P", cells))), t=1.0)
    events = tracker.update("p.xlsx", read_xlsx(xlsx(("P", {**cells, "C2": "lena@example.org", "D2": "erledigt"}))), t=2.0)
    # Geschwärzt sehen beide Adressen gleich aus: Nur der Status hat sich geändert.
    assert [e.label for e in events] == ["Status"]
    text = " ".join([events[0].text, *events[0].context])
    assert "example.org" not in text and "1990" not in text


@pytest.mark.parametrize("data", [b"", b"kein zip", xlsx(("A", {}))[:50]])
def test_not_an_excel_file(data):
    with pytest.raises(ValueError):
        read_xlsx(data)


# --- In der Aufnahme ---


def test_a_setting_beside_the_header_row_keeps_the_table():
    from voiceagent.xlsx import Cell, Sheet

    cells = {(1, 1): Cell("Monat"), (1, 2): Cell("Patienten"), (1, 7): Cell("Zinssatz"), (1, 8): Cell("3 %")}
    for r, month in enumerate(["Januar", "Februar", "März", "April", "Mai", "Juni"], start=2):
        cells[(r, 1)], cells[(r, 2)] = Cell(month), Cell(str(400 + r))
    # Rechts in der ersten Datenzeile ein Wort: Die Zeile sieht wie eine Kopfzeile aus, hat aber schon ihre Zahl.
    cells[(2, 7)], cells[(2, 8)] = Cell("Summe mit Zins"), Cell("2500", "SUM(B2:B7)*(1+H1)")
    plan = layout(Sheet("S", cells))
    assert plan.header_row == 1 and plan.headers[1] == "Monat" and plan.headers[2] == "Patienten"
    assert plan.names[2] == "Januar"
