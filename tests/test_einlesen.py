"""Tabellen aus Dateien in ein neues Excel-Blatt: CSV, Excel, Word, PDF."""

import io
import zipfile
from xml.sax.saxutils import escape

import pytest
from test_excel_structure import mappe
from test_helfer import FakeLink, excel_shows, snapshot_of, wahl

from voiceagent.einlesen import date, decimal_mark, draft, number, tables_from
from voiceagent.excel import Snapshot
from voiceagent.helfer import Helper
from voiceagent.schreiben import applescript
from voiceagent.xlsx import Cell


def pdf_with(lines: list[list[tuple[int, str]]]) -> bytes:
    """Ein kleines PDF: Text an festen Stellen, wie eine Tabelle in einem Bericht."""
    content = "BT /F1 10 Tf\n"
    y = 750
    for line in lines:
        for x, text in line:
            content += f"1 0 0 1 {x} {y} Tm ({text}) Tj\n"
        y -= 14
    content += "ET"
    objects = [
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        f"<< /Length {len(content)} >>\nstream\n{content}\nendstream",
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
    ]
    out, offsets = "%PDF-1.4\n", []
    for i, body in enumerate(objects, start=1):
        offsets.append(len(out.encode("latin-1")))
        out += f"{i} 0 obj\n{body}\nendobj\n"
    xref = len(out.encode("latin-1"))
    out += f"xref\n0 {len(objects) + 1}\n0000000000 65535 f \n" + "".join(f"{o:010d} 00000 n \n" for o in offsets)
    out += f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n"
    return out.encode("latin-1")


def docx_with(*tables: list[list[str]], text: str = "Bericht") -> bytes:
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = f"<w:p><w:r><w:t>{escape(text)}</w:t></w:r></w:p>"
    for table in tables:
        rows = "".join("<w:tr>" + "".join(f"<w:tc><w:p><w:r><w:t>{escape(c)}</w:t></w:r></w:p></w:tc>" for c in row) + "</w:tr>"
                       for row in table)
        body += f"<w:tbl>{rows}</w:tbl>"
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("word/document.xml", f'<w:document xmlns:w="{w}"><w:body>{body}</w:body></w:document>')
    return buffer.getvalue()


# --- Werte erkennen ---


@pytest.mark.parametrize("text, mark, expected", [
    ("1.234,50", ",", (1234.5, False)),
    ("1'234.50", ".", (1234.5, False)),
    ("12,5 %", ",", (0.125, True)),
    ("CHF 80.–", ".", (80.0, False)),
    ("-3,2", ",", (-3.2, False)),
    ("1,234.5", ".", (1234.5, False)),
    ("2 500", ",", (2500.0, False)),
])
def test_numbers_in_every_spelling(text, mark, expected):
    assert number(text, mark) == expected


@pytest.mark.parametrize("text", ["08001", "0041 44 123 45 67", "CH93 0076 2011 6238 5295 7", "1234567890123456", "1.2.3", "abc"])
def test_what_only_looks_like_a_number_stays_text(text):
    assert number(text, ",") is None


def test_dates_and_decimal_mark():
    assert date("1.3.2024") == (2024, 3, 1) and date("2024-03-01") == (2024, 3, 1) and date("31/12/24") is None
    assert date("01.03.24") == (2024, 3, 1) and date("32.01.2024") is None
    assert decimal_mark([["1,5", "2,25"], ["3.5", "x"]]) == ","
    assert decimal_mark([["1.5", "2.25"], ["12,000"]]) == "."


# --- Dateien ---


def test_german_csv_becomes_numbers_dates_and_text():
    data = "Kunde;PLZ;Datum;Betrag;Anteil\nMüller AG;08001;01.03.2024;1.234,50;12,5 %\nHuber;3000;2.3.2024;-80;7 %\n".encode("cp1252")
    [table] = tables_from("Umsatz März.csv", data)
    assert table.kind == "CSV" and table.size == (3, 5) and table.rows[1][0] == "Müller AG"
    result = draft(table, ["Tabelle1"], "Umsatz März.csv")
    assert result.name == "Umsatz März"
    assert result.grid[1] == ["Müller AG", "'08001", "2024-03-01", "1234.5", "12.5%"]
    assert result.grid[2] == ["Huber", "3000", "2024-03-02", "-80", "7%"]
    assert result.expected[(2, 3)] == "01.03.2024" and result.expected[(2, 4)] == 1234.5 and result.expected[(2, 5)] == 0.125
    assert ("C2:C3", "dd.mm.yyyy") in result.formats and result.bold == ["A1:E1"]
    assert any("Null vorne" in n for n in result.notes) and any("Komma" in n for n in result.notes)


def test_english_csv_with_commas():
    [table] = tables_from("export.csv", b'name,amount,date\n"Smith, J.",1234.5,2024-03-01\nLee,"2,000.25",2024-03-02\n')
    result = draft(table, [], "export.csv")
    assert result.grid[1] == ["Smith, J.", "1234.5", "2024-03-01"] and result.grid[2][1] == "2000.25"


def test_excel_file_gives_each_sheet():
    data = mappe(("Daten", {"A1": "Station", "B1": "Fälle", "A2": "Kardio", "B2": 12, "A3": "Notfall", "B3": 31}, "", {}),
                 ("Leer", {}, "", {}))
    tables = tables_from("Fälle.xlsx", data)
    assert [t.name for t in tables] == ["Daten"] and tables[0].rows == [["Station", "Fälle"], ["Kardio", "12"], ["Notfall", "31"]]


def test_word_tables_each_count():
    data = docx_with([["Station", "Betten"], ["Kardio", "24"]], [["Name", "Telefon"], ["Huber", "044 123 45 67"], ["Keller", "031 1"]])
    tables = tables_from("Bericht.docx", data)
    assert [t.name for t in tables] == ["Tabelle 1", "Tabelle 2"] and tables[1].rows[1] == ["Huber", "044 123 45 67"]
    result = draft(tables[1], [], "Bericht.docx")
    assert result.name == "Bericht Tabelle 2" and result.grid[1][1] == "'044 123 45 67"


def test_pdf_table_from_a_report():
    data = pdf_with([[(72, "Bericht Quartal 3")], [(72, "Station"), (200, "Faelle"), (300, "Kosten")],
                     [(72, "Kardio"), (200, "12"), (300, "1'200.50")], [(72, "Chirurgie"), (200, "7"), (300, "980.00")],
                     [(72, "Notfall"), (200, "31"), (300, "2'310.75")], [(72, "Ende.")]])
    [table] = tables_from("Quartal.pdf", data)
    assert table.kind == "PDF" and table.rows[0] == ["Station", "Faelle", "Kosten"] and table.rows[1] == ["Kardio", "12", "1'200.50"]
    result = draft(table, [], "Quartal.pdf")
    assert result.grid[3] == ["Notfall", "31", "2310.75"]


def test_files_without_tables_say_why():
    with pytest.raises(ValueError, match="keine Tabelle"):
        tables_from("Brief.docx", docx_with(text="Sehr geehrte Damen und Herren"))
    with pytest.raises(ValueError, match="Alte Office-Dateien"):
        tables_from("alt.xls", b"x")
    with pytest.raises(ValueError, match="Ich lese CSV"):
        tables_from("Bild.png", b"x")
    big = tables_from("gross.csv", ("a;b\n" + "1;2\n" * 30_000).encode())[0]
    with pytest.raises(ValueError, match="mehr als 50000"):
        draft(big, [], "gross.csv")


def test_umlauts_reach_excel_as_character_ids():
    [table] = tables_from("Umsatz.csv", "Kunde;Betrag\nMüller;5\n".encode())
    script = applescript(draft(table, [], "Umsatz.csv"), "Mappe1")
    assert '("M" & (character id 252) & "ller")' in script


# --- Der ganze Ablauf im Helfer ---


def test_upload_preview_write_and_check():
    link = FakeLink(snapshot_of(wahl(), sheets=("Daten", "Umsatz")), answer="sheet\tUmsatz 2\npivot\tfalse")
    helper = Helper(link)
    data = "Kunde;Betrag\nMüller;1.234,50\nHuber;80\n".encode()
    shown = helper.load("Umsatz.csv", data)
    assert shown["ok"] and shown["sheet"] == "Umsatz 2" and shown["tables"] == [{"name": "Umsatz", "rows": 3, "columns": 2, "kind": "CSV"}]
    assert shown["preview"][1] == ["Müller", "1.234,50"]
    pending = helper._pending[shown["token"]]
    link.after = Snapshot("Wahl.xlsx", "", "Umsatz 2", [], {**excel_shows(pending.draft, {k: v for k, v in pending.draft.expected.items() if isinstance(v, float)}),
                                                              (1, 1): Cell("Kunde"), (1, 2): Cell("Betrag"), (2, 1): Cell("Müller"), (3, 1): Cell("Huber")})
    done = helper.run(shown["token"])
    assert done["ok"] and "alle 6 Werte in Excel nachgeprüft" in done["text"]
    # Ein importiertes Blatt darf danach Quelle einer Übersicht sein.
    assert helper._made == {}


def test_upload_with_several_tables_and_wrong_book():
    link = FakeLink(snapshot_of(wahl()))
    helper = Helper(link)
    shown = helper.load("Bericht.docx", docx_with([["A", "B"], ["1", "2"]], [["Name", "Zahl"], ["x", "3"], ["y", "4"]]))
    assert len(shown["tables"]) == 2 and shown["index"] == 0
    second = helper.table(shown["file"], 1)
    assert second["ok"] and second["preview"][0] == ["Name", "Zahl"]
    link.snapshot = Snapshot("Andere.xlsx", "", "Tabelle1", ["Tabelle1"], {})
    result = helper.run(second["token"])
    assert result["state"] == "moved" and "Wahl.xlsx" in result["text"] and link.scripts == []
    assert helper.table("falsch", 0)["state"] == "expired"
    assert helper.load("Brief.docx", docx_with(text="Hallo"))["text"] == "In der Datei finde ich keine Tabelle."
