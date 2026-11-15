"""CSV aus einem Export (KIS, SAP, Statistik) im Fenster unter „Datei“: als ein Blatt gelesen, wie Excel es öffnen
sollte, und auf seine Daten geprüft. Andere Formate (.ods, Numbers, .xls) bekommen einen Satz, was zu tun ist."""

import io
import random
import zipfile

from voiceagent import einlesen, pruefer

STATIONS = {"Kardiologie": "Dr. Keller", "Chirurgie": "Dr. Brunner", "Neurologie": "Dr. Steiner"}


def _rows(**planted) -> list[list[str]]:
    rnd = random.Random(3)
    rows = [["Fall-Nr", "Eintritt", "Alter", "Gewicht kg", "Station", "Arzt", "PLZ"]]
    for i in range(40):
        station = list(STATIONS)[i % 3]
        rows.append([str(240100 + i), f"{1 + i % 28:02d}.0{1 + i % 9}.2024", str(rnd.randint(20, 90)),
                     f"{rnd.randint(50, 110)},{rnd.randint(0, 9)}", station, STATIONS[station], f"0{rnd.randint(100, 999)}"])
    if planted.get("errors"):
        rows[21][0] = rows[6][0]          # Fall-Nr kopiert
        rows[9][2] = "212"                # Alter unmöglich
        rows[13][4] = "Kardiolgie"        # Tippfehler
    return rows


def _csv(rows, delimiter=";", mark=",", encoding="cp1252", bom=False) -> bytes:
    lines = []
    for row in rows:
        cells = [c.replace(",", ".") if mark == "." and c.replace(",", "").isdigit() else c for c in row]
        lines.append(delimiter.join(f'"{c}"' if delimiter in c else c for c in cells))
    text = "\r\n".join(lines) + "\r\n"
    return (b"\xef\xbb\xbf" if bom else b"") + text.encode(encoding)


def _keys(data: bytes, name="export.csv") -> dict:
    done = pruefer.Checker(None).check_file(name, data)
    assert done["state"] == "ok", done
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


def test_an_export_is_checked_for_its_data():
    found = _keys(_csv(_rows(errors=True)))
    assert found["id:A22"]["level"] == "fehler" and "„240120“ fehlt" in found["id:A22"]["text"]
    assert "212 in C10" in found["impossible:C10"]["text"]
    typo = found["typo:E14"]
    assert "„Kardiolgie“ ist wohl ein Tippfehler für „Kardiologie“" in typo["text"]
    # Postleitzahlen mit Null vorn bleiben Text und sind kein Fehler.
    assert not any(k.endswith(("G2", "G3")) or ":PLZ" in k for k in found)


def test_any_delimiter_decimal_mark_and_encoding_reads_the_same():
    semicolon = _keys(_csv(_rows(errors=True)))
    comma = _keys(_csv(_rows(errors=True), delimiter=",", mark=".", encoding="utf-8", bom=True), "export.txt")
    tab = _keys(_csv(_rows(errors=True), delimiter="\t", encoding="utf-8"), "export.tsv")
    assert sorted(semicolon) == sorted(comma) == sorted(tab)


def test_a_clean_export_has_nothing_to_say():
    assert _keys(_csv(_rows())) == {}


def test_numbers_dates_and_text_in_a_csv_are_what_they_mean():
    sheet = einlesen.csv_sheets("x.csv", "Betrag;Datum;Anteil;Code\n1.234,50;2024-03-05;12,5 %;0042\n".encode())[0]
    cells = sheet.cells
    assert (cells[(2, 1)].value, cells[(2, 2)].value, cells[(2, 3)].value) == ("1234,5", "05.03.2024", "12,5 %")
    assert cells[(2, 4)].value == "0042" and cells[(2, 4)].text == "0042" and sheet.name == "x"


def test_other_spreadsheet_formats_say_what_to_do():
    def zipped(**files):
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            for path, content in files.items():
                archive.writestr(path.replace("__", "/"), content)
        return buffer.getvalue()

    checker = pruefer.Checker(None)
    ods = checker.check_file("liste.ods", zipped(mimetype="application/vnd.oasis.opendocument.spreadsheet"))
    assert ods["state"] == "error" and "LibreOffice" in ods["text"] and "Speichere sie dort als Excel-Datei" in ods["text"]
    numbers = checker.check_file("liste.numbers", zipped(Index__Document_iwa="x"))
    assert "Numbers-Datei" in numbers["text"] and "Exportieren" in numbers["text"]
    binary = checker.check_file("kaputt.csv", b"\x00\x01\x02" * 10)
    assert binary["state"] == "error" and "keine Textdatei" in binary["text"]
    assert checker.check_file("leer.csv", b"  \n")["text"] == "Die Datei ist leer."
