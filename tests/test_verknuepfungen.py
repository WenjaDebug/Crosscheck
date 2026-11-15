"""Werte aus einer anderen Datei: =[Budget.xlsx]Plan!B3 holt, was dort zuletzt stand. Ein Hinweis je Blatt nennt die
Zellen und die Dateien. Gibt es das Blatt, auf das gezeigt wird, auch in dieser Mappe (beim Kopieren mitgekommen),
sagt der Hinweis, was die Formel hier rechnen würde, und Einsetzen schreibt den Bezug auf das eigene Blatt."""

import io
import zipfile

from test_workbook import xlsx

from voiceagent import einsetzen, pruefer
from voiceagent.verknuepfungen import links_in

_NS = "http://schemas.openxmlformats.org"


def _with_link(data: bytes, target: str) -> bytes:
    """Die Mappe mit einer Verknüpfung [1] auf target, wie Excel sie speichert."""
    source, out = zipfile.ZipFile(io.BytesIO(data)), io.BytesIO()
    with zipfile.ZipFile(out, "w") as book:
        for item in source.infolist():
            body = source.read(item)
            if item.filename == "xl/workbook.xml":
                body = body.replace(b"</sheets>", b'</sheets><externalReferences><externalReference r:id="rId99"/>'
                                                  b"</externalReferences>")
            if item.filename == "xl/_rels/workbook.xml.rels":
                body = body.replace(b"</Relationships>", f'<Relationship Id="rId99" Type="{_NS}/officeDocument/2006/'
                                    'relationships/externalLink" Target="externalLinks/externalLink1.xml"/>'
                                    "</Relationships>".encode())
            book.writestr(item, body)
        book.writestr("xl/externalLinks/externalLink1.xml",
                      f'<externalLink xmlns="{_NS}/spreadsheetml/2006/main"><externalBook xmlns:r="{_NS}/'
                      'officeDocument/2006/relationships" r:id="rId1"/></externalLink>')
        book.writestr("xl/externalLinks/_rels/externalLink1.xml.rels",
                      f'<Relationships xmlns="{_NS}/package/2006/relationships"><Relationship Id="rId1" '
                      f'Type="{_NS}/officeDocument/2006/relationships/externalLinkPath" Target="{target}" '
                      'TargetMode="External"/></Relationships>')
    return out.getvalue()


PLAN = {"A1": "Größe", "B1": "Wert", "A2": "Satz", "B2": 0.081, "A3": "Basis", "B3": 1000.0}


def _bericht() -> dict:
    return {"A1": "Posten", "B1": "Betrag", "A2": "Umsatz", "B2": 5000.0, "A3": "MwSt", "B3": ("=", "B2*[1]Plan!$B$2", 385.0),
            "A4": "Basis", "B4": ("=", "[1]Plan!B3", 1200.0), "A5": "Kurs", "B5": ("=", "[1]Kurse!B2", 0.94)}


def _hints(*sheets) -> dict:
    data = _with_link(xlsx(*sheets), "file:///Users/x/Budget%202025.xlsx")
    done = pruefer.Checker(None).check_file("bericht.xlsx", data)
    assert done["state"] == "ok", done
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


def test_a_copied_sheet_still_pointing_at_the_old_file_gets_its_own_sheet():
    found = _hints(("Bericht", _bericht()), ("Plan", PLAN))
    hint = found["copylink:B3"]
    assert hint["level"] == "pruefen" and hint["fix"]["cells"] == {"B3": "=B2*Plan!$B$2", "B4": "=Plan!B3"}
    assert "ein Blatt „Plan“ gibt es auch in dieser Mappe" in hint["text"]
    assert "Mit dem Blatt dieser Mappe ergäbe B3 405 statt 385. Ebenso in B4." in hint["text"]
    assert hint["tip"] == "Korrigiert in B3: =B2*Plan!$B$2, nachgerechnet 405."
    # Die Kurse gibt es hier nicht: Dafür bleibt der Hinweis auf die andere Datei, mit ihrem Namen.
    other = found["extern:B5"]
    assert other["level"] == "info" and other["text"].startswith("B5 holt Werte aus einer anderen Datei: „Budget 2025.xlsx“.")
    # Umsatz über MwSt und Basis ist eine Eingabe, keine vergessene Formel.
    assert not any(k.startswith("typed:") for k in found)
    after = einsetzen.applied({}, [hint["fix"]])
    assert after[(3, 2)].formula == "B2*Plan!$B$2"


def test_links_to_a_file_without_such_a_sheet_are_one_hint():
    found = _hints(("Bericht", _bericht()))
    assert [k for k in found if k.startswith(("extern:", "copylink:"))] == ["extern:B3"]
    hint = found["extern:B3"]
    assert hint["text"].startswith("B3, B4 und B5 holen Werte aus einer anderen Datei: „Budget 2025.xlsx“.")
    assert "Arbeitsmappenlinks" in hint["tip"]


def test_every_notation_of_a_link_and_nothing_else():
    assert links_in("B2*[1]Plan!$B$2") == [("[1]Plan!", "1", "Plan")]
    assert links_in("'C:\\Daten\\[Budget.xlsx]Plan 2025'!B3") == [("'C:\\Daten\\[Budget.xlsx]Plan 2025'!", "Budget.xlsx",
                                                                     "Plan 2025")]
    assert links_in("'/Users/x/[Budget 2025.xlsx]Bob''s Plan'!$B$2+1") == [
        ("'/Users/x/[Budget 2025.xlsx]Bob''s Plan'!", "Budget 2025.xlsx", "Bob's Plan")]
    assert links_in("[1]!Zins*2") == [("[1]!", "1", "")]
    assert links_in('"[1]Plan!B3"&A1') == []
    assert links_in("SUM(Tabelle1[@Menge])*Tabelle1[[#Totals],[Preis]]") == []


def test_the_pane_shows_the_path_and_einsetzen_writes_the_own_sheet():
    values = [["Posten", "Betrag"], ["Umsatz", 5000], ["MwSt", 385]]
    formulas = [["Posten", "Betrag"], ["Umsatz", 5000], ["MwSt", "=B2*'/Users/x/[Budget 2025.xlsx]Plan'!$B$2"]]
    plan = [["Größe", "Wert"], ["Satz", 0.081]]
    done = pruefer.Checker(None).check_payload({"book": "Bericht.xlsx", "active": "Bericht", "sheets": [
        {"name": "Bericht", "row": 1, "col": 1, "values": values, "formulas": formulas},
        {"name": "Plan", "row": 1, "col": 1, "values": plan, "formulas": plan}]})
    hint = next(h for h in done["hints"] if h["key"] == "copylink:B3")
    assert hint["fix"]["cells"] == {"B3": "=B2*Plan!$B$2"} and "ergäbe B3 405 statt 385" in hint["text"]
    assert not any(h["key"].startswith("extern:") for h in done["hints"])
