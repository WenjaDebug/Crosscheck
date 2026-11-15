"""Ausgeblendete Blätter: Formeln anderer Blätter rechnen damit, darum liest der Prüfer sie mit, prüft sie und sagt
dazu, dass sie ausgeblendet sind. Ein Klick auf ein solches Blatt im Seitenfenster erklärt, wie man es einblendet."""

import io
import zipfile

from test_workbook import xlsx

from voiceagent import pruefer
from voiceagent.xlsx import read_book


def _hide(data: bytes, name: str, state: str = "hidden") -> bytes:
    source, out = zipfile.ZipFile(io.BytesIO(data)), io.BytesIO()
    with zipfile.ZipFile(out, "w") as book:
        for item in source.infolist():
            body = source.read(item)
            if item.filename == "xl/workbook.xml":
                body = body.replace(f'<sheet name="{name}"'.encode(), f'<sheet name="{name}" state="{state}"'.encode())
            book.writestr(item, body)
    return out.getvalue()


def _liste() -> dict:
    cells = {"A1": "Artikel", "B1": "Netto", "C1": "MwSt", "D1": "Brutto"}
    for r in range(2, 14):
        net = float(100 + r * 7)
        cells.update({f"A{r}": f"Art {r}", f"B{r}": net, f"C{r}": ("=", f"B{r}*Parameter!$B$2", net * 0.19),
                      f"D{r}": ("=", f"B{r}+C{r}", net * 1.19)})
    cells["C8"] = ("=", "B8*Parameter!$B$3", 156 * 0.05)
    cells["D8"] = ("=", "B8+C8", 156 * 1.05)
    return cells


PARAMETER = {"A1": "Größe", "B1": "Wert", "A2": "MwSt", "B2": 0.19, "A3": "Rabatt", "B3": 0.05}


def test_formulas_on_a_hidden_sheet_are_recomputed_and_fixed():
    for state in ("hidden", "veryHidden"):
        data = _hide(xlsx(("Liste", _liste()), ("Parameter", PARAMETER)), "Parameter", state)
        assert [(s.name, s.hidden) for s in read_book(data).sheets] == [("Liste", False), ("Parameter", True)]
        done = pruefer.Checker(None).check_file("preise.xlsx", data)
        hint = next(h for s in done["sheets"] for h in s["hints"] if h["key"] == "pattern:C8")
        assert hint["fix"]["cells"] == {"C8": "=B8*Parameter!$B$2"} and "nachgerechnet 29,64" in hint["tip"]
        assert [(s["name"], s.get("hidden", False)) for s in done["sheets"]] == [("Liste", False), ("Parameter", True)]


def test_a_hidden_sheet_gets_its_own_hints():
    parameter = {**PARAMETER, "A4": "Summe", "B4": ("=", "SUM(B2:B2)", 0.19)}
    data = _hide(xlsx(("Liste", _liste()), ("Parameter", parameter)), "Parameter")
    done = pruefer.Checker(None).check_file("preise.xlsx", data)
    hidden = next(s for s in done["sheets"] if s["name"] == "Parameter")
    assert hidden["hidden"] and hidden["hints"]


def test_visible_sheets_get_their_cells_first():
    """Reichen die Zellen nicht für alle Blätter, bekommt ein großes ausgeblendetes Blatt nur den Rest."""
    from voiceagent import xlsx as reader

    big = {f"A{r}": float(r) for r in range(1, 40)}
    data = _hide(xlsx(("Daten", big), ("Liste", _liste())), "Daten", "veryHidden")
    kept, reader.MAX_BOOK_CELLS = reader.MAX_BOOK_CELLS, 60
    try:
        sheets = {s.name: s for s in read_book(data).sheets}
    finally:
        reader.MAX_BOOK_CELLS = kept
    assert len(sheets["Liste"].cells) == len(_liste()) and len(sheets["Daten"].cells) <= 60 - len(_liste())
    assert list(sheets) == ["Daten", "Liste"]


def test_the_pane_marks_hidden_sheets_among_the_others():
    liste = [["Artikel", "Netto"], ["A", 1], ["B", 2]]
    parameter = [["Größe", "Wert"], ["MwSt", 0.19], ["Summe", 0.19]]
    formulas = [["Größe", "Wert"], ["MwSt", 0.19], ["Summe", "=SUM(B2:B2)"]]
    done = pruefer.Checker(None).check_payload({"book": "P.xlsx", "active": "Liste", "sheets": [
        {"name": "Liste", "row": 1, "col": 1, "values": liste, "formulas": liste},
        {"name": "Parameter", "row": 1, "col": 1, "values": parameter, "formulas": formulas, "hidden": True}]})
    other = next(o for o in done["others"] if o["name"] == "Parameter")
    assert other["hidden"] is True
