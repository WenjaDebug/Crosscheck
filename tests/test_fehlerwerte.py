"""Fehlerwerte: wo sie entstehen, warum, und was Excel in der Übungsmappe wirklich zeigt (Test in Excel am 30.10.)."""

import re
import zipfile
from pathlib import Path

from voiceagent import formel, pruefer
from voiceagent.addin import _pivots
from voiceagent.analyse import Sheet as Checked
from voiceagent.analyse import analyse
from voiceagent.uebung import DEMO
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Sheet, read_xlsx

# Funktionen ab Excel 2010: In der Datei brauchen sie „_xlfn.“, sonst zeigt Excel #NAME?.
FUTURE = {"STDEV.S", "STDEV.P", "VAR.S", "VAR.P", "T.TEST", "T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T",
          "NORM.S.DIST", "NORM.S.INV", "CONFIDENCE.NORM", "CONFIDENCE.T", "PERCENTILE.INC", "QUARTILE.INC", "RANK.EQ",
          "IFNA", "IFS", "MAXIFS", "MINIFS", "SWITCH", "CONCAT", "TEXTJOIN", "XLOOKUP", "XMATCH", "FLOOR.MATH", "DAYS"}


def filled(cells, name="T", book=None):
    """Was Excel zeigt: jede Formel ausgerechnet, Fehler als Code."""
    for _ in range(3):
        calc = Checked(cells, layout(Sheet(name, cells)), name, book or {}).evaluator()[0]
        for pos, cell in list(cells.items()):
            if cell.formula and cell.value != "#NAME?":
                value = calc.run(formel.parse(cell.formula))
                text = value.code if isinstance(value, formel.Err) else repr(value) if isinstance(value, float) else str(value)
                cells[pos] = Cell(text, cell.formula)
    return cells


def errors(cells, name="T", book=None):
    found = analyse(cells, layout(Sheet(name, cells)), name=name, exact=True, book=book).findings
    return [f for f in found if f.key.startswith("error:")], found


def test_demo_stores_new_functions_the_way_excel_reads_them():
    with zipfile.ZipFile(DEMO) as archive:
        for name in archive.namelist():
            if not re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name):
                continue
            for body in re.findall(r"<f(?:\s[^>]*)?>([^<]*)</f>", archive.read(name).decode("utf-8")):
                for function in re.findall(r"(?<![A-Za-z0-9_.])([A-Z][A-Z0-9.]*)\(", body):
                    assert function not in FUTURE, (name, body)


def test_demo_stores_what_excel_computes():
    sheets = {s.name: s.cells for s in read_xlsx(DEMO.read_bytes())}
    rechnung = sheets["Rechnung"]
    assert [rechnung[(r, 5)].value for r in range(3, 9)] == ["#DIV/0!"] * 6
    for name, cells in sheets.items():
        checked = Checked(cells, layout(Sheet(name, cells)), name, sheets)
        calc, value = checked.evaluator(chain=set(checked.nodes))
        for pos in checked.nodes:
            if pos in checked.unknown:
                continue
            try:
                got = value(pos)
            except formel.Unknown:  # HEUTE() ändert sich jeden Tag.
                continue
            have = checked.inputs.get(pos)
            if isinstance(got, float) and isinstance(have, float):
                assert abs(got - have) <= 1e-6 * max(1.0, abs(got)), (name, pos)
            else:
                assert got == have, (name, pos, got, have)


def test_at_sign_before_a_function_changes_nothing():
    assert formel.parse("@STDEV.S(E2:E21)") == formel.parse("STDEV.S(E2:E21)")
    assert formel.parse("@INDEX(A1:A5,2)") == formel.parse("INDEX(A1:A5,2)")
    assert formel.parse("@A1:A10").kind == "unknown"


def excel_statistik():
    """Die Übungsmappe vor der Korrektur, so wie Excel sie dem Seitenfenster gibt: #NAME? und „@“."""
    sheets = read_xlsx(DEMO.read_bytes())
    stat = next(s for s in sheets if s.name == "Statistik").cells
    for row in range(3, 8):
        cell = stat[(row, 8)]
        formula = cell.formula.replace("_xlfn.", "").replace(",", ", ")
        unknown = any(f in formula for f in ("STDEV.S", "T.TEST", "CONFIDENCE"))
        stat[(row, 8)] = Cell("#NAME?", ("@" if unknown else "") + formula)
    return sheets


def test_name_errors_from_a_file_excel_did_not_read_are_re_entered():
    sheets = excel_statistik()
    done = pruefer.Checker(None).check_book("Übungsmappe.xlsx", "Statistik", sheets)
    hints = {h["key"]: h for h in done["hints"]}
    assert not any(k.startswith("constant:") for k in hints)
    [mine] = [h for h in hints.values() if h["key"].startswith("error:")]
    assert mine["cell"] == "H3:H6"
    assert "obwohl die Formeln richtig geschrieben sind" in mine["text"]
    assert "„@“" in mine["text"] and "auch in H7" in mine["text"]
    assert mine["fix"]["cells"] == {
        "H3": "=STDEV.S(E2:E21)", "H4": "=T.TEST(E2:E11, E12:E21, 2, 1)", "H5": "=T.TEST(C2:C21, D2:D21, 2, 3)",
        "H6": "=CONFIDENCE.NORM(0.95, H3, 20)"}
    assert mine["fix"]["was"]["H3"] == "@STDEV.S(E2:E21)"


def test_sliding_denominator_says_which_rows_show_div0():
    sheets = read_xlsx(DEMO.read_bytes())
    done = pruefer.Checker(None).check_book("Übungsmappe.xlsx", "Rechnung", sheets)
    keys = [h["key"] for h in done["hints"]]
    assert not [k for k in keys if k.startswith(("error:", "emptyref:"))]
    [slide] = [h for h in done["hints"] if h["key"] == "slide:E2"]
    assert "E3 bis E8 zeigen #DIV/0!: Die Formeln teilen durch D10, D11, D12, D13, D14 und 1 weitere, und die sind leer." \
        in slide["text"]
    assert "Die Korrektur behebt das mit." in slide["text"]


def test_empty_denominator_with_rows_below_and_followers():
    cells = {(1, 1): Cell("Menge"), (1, 2): Cell("Teiler"), (1, 3): Cell("Quote"), (1, 4): Cell("Doppelt")}
    for r in range(2, 6):
        cells[(r, 1)] = Cell(str(r * 10))
        cells[(r, 3)] = Cell("", f"A{r}/B{r}")
    cells[(2, 2)] = Cell("4")
    cells[(2, 4)] = Cell("", "C3*2")
    filled(cells)
    [one], _ = errors(cells)
    assert one.cell == "C3:C5"
    assert one.text.startswith("C3 bis C5 zeigen #DIV/0!: Die Formeln teilen durch B3, B4 und B5, und die sind leer.")
    assert one.tip.startswith("Trag dort die Werte ein")
    assert "auch in D2" in one.text


def test_division_by_zero_and_by_an_expression():
    cells = {(1, 1): Cell("5"), (1, 2): Cell("0"), (1, 3): Cell("", "A1/B1"), (2, 1): Cell("3"), (2, 2): Cell("3"),
             (2, 3): Cell("", "A1/(A2-B2)")}
    filled(cells)
    found, _ = errors(cells)
    texts = {f.key: f.text for f in found}
    assert "teilt durch B1, und B1 ist 0" in texts["error:C1"]
    assert "teilt durch A2-B2, und das ergibt 0" in texts["error:C2"]


def test_average_without_numbers_and_spread_of_one():
    cells = {(1, 1): Cell("Alter"), (2, 1): Cell("k. A."), (3, 1): Cell("12", text="12"),
             (1, 3): Cell("", "AVERAGE(A2:A3)"), (1, 4): Cell("7"), (1, 5): Cell("", "STDEV.S(D1:D1)"),
             (1, 6): Cell("", "AVERAGE(G1:G3)")}
    filled(cells)
    found, _ = errors(cells)
    texts = {f.key: f.text for f in found}
    assert "MITTELWERT(A2:A3) findet keine einzige Zahl" in texts["error:C1"]
    assert "STABW.S braucht mindestens zwei Zahlen, in D1:D1 steht nur eine" in texts["error:E1"]
    assert "G1:G3 ist leer" in texts["error:F1"]


def test_lookup_value_not_found_and_empty_key():
    cells = {(1, 1): Cell("Code"), (1, 2): Cell("Satz"), (2, 1): Cell("A"), (2, 2): Cell("1"), (3, 1): Cell("B"),
             (3, 2): Cell("2"), (5, 1): Cell("Z"), (5, 2): Cell("", "VLOOKUP(A5,$A$2:$B$3,2,FALSE)"),
             (6, 2): Cell("", "VLOOKUP(A6,$A$2:$B$3,2,FALSE)")}
    filled(cells)
    found, _ = errors(cells)
    texts = {f.key: f.text for f in found}
    assert "SVERWEIS findet „Z“ nicht in der ersten Spalte von A2:B3" in texts["error:B5"]
    assert "Der Suchwert A6 ist leer" in texts["error:B6"]


def test_text_in_arithmetic():
    cells = {(1, 1): Cell("k. A."), (1, 2): Cell("", "A1+1")}
    filled(cells)
    [one], _ = errors(cells)
    assert "rechnet mit A1, und dort steht Text („k. A.“)" in one.text


def test_name_typo_unquoted_text_and_missing_colon_get_a_verified_fix():
    cells = {(1, 1): Cell("Ost"), (2, 1): Cell("West"), (1, 2): Cell("5"), (2, 2): Cell("7"),
             (1, 4): Cell("#NAME?", "SUMM(B1:B2)"), (2, 4): Cell("#NAME?", "SUMIF(A1:A2,Ost,B1:B2)"),
             (3, 4): Cell("#NAME?", "SUM(B1B2)"), (4, 4): Cell("#NAME?", "VLOKUP(A1,A1:B2,2,FALSE)"),
             (5, 4): Cell("#NAME?", "SUMME(B1:B2)"), (6, 4): Cell("#NAME?", "GEHEIM(B1)")}
    found, _ = errors(cells)
    by = {f.key: f for f in found}
    assert "Excel kennt keine Funktion „SUMM“. Gemeint ist wohl SUMME" in by["error:D1"].text
    assert by["error:D1"].fix["cells"] == {"D1": "=SUM(B1:B2)"}
    assert "„Ost“ steht ohne Anführungszeichen" in by["error:D2"].text
    assert by["error:D2"].fix["cells"] == {"D2": '=SUMIF(A1:A2,"Ost",B1:B2)'}
    assert "Zwischen B1 und B2 fehlt der Doppelpunkt" in by["error:D3"].text
    assert by["error:D3"].fix["cells"] == {"D3": "=SUM(B1:B2)"}
    assert "Gemeint ist wohl SVERWEIS (englisch VLOOKUP)" in by["error:D4"].text
    assert "In der Datei steht SUMME mit dem deutschen Namen" in by["error:D5"].text and by["error:D5"].fix["cells"] == {"D5": "=SUM(B1:B2)"}
    assert "Excel kennt keine Funktion „GEHEIM“" in by["error:D6"].text and by["error:D6"].fix is None


def test_deleted_reference_typed_error_and_error_from_another_sheet():
    data = {(1, 1): Cell("#NV")}
    cells = {(1, 1): Cell("#REF!", "A2+#REF!"), (2, 1): Cell("1"), (3, 1): Cell("#N/A"), (4, 1): Cell("#N/A", "Daten!A1")}
    found, _ = errors(cells, book={"Daten": data})
    texts = {f.key: f.text for f in found}
    assert "statt einer Zelle" in texts["error:A1"] and "gelöscht" in texts["error:A1"]
    assert "als fester Wert" in texts["error:A3"]
    assert "liest Daten!A1, und dort steht schon #NV" in texts["error:A4"]


def test_one_hint_per_error_cell_in_the_pane():
    cells = {(1, 1): Cell("1"), (1, 2): Cell(""), (1, 3): Cell("#DIV/0!", "A1/B1")}
    sheet = Sheet("T", cells)
    done = pruefer.Checker(None).check_book("x.xlsx", "T", [sheet])
    assert [h["key"] for h in done["hints"] if "C1" in h["key"]] == ["error:C1"]


def test_pivot_sum_office_calls_automatic_is_read_from_the_caption():
    [pivot] = _pivots([{"name": "P", "area": "E3:F7", "body": "F4:F7", "rows": ["Region"], "columns": [], "filters": [],
                        "values": [{"caption": "Summe von Umsatz", "field": "Umsatz", "function": "Automatic",
                                    "showAs": "None"}],
                        "skipped": {}, "grand": [True, True], "filtered": False, "source": "T!$A$1:$C$14", "table": ""}])
    assert pivot.value_fields[0].function == "sum"
    [own] = _pivots([{"name": "P", "values": [{"caption": "Mein Umsatz", "field": "Umsatz", "function": "Automatic"}]}])
    assert own.value_fields[0].function == ""


def test_demo_path_is_the_practice_workbook():
    assert Path(DEMO).name == "pruefer-uebung.xlsx"


def test_number_errors_name_the_function_and_the_number():
    cells = {(1, 1): Cell("#NUM!", "ASIN(2)"), (2, 1): Cell("#NUM!", "CEILING(2,-2)"), (3, 1): Cell("#N/A", "NA()"),
             (4, 1): Cell("#NUM!", "DATEDIF(DATE(2024,5,1),DATE(2024,1,1),\"m\")"), (5, 1): Cell("#NUM!", "DATE(1900,1,-5)")}
    found, _ = errors(cells)
    texts = {f.key: f.text for f in found}
    assert "ARCSIN bekommt 2, rechnet aber nur mit Zahlen von -1 bis 1" in texts["error:A1"]
    assert "#ZAHL!" in texts["error:A1"] and "#NUM!" not in texts["error:A1"]
    assert "negatives Vielfaches (-2)" in texts["error:A2"]
    assert "mit NV() absichtlich #NV" in texts["error:A3"]
    # DATUM rechnet hier richtig, der Fehler kommt aus DATEDIF.
    assert "Enddatum" in texts["error:A4"] and "1.1.1900" not in texts["error:A4"]
    assert "vor dem 1.1.1900" in texts["error:A5"]


def test_a_number_as_text_owns_the_errors_it_causes():
    """6'480 als Text in G5: H5 und I5 zeigen #WERT!. Das sagt der Hinweis zur Zahl als Text, mit den Zellen."""
    from voiceagent import pruefer

    values, formulas = [["Name", "Lohn", "AHV", "Netto"]], [["Name", "Lohn", "AHV", "Netto"]]
    for r in range(2, 14):
        lohn = 5000 + 100 * r
        if r == 5:
            values.append([f"M{r}", "6'480", "#VALUE!", "#VALUE!"])
        else:
            values.append([f"M{r}", lohn, lohn * 0.05, lohn * 0.95])
        formulas.append([f"M{r}", "6'480" if r == 5 else lohn, f"=B{r}*0.05", f"=B{r}-C{r}"])
    done = pruefer.Checker(None).check_payload({"book": "L.xlsx", "active": "L", "sheets": [
        {"name": "L", "row": 1, "col": 1, "values": values, "formulas": formulas}]})
    keys = {h["key"]: h for h in done["hints"]}
    assert not [k for k in keys if k.startswith("error:")], list(keys)
    text = next(h["text"] for k, h in keys.items() if k.startswith("astext:"))
    assert "#WERT!" in text and "C5" in text and "D5" in text
