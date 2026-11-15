"""Ausgeblendete und weggefilterte Zeilen: SUMME zählt sie mit, TEILERGEBNIS(109; …) nicht. Der Prüfer liest sie aus der
Datei und aus dem Seitenfenster, rechnet TEILERGEBNIS wie Excel und sagt, was man sieht und was die Summe zählt."""

import io
import re
import zipfile

from test_workbook import xlsx

from voiceagent import einsetzen, formel, pruefer
from voiceagent.sichtbar import _rows
from voiceagent.xlsx import read_book


def _liste() -> dict:
    cells = {"A1": "Fall", "B1": "Station", "C1": "Kosten"}
    stations = ["Nord", "Süd", "Nord", "Ost", "Nord", "Süd", "Nord", "Ost", "Nord", "Süd", "Nord", "Ost"]
    for i, station in enumerate(stations):
        r = i + 2
        cells.update({f"A{r}": f"F{r:03d}", f"B{r}": station, f"C{r}": float(100 + 10 * i)})
    total = sum(100 + 10 * i for i in range(12))
    cells.update({"B15": "Summe", "C15": ("=", "SUM(C2:C13)", float(total))})
    return cells


def _hidden(data: bytes, rows: list[int], autofilter: str = "", grouped: bool = False) -> bytes:
    """Die Zeilen ausgeblendet; mit autofilter (z. B. A1:C13) als Filter auf „Station“."""
    source, out = zipfile.ZipFile(io.BytesIO(data)), io.BytesIO()
    with zipfile.ZipFile(out, "w") as book:
        for item in source.infolist():
            body = source.read(item)
            if item.filename == "xl/worksheets/sheet1.xml":
                text = body.decode()
                for r in rows:
                    extra = ' outlineLevel="1"' if grouped else ""
                    text = text.replace(f'<row r="{r}">', f'<row r="{r}" hidden="1"{extra}>')
                if autofilter:
                    text = text.replace("</sheetData>", f'</sheetData><autoFilter ref="{autofilter}"><filterColumn '
                                        'colId="1"><filters><filter val="Nord"/></filters></filterColumn></autoFilter>')
                body = text.encode()
            book.writestr(item, body)
    return out.getvalue()


def _hints(data: bytes) -> dict:
    done = pruefer.Checker(None).check_file("liste.xlsx", data)
    assert done["state"] == "ok", done
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


NOT_NORD = [3, 5, 7, 9, 11, 13]


def test_hidden_rows_are_read_and_subtotal_leaves_them_out():
    data = _hidden(xlsx(("Liste", _liste())), NOT_NORD, "A1:C13")
    cells = read_book(data).sheets[0].cells
    assert cells.hidden_rows == frozenset(NOT_NORD) and cells.filtered_rows == frozenset(NOT_NORD)
    calc = formel.Evaluator(lambda p: cells[p].value and float(cells[p].value) if p in cells and p[1] == 3 and p[0] > 1 else None)
    calc.nested, calc.hidden, calc.filtered = (lambda p: False), cells.hidden_rows, cells.filtered_rows
    nord = sum(100 + 10 * i for i in range(0, 12, 2))
    assert calc.run(formel.parse("SUBTOTAL(109,C2:C13)")) == calc.run(formel.parse("SUBTOTAL(9,C2:C13)")) == nord
    # Von Hand ausgeblendet: TEILERGEBNIS(9; …) zählt sie, TEILERGEBNIS(109; …) nicht.
    calc.filtered = frozenset()
    assert calc.run(formel.parse("SUBTOTAL(9,C2:C13)")) == sum(100 + 10 * i for i in range(12))


def test_a_sum_below_a_filtered_list_says_what_the_filter_shows():
    found = _hints(_hidden(xlsx(("Liste", _liste())), NOT_NORD, "A1:C13"))
    hint = found["filter:C15"]
    nord = sum(100 + 10 * i for i in range(0, 12, 2))
    assert hint["level"] == "pruefen"
    assert hint["text"] == (f"C15 (Kosten) zählt auch die 6 weggefilterten Zeilen (3, 5, 7 und 3 weitere) mit: Was der Filter "
                            f"zeigt, ergibt {nord}, C15 zeigt 1860.")
    assert hint["fix"]["cells"] == {"C15": "=SUBTOTAL(109,C2:C13)"}
    assert hint["tip"].startswith(f"Soll C15 nur zählen, was zu sehen ist: Korrigiert in C15: =TEILERGEBNIS(109;C2:C13), "
                                  f"nachgerechnet {nord}.")
    after = einsetzen.applied({}, [hint["fix"]])
    assert after[(15, 3)].formula == "SUBTOTAL(109,C2:C13)"


def test_rows_hidden_by_hand_are_a_note_and_a_collapsed_group_is_quiet():
    found = _hints(_hidden(xlsx(("Liste", _liste())), [8, 9, 10]))
    hint = found["ausgeblendet:C15"]
    assert hint["level"] == "info" and hint["text"].startswith("C15 (Kosten) zählt auch die ausgeblendeten Zeilen 8 bis 10 mit:")
    assert not any(k.startswith(("filter:", "ausgeblendet:")) for k in _hints(_hidden(xlsx(("Liste", _liste())), [8, 9, 10],
                                                                                   grouped=True)))


def test_quiet_when_nothing_changes_or_only_hidden_rows_are_summed():
    cells = _liste()
    cells.update({"C8": 0.0, "C9": 0.0, "B16": "Teil", "C16": ("=", "SUM(C8:C9)", 0.0)})
    cells["C15"] = ("=", "SUM(C2:C13)", 1860.0 - 160 - 170)
    found = _hints(_hidden(xlsx(("Liste", cells)), [8, 9]))
    assert not any(k.startswith(("filter:", "ausgeblendet:")) for k in found)
    # Eine Summe nur über die ausgeblendeten Zeilen (das Ergebnis einer eingeklappten Gruppe) ist so gemeint.
    cells = _liste()
    cells.update({"B16": "Teil", "C16": ("=", "SUM(C8:C10)", 160.0 + 170 + 180)})
    assert "ausgeblendet:C16" not in _hints(_hidden(xlsx(("Liste", cells)), [8, 9, 10]))


def test_totals_side_by_side_are_one_hint():
    cells = _liste()
    for r in range(2, 14):
        cells[f"D{r}"] = float(r)
    cells.update({"D1": "Tage", "D15": ("=", "SUM(D2:D13)", float(sum(range(2, 14))))})
    found = _hints(_hidden(xlsx(("Liste", cells)), NOT_NORD, "A1:D13"))
    hint = found["filter:C15"]
    assert "D15" not in "".join(k for k in found if k != "filter:C15") and hint["text"].endswith(" Ebenso D15.")
    assert hint["fix"]["cells"] == {"C15": "=SUBTOTAL(109,C2:C13)", "D15": "=SUBTOTAL(109,D2:D13)"}


def test_the_pane_sends_hidden_rows():
    values = [["Fall", "Kosten"], ["F1", 10], ["F2", 20], ["F3", 30], ["Summe", 60]]
    formulas = [["Fall", "Kosten"], ["F1", 10], ["F2", 20], ["F3", 30], ["Summe", "=SUM(B2:B4)"]]
    done = pruefer.Checker(None).check_payload({"book": "P.xlsx", "active": "Liste", "sheets": [
        {"name": "Liste", "row": 1, "col": 1, "values": values, "formulas": formulas, "hiddenRows": [3],
         "filtered": ["Liste!A1:B4"]}]})
    hint = next(h for h in done["hints"] if h["key"] == "filter:B5")
    assert "die weggefilterte Zeile 3 mit: Was der Filter zeigt, ergibt 40, B5 zeigt 60." in hint["text"]
    # Ohne die Angabe (ein älteres Fenster) bleibt es beim Hinweis, dass die Tabelle gefiltert ist.
    done = pruefer.Checker(None).check_payload({"book": "P.xlsx", "active": "Liste", "sheets": [
        {"name": "Liste", "row": 1, "col": 1, "values": values, "formulas": formulas, "filtered": ["Liste!A1:B4"]}]})
    assert "ergibt" not in next(h for h in done["hints"] if h["key"] == "filter:B5")["text"]


def test_rows_in_words():
    assert _rows([8]) == "8" and _rows([8, 9]) == "8 und 9" and _rows([8, 9, 10]) == "8 bis 10"
    assert _rows([3, 5, 7]) == "3, 5 und 7" and _rows([3, 5, 7, 9, 11]) == "3, 5, 7 und 2 weitere"
    assert _rows([2, 3, 4, 8, 9]) == "2 bis 4, 8 und 9"
    assert re.fullmatch(r"\d+ bis \d+", _rows(list(range(2, 40))))


def test_sum_average_and_count_below_the_list_are_one_hint_and_all_stay():
    cells = _liste()
    values = [100 + 10 * i for i in range(12)]
    cells.update({"B16": "Mittel", "C16": ("=", "AVERAGE(C2:C13)", sum(values) / 12),
                  "B17": "Anzahl", "C17": ("=", "COUNT(C2:C13)", 12.0), "B18": "Größter", "C18": ("=", "MAX(C2:C13)", 210.0)})
    hint = _hints(_hidden(xlsx(("Liste", cells)), NOT_NORD, "A1:C13"))["filter:C15"]
    # Der größte Wert (210 in Zeile 13) ist weggefiltert, darum ändert sich auch MAX.
    assert hint["text"].endswith(" Ebenso C16, C17 und C18.")
    assert hint["fix"]["cells"] == {"C15": "=SUBTOTAL(109,C2:C13)", "C16": "=SUBTOTAL(101,C2:C13)",
                                    "C17": "=SUBTOTAL(102,C2:C13)", "C18": "=SUBTOTAL(104,C2:C13)"}
