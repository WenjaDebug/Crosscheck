"""Korrekturen zum Einsetzen: aus dem deutschen Tipp die Formel, die das Seitenfenster mit einem Klick in Excel schreibt."""

import pytest
from test_excel_live import FakeExcel, client_with
from test_statistik import AGES, filled, found, shares, table

from voiceagent import einsetzen, formel
from voiceagent.analyse import Finding
from voiceagent.analyse import Sheet as Checked
from voiceagent.bereiche import _german_formula
from voiceagent.workbook import FUNCTIONS, layout
from voiceagent.xlsx import Cell, Sheet


@pytest.mark.parametrize("german, english", [
    ('SUMMEWENNS(D:D;B:B;"Nord";C:C;"Ost; West")', 'SUMIFS(D:D,B:B,"Nord",C:C,"Ost; West")'),
    ("RUNDEN(A1*0,5;2)", "ROUND(A1*0.5,2)"),
    ("SVERWEIS(A2;Tarife!$F$2:$H$5;3;FALSCH)", "VLOOKUP(A2,Tarife!$F$2:$H$5,3,FALSE)"),
    ('WENNFEHLER(B2/C2;"")', 'IFERROR(B2/C2,"")'),
    ("MEDIAN(FILTER(B2:B13;B2:B13<>999))", "MEDIAN(FILTER(B2:B13,B2:B13<>999))"),
    ('ANZAHL(B2:B13)-ZÄHLENWENN(B2:B13;"WAHR")', 'COUNT(B2:B13)-COUNTIF(B2:B13,"WAHR")'),
])
def test_german_formulas_become_what_excel_takes_from_office_js(german, english):
    assert einsetzen.english(german) == english
    assert _german_formula(english) == german


def test_every_function_name_goes_there_and_back():
    for name, german in FUNCTIONS.items():
        assert einsetzen.english(f"{german}(A1;2,5)") == f"{name}(A1,2.5)"


def applied(cells, fix):
    """Die Korrektur so einsetzen, wie das Seitenfenster es tut, und neu rechnen."""
    cells = dict(cells)
    for where, formula in fix["cells"].items():
        cells[formel.position(where)] = Cell("0", formula[1:])
    return filled(cells)


def test_one_cell_with_what_stood_there_before():
    cells = table(AGES, [("Mittelwert Alter", "AVERAGE(B2:B13)")])
    hint = found(cells)["statcode:E2"]
    assert hint.fix == {"range": "E2", "cells": {"E2": '=AVERAGEIF(B2:B13,"<>999")'}, "was": {"E2": "AVERAGE(B2:B13)"},
                        "german": '=MITTELWERTWENN(B2:B13;"<>999")'}
    # Eingesetzt ist der Fehler weg, und keine andere Prüfung hält die Korrektur für falsch.
    after = found(applied(cells, hint.fix))
    assert not [k for k, f in after.items() if k.endswith(":E2") and f.level == "fehler"]


def test_a_sum_that_leaves_out_the_codes_is_right():
    """SUMMEWENN(…;"<>999") unter „Summe Alter“ nimmt genau die echten Werte: kein Hinweis auf eine Bedingung."""
    after = found(table(AGES, [("Summe Alter", 'SUMIF(B2:B13,"<>999")'), ("Anzahl Alter", 'COUNTIF(B2:B13,"<>999")')]))
    assert not [k for k in after if k.endswith((":E2", ":E3"))]


def test_a_column_of_shares_is_written_as_one_block():
    cells = shares("B{r}/$F$2")
    hint = found(cells)["base:C2"]
    assert hint.fix["range"] == "C2:C4" and hint.fix["german"] == "=B2/$B$5"
    assert hint.fix["cells"] == {"C2": "=B2/$B$5", "C3": "=B3/$B$5", "C4": "=B4/$B$5"}
    assert hint.fix["was"] == {"C2": "B2/$F$2", "C3": "B3/$F$2", "C4": "B4/$F$2"}
    assert "base:C2" not in found(applied(cells, hint.fix))


def test_filled_down_to_the_end_of_the_column():
    """„… dann bis unten herunterziehen“: so weit, wie die Formel heute heruntergezogen ist."""
    cells = {(1, 1): Cell("Fall"), (1, 2): Cell("Punkte"), (1, 3): Cell("Doppelt")}
    for r in range(2, 8):
        cells[(r, 1)], cells[(r, 2)], cells[(r, 3)] = Cell(f"F{r}"), Cell(str(r)), Cell("0", f"B{r}*2")
    sheet = Checked(filled(cells), layout(Sheet("T", cells)), "T", {})
    tip = "Korrigiert: =B2*3, dann bis unten herunterziehen."
    fix = einsetzen.fix(sheet, Finding("pattern:C2", "", "fehler", tip))
    assert fix["range"] == "C2:C7" and fix["cells"]["C7"] == "=B7*3" and len(fix["cells"]) == 6


@pytest.mark.parametrize("tip", [
    # Der Tipp nennt einen anderen Wert, als die Formel ergibt.
    "Korrigiert: =SUMME(B2:B3), nachgerechnet 99.",
    # Eine Funktion, die der eigene Formelrechner nicht kennt.
    "Korrigiert: =GIBTESNICHT(B2:B3).",
    # Ein Fehlerwert, wo der Tipp keinen nennt.
    "Korrigiert: =B2/0.",
    # Nach oben herunterziehen gibt es nicht.
    "Korrigiert in C5: =B5, dann bis C2 herunterziehen.",
    "Prüf die Zeilen selbst.",
])
def test_nothing_to_insert_when_it_is_not_sure(tip):
    cells = {(1, 1): Cell("Fall"), (1, 2): Cell("Punkte"), (2, 2): Cell("3"), (3, 2): Cell("4")}
    sheet = Checked(cells, layout(Sheet("T", cells)), "T", {})
    assert einsetzen.fix(sheet, Finding("x:C2", "", "fehler", tip)) is None


def test_sum_plus_the_missing_cell_counts_that_cell():
    """=SUMME(D2:D12)+D13 unter „Umsatz Nord“ nimmt alle Nord-Zeilen: kein Hinweis, die Zeile fehle."""
    cells = {(1, 1): Cell("Region"), (1, 4): Cell("Umsatz"), (1, 8): Cell("Umsatz Nord")}
    for r in range(2, 21):
        cells[(r, 1)] = Cell("Nord" if r <= 13 else "Ost")
        cells[(r, 4)] = Cell(str(100 + r * 7))
    cells[(2, 8)] = Cell("0", "SUM(D2:D12)+D13")
    assert not [k for k in found(filled(dict(cells))) if k.endswith(":H2")]
    cells[(2, 8)] = Cell("0", "SUM(D2:D12)")
    assert "group:H2" in found(filled(dict(cells))) or "label:H2" in found(filled(dict(cells)))


def test_the_side_pane_gets_the_fix(tmp_path):
    values = [["Fall", "Alter", "", "", "Mittelwert Alter"]] + [[f"F{i}", a, "", "", ""] for i, a in enumerate(AGES)]
    formulas = [[None] * 5 for _ in values]
    values[1][4], formulas[1][4] = sum(AGES) / len(AGES), "=AVERAGE(B2:B13)"
    sheet = {"name": "T", "row": 1, "col": 1, "values": values, "formulas": formulas}
    _, client = client_with(tmp_path, FakeExcel(None))
    result = client.post("/api/addin/pruefen", json={"book": "A.xlsx", "active": "T", "sheets": [sheet]},
                         headers={"X-VoiceAgent": "1"}).json()
    hint = next(h for h in result["hints"] if h["key"] == "statcode:E2")
    # Was vorher dastand, so wie das Seitenfenster es aus Excel liest (ohne „=“): Daran sieht es, ob sich die Zelle
    # seit dem Prüfen geändert hat.
    assert hint["fix"]["was"] == {"E2": "AVERAGE(B2:B13)"} and hint["fix"]["cells"] == {"E2": '=AVERAGEIF(B2:B13,"<>999")'}


def test_shares_whose_denominator_slides_get_it_fixed_with_a_dollar():
    cells = {(1, 3): Cell("Station"), (1, 4): Cell("Fälle"), (1, 5): Cell("Anteil")}
    for r in range(2, 9):
        cells[(r, 3)], cells[(r, 4)], cells[(r, 5)] = Cell(f"S{r}"), Cell(str(10 * r)), Cell("0", f"D{r}/D{r + 7}")
    cells[(9, 3)], cells[(9, 4)] = Cell("Summe"), Cell("0", "SUM(D2:D8)")
    cells = filled(cells)
    hint = found(cells)["slide:E2"]
    assert hint.tip == ("Korrigiert in E2: =D2/D$9, dann bis E8 herunterziehen. Nachgerechnet teilt dann jede Zeile "
                        "durch die Summe.")
    assert hint.fix["range"] == "E2:E8" and hint.fix["cells"]["E8"] == "=D8/D$9" and hint.fix["was"]["E8"] == "D8/D15"
    assert "slide:E2" not in found(applied(cells, hint.fix))


@pytest.mark.parametrize("said, ok", [("50 %", True), ("0,5", True), ("60 %", False)])
def test_a_value_said_in_percent(said, ok):
    cells = {(1, 1): Cell("Fälle"), (1, 2): Cell("Davon"), (2, 1): Cell("2"), (3, 1): Cell("2"), (2, 2): Cell("1"),
             (3, 2): Cell("1")}
    sheet = Checked(cells, layout(Sheet("T", cells)), "T", {})
    tip = f"Gemeint ist meist die Quote über alle Zeilen: =SUMME(B2:B3)/SUMME(A2:A3), nachgerechnet {said}. Soll …"
    fix = einsetzen.fix(sheet, Finding("ratio:D4", "", "pruefen", tip))
    assert (fix == {"range": "D4", "cells": {"D4": "=SUM(B2:B3)/SUM(A2:A3)"}, "was": {"D4": ""},
                    "german": "=SUMME(B2:B3)/SUMME(A2:A3)"}) is ok


def test_one_cell_that_left_its_column_gets_the_column_formula_back():
    cells = {(1, 2): Cell("Region"), (1, 4): Cell("Umsatz"), (1, 5): Cell("Laufend")}
    for r in range(2, 8):
        cells[(r, 2)], cells[(r, 4)], cells[(r, 5)] = Cell("Nord"), Cell(str(r)), Cell("0", f"SUM(D$2:D{r})")
    cells[(4, 5)] = Cell("0", "SUM(D$2:D3)")
    sheet = Checked(filled(cells), layout(Sheet("T", cells)), "T", {})
    fix = einsetzen.fix(sheet, Finding("pattern:E4", "", "fehler", "Korrigiert: =SUMME(D$2:D4), nachgerechnet 9."))
    assert fix["cells"] == {"E4": "=SUM(D$2:D4)"} and fix["was"] == {"E4": "SUM(D$2:D3)"}


def test_a_running_total_is_not_a_group_that_ends_too_early():
    """=SUMME(D$2:D15) heruntergezogen: dass Zeile 15 kurz vor dem Ende von „Ost“ liegt, ist kein Fehler."""
    cells = {(1, 2): Cell("Region"), (1, 4): Cell("Umsatz"), (1, 5): Cell("Laufend")}
    for r in range(2, 21):
        cells[(r, 2)] = Cell("Ost" if r <= 17 else "West")
        cells[(r, 4)], cells[(r, 5)] = Cell(str(50 + r)), Cell("0", f"SUM(D$2:D{r})")
    assert not [k for k in found(filled(cells)) if k.startswith("group:")]


# --- Mehr einsetzen: Korrekturen, die der Tipp nur in Worten sagt, und alle zusammen ---


def months(figures, text=None):
    """Monate in A:C (Umsatz, Kosten), Kennzahlen in E:F. text: Zeile → Umsatz als Text."""
    cells = {(1, 1): Cell("Monat"), (1, 2): Cell("Umsatz"), (1, 3): Cell("Kosten"), (1, 5): Cell("Kennzahl"), (1, 6): Cell("Wert")}
    for r in range(2, 14):
        cells[(r, 1)], cells[(r, 2)], cells[(r, 3)] = Cell(f"M{r}"), Cell(str(100 + r * 7)), Cell(str(50 + r))
    for r, raw in (text or {}).items():
        cells[(r, 2)] = Cell(raw, None, raw)
    for r, (name, f) in enumerate(figures, start=2):
        cells[(r, 5)], cells[(r, 6)] = Cell(name), Cell("0", f)
    return filled(cells)


def test_a_range_that_stops_too_early_is_pulled_to_the_last_row_with_its_dollars():
    cells = months([("Mittel Umsatz", "AVERAGE($B$2:$B$11)"), ("Median Kosten", "MEDIAN(C2:C11)")])
    hints = found(cells)
    assert hints["range:F2"].fix == {"range": "F2", "cells": {"F2": "=AVERAGE($B$2:$B$13)"}, "was": {"F2": "AVERAGE($B$2:$B$11)"},
                                     "german": "=MITTELWERT($B$2:$B$13)"}
    assert hints["range:F3"].fix["cells"] == {"F3": "=MEDIAN(C2:C13)"}
    after = found(applied(cells, hints["range:F2"].fix))
    assert "range:F2" not in after and "range:F3" in after


def test_ranges_next_to_it_in_the_same_call_are_pulled_too():
    cells = months([])
    sheet = Checked(cells, layout(Sheet("T", cells)), "T", {})
    cells[(2, 6)] = Cell("0", 'SUMIF($A$2:$A$11,"M5",B2:B11)')
    sheet = Checked(cells, layout(Sheet("T", cells)), "T", {})
    fix = einsetzen.stretched(sheet, [(2, 6)], "SUMIF", (2, 2, 11, 2), (2, 13))
    assert fix["cells"] == {"F2": '=SUMIF($A$2:$A$13,"M5",B2:B13)'}


def test_a_range_is_not_pulled_over_a_subtotal_or_onto_itself():
    from voiceagent.auswertung import _longer

    cells = months([])
    cells[(12, 2)] = Cell("0", "SUM(B2:B11)")
    sheet = Checked(cells, layout(Sheet("T", cells)), "T", {})
    cells[(2, 6)] = Cell("0", "AVERAGE(B2:B11)")
    sheet = Checked(cells, layout(Sheet("T", cells)), "T", {})
    rows = set(range(2, 12)) | {13}
    assert _longer(sheet, (2, 6), sheet.nodes[(2, 6)], (2, 2, 11, 2), (2, 13), rows, set()) is None
    assert _longer(sheet, (2, 6), sheet.nodes[(2, 6)], (2, 2, 11, 2), (2, 13), rows | {12}, set()) is not None
    # Nimmt der Aufruf B13 schon dazu (MITTELWERT(B2:B11;B13)), zählte es doppelt.
    assert _longer(sheet, (2, 6), sheet.nodes[(2, 6)], (2, 2, 11, 2), (2, 13), rows | {12}, {13}) is None
    # Die Formel stünde in ihrem eigenen Bereich: ein Kreisbezug.
    assert _longer(sheet, (12, 2), sheet.nodes[(12, 2)], (2, 2, 11, 2), (2, 13), rows | {12}, set()) is None


def test_numbers_as_text_are_written_as_numbers_in_the_format_of_the_column():
    cells = months([("Mittel Umsatz", "AVERAGE(B2:B13)")], text={6: "142", 8: "4855.45"})
    fix = found(cells)["astext:B6"].fix
    assert fix == {"range": "B6 und B8", "cells": {"B6": 142.0, "B8": 4855.45}, "was": {"B6": {"text": "142"}, "B8": {"text": "4855.45"}},
                   "german": "2 echte Zahlen", "format": {"B6": "B2", "B8": "B2"}}
    after = found(filled(einsetzen.applied(cells, [fix])))
    assert not [k for k in after if k.startswith("astext:")]


def test_a_number_that_could_be_a_thousand_times_more_is_not_written():
    """„1,234“ ist auf Deutsch 1,234 und aus einem englischen Export 1234: Das entscheidet man besser selbst."""
    cells = months([("Mittel Umsatz", "AVERAGE(B2:B13)")], text={6: "1,234"})
    hint = found(cells)["astext:B6"]
    assert hint.fix is None


def test_a_date_as_text_becomes_the_date_number():
    cells = {(1, 1): Cell("Datum"), (1, 2): Cell("Umsatz")}
    for r in range(2, 12):
        cells[(r, 1)], cells[(r, 2)] = Cell(f"{r:02d}.01.2026"), Cell(str(100 + r))
    cells[(9, 1)] = Cell("09.01.2026", None, "09.01.2026")
    cells[(2, 4)] = Cell("Ab 5.1.")
    cells[(2, 5)] = Cell("0", 'SUMIFS(B2:B11,A2:A11,">="&DATE(2026,1,5))')
    fix = found(filled(cells))["astext:A9"].fix
    assert fix["cells"] == {"A9": 46031.0} and fix["format"] == {"A9": "A2"} and fix["german"] == "„09.01.2026“ als Datum"
    # Eingesetzt sieht der Prüfer die Zelle wieder als Datum, wie Excel sie dann zeigt.
    assert einsetzen.applied(cells, [fix])[(9, 1)] == Cell("09.01.2026")


def test_an_unsorted_lookup_is_set_to_exact_search_when_every_key_is_found():
    from voiceagent.analyse import _lookups

    cells = {(1, 1): Cell("Code"), (1, 3): Cell("Satz"), (1, 6): Cell("Code"), (1, 7): Cell("Satz")}
    for r, (code, rate) in enumerate(zip("CADB", ("0.3", "0.1", "0.4", "0.2")), start=2):
        cells[(r, 6)], cells[(r, 7)] = Cell(code), Cell(rate)
    # Ungefähr in unsortierten Werten rechnet der eigene Formelrechner nicht nach: Was Excel zeigt, bleibt "0".
    for r, code in enumerate("ABCDB", start=2):
        cells[(r, 1)], cells[(r, 3)] = Cell(code), Cell("0", f"VLOOKUP(A{r},$F$2:$G$5,2)")
    found_ = {f.key: f for f in _lookups(Checked(cells, layout(Sheet("T", cells)), "T", {}))}
    fix = next(f.fix for k, f in found_.items() if k.startswith(("approx:", "nearest:")))
    assert fix["range"] == "C2:C6" and fix["cells"]["C2"] == "=VLOOKUP(A2,$F$2:$G$5,2,FALSE)"
    assert fix["was"]["C6"] == "VLOOKUP(A6,$F$2:$G$5,2)" and fix["german"] == "=SVERWEIS(A2;$F$2:$G$5;2;FALSCH)"
    # Fehlt ein Schlüssel, wären es Stufen (Tarife): dann nicht einfach genau suchen.
    cells[(4, 1)] = Cell("E")
    sheet = Checked(cells, layout(Sheet("T", cells)), "T", {})
    places = [(r, 3) for r in range(2, 7)]
    assert einsetzen.exact(sheet, places, "VLOOKUP", (None, (2, 6, 5, 7)), errors=False) is None


def test_all_fixes_of_a_sheet_together():
    from voiceagent import pruefer

    cells = months([("Mittel Umsatz", "AVERAGE($B$2:$B$11)"), ("Median Kosten", "MEDIAN(C2:C11)")], text={6: "142"})
    hints = pruefer.sheet_result("A.xlsx", "T", {"T": cells})["hints"]
    keys = [h["key"] for h in hints if h.get("fix")]
    assert keys == ["astext:B6", "range:F2", "range:F3"]
    joint = pruefer.together("A.xlsx", "T", {"T": cells}, hints)
    assert joint["state"] == "ok" and joint["count"] == 3 and joint["left"] == 0
    fix = joint["fix"]
    assert fix["cells"] == {"B6": 142.0, "F2": "=AVERAGE($B$2:$B$13)", "F3": "=MEDIAN(C2:C13)"}
    assert fix["range"] == "B6, F2 und F3" and fix["german"] == "3 Korrekturen" and fix["format"] == {"B6": "B2"}
    assert fix["parts"] == ["„142“ als Zahl", "=MITTELWERT($B$2:$B$13)", "=MEDIAN(C2:C13)"]
    # Zusammen eingesetzt und neu gerechnet, wie Excel es tut: keiner der drei Fehler ist mehr da.
    new = einsetzen.refilled(einsetzen.applied(cells, [fix]), layout(Sheet("T", cells)), "T", {}, cells)
    assert not set(keys) & {h["key"] for h in pruefer.sheet_result("A.xlsx", "T", {"T": new})["hints"]}


def test_fixes_that_only_go_wrong_together_are_left_out():
    """E2 auf E3 und E3 auf E2: Jede für sich rechnet, zusammen wäre es ein Kreisbezug."""
    from voiceagent import pruefer

    cells = months([("A", "B2*2"), ("B", "B3*2"), ("C", "B4*2")])

    def hint(key, where, formula):
        return {"key": key, "level": "fehler", "fix": {"range": where, "cells": {where: "=" + formula},
                                                       "was": {where: cells[formel.position(where)].formula}, "german": "=" + formula}}

    hints = [hint("x:F2", "F2", "F3+1"), hint("x:F3", "F3", "F2+1"), hint("x:F4", "F4", "B4*3")]
    joint = pruefer.together("A.xlsx", "T", {"T": cells}, hints)
    assert joint["state"] == "ok" and joint["count"] == 2 and joint["left"] == 1
    assert set(joint["fix"]["cells"]) == {"F2", "F4"}
    # Bleibt nur eine, kommt sie allein: Was dann noch fehlt, zeigt das neue Prüfen.
    alone = pruefer.together("A.xlsx", "T", {"T": cells}, hints[:2])
    assert alone["count"] == 1 and alone["left"] == 1 and alone["fix"] == hints[0]["fix"]
    assert pruefer.together("A.xlsx", "T", {"T": cells}, hints[:1])["state"] == "none"


def test_the_data_comes_before_a_formula_that_works_around_it():
    """Eine Zahl als Text, die SVERWEIS nicht findet: Die Zahl umwandeln ist die Ursache, WERT() in der Formel nur ein
    Umweg. Beides zusammen wäre doppelt."""
    from voiceagent import pruefer

    cells = months([("A", "B2*2")])
    hints = [{"key": "looknear:F2", "level": "fehler", "fix": {"range": "F2", "cells": {"F2": "=VALUE(B2)*2"},
                                                              "was": {"F2": "B2*2"}, "german": "=WERT(B2)*2"}},
             {"key": "astext:B6", "level": "fehler", "fix": {"range": "B6", "cells": {"B6": 142.0}, "was": {"B6": {"text": "142"}},
                                                            "german": "„142“ als Zahl"}}]
    assert pruefer.together("A.xlsx", "T", {"T": cells}, hints)["fix"]["parts"] == ["„142“ als Zahl", "=WERT(B2)*2"]


def test_the_side_pane_gets_all_fixes_together(tmp_path):
    values = [["Monat", "Umsatz", "Kosten", "", "Kennzahl", "Wert"]]
    values += [[f"M{r}", 100 + r * 7, 50 + r, "", "", ""] for r in range(2, 14)]
    values[5][1] = "142"
    formulas = [[None] * 6 for _ in values]
    ok = [v[1] for v in values[1:11] if isinstance(v[1], int)]
    values[1][4:6], formulas[1][5] = ["Mittel Umsatz", sum(ok) / len(ok)], "=AVERAGE($B$2:$B$11)"
    values[2][4:6], formulas[2][5] = ["Median Kosten", 56.5], "=MEDIAN(C2:C11)"
    sheet = {"name": "T", "row": 1, "col": 1, "values": values, "formulas": formulas}
    _, client = client_with(tmp_path, FakeExcel(None))
    head = {"X-VoiceAgent": "1"}
    assert client.post("/api/addin/alle", json={"sheet": "T"}, headers=head).json()["state"] == "again"
    client.post("/api/addin/pruefen", json={"book": "A.xlsx", "active": "T", "sheets": [sheet]}, headers=head)
    joint = client.post("/api/addin/alle", json={"sheet": "T"}, headers=head).json()
    assert joint["state"] == "ok" and joint["count"] == 3
    assert joint["fix"]["was"] == {"B6": {"text": "142"}, "F2": "AVERAGE($B$2:$B$11)", "F3": "MEDIAN(C2:C11)"}


def test_a_running_total_near_the_end_is_not_a_range_that_stops_too_early():
    """=SUMME(D$2:D18) in Zeile 18, heruntergezogen: Sie endet mit Absicht in ihrer Zeile. Einsetzen machte sonst aus
    jeder laufenden Summe die Gesamtsumme."""
    cells = {(1, 2): Cell("Region"), (1, 4): Cell("Umsatz"), (1, 5): Cell("Laufend")}
    for r in range(2, 21):
        cells[(r, 2)] = Cell("Ost")
        cells[(r, 4)], cells[(r, 5)] = Cell(str(50 + r)), Cell("0", f"SUM(D$2:D{r})")
    assert not [k for k in found(filled(cells)) if k.startswith("range:")]
    # Ein fester Bereich, der zu früh endet, bleibt ein Fehler, auch heruntergezogen.
    for r in range(2, 21):
        cells[(r, 5)] = Cell("0", f"D{r}/SUM($D$2:$D$18)")
    hint = found(filled(cells))["range:E2"]
    # Die ganze heruntergezogene Spalte bekommt den längeren Bereich, nicht nur E2.
    assert hint.fix["range"] == "E2:E20" and hint.fix["cells"]["E20"] == "=D20/SUM($D$2:$D$20)"
    assert not [k for k in found(applied(filled(cells), hint.fix)) if k.startswith("range:")]


# --- Ein Hinweis je Fehler, mit Knopf: Muster der Spalte und Bereiche auf einem anderen Blatt ---


def test_a_formula_out_of_its_column_pattern_gets_the_pattern_to_insert():
    cells = {(1, 1): Cell("Fall"), (1, 2): Cell("Menge"), (1, 3): Cell("Doppelt")}
    for r in range(2, 14):
        cells[(r, 1)], cells[(r, 2)] = Cell(f"F{r}"), Cell(str(r))
        cells[(r, 3)] = Cell("0", "B7*3" if r == 7 else f"B{r}*2")
    cells = filled(cells)
    hint = found(cells)["pattern:C7"]
    assert hint.fix["cells"] == {"C7": "=B7*2"} and hint.tip.startswith("Korrigiert: =B7*2, nachgerechnet 14. Ist es Absicht")
    assert not [k for k in found(applied(cells, hint.fix)) if k.endswith(":C7")]


def test_a_lookup_with_the_wrong_column_is_one_hint_with_its_fix():
    """Vorher kam dazu „rechnet anders als die übrigen“ ohne Knopf für dieselbe Zelle."""
    from test_haertetest import _hints

    rates = {"A1": "Station", "B1": "Klinik", "C1": "Betten", "A2": "K1", "B2": "Kardiologie", "C2": 20,
             "A3": "C1", "B3": "Chirurgie", "C3": 18, "A4": "M1", "B4": "Medizin", "C4": 24}
    cases = {"A1": "Fall", "B1": "Station", "C1": "Klinik"}
    for r in range(2, 14):
        station = ["K1", "C1", "M1"][r % 3]
        column = 3 if r == 9 else 2
        shown = rates[f"{'BC'[column - 2]}{['K1', 'C1', 'M1'].index(station) + 2}"]
        cases[f"A{r}"], cases[f"B{r}"] = f"F{r}", station
        cases[f"C{r}"] = ("=", f"VLOOKUP(B{r},Stammdaten!$A$2:$C$4,{column},FALSE)", shown)
    found = {k: h for (s, k), h in _hints(("Fälle", cases), ("Stammdaten", rates)).items() if s == "Fälle"}
    at_c9 = [k for k in found if k.endswith(":C9")]
    assert at_c9 == ["lookcol:C9"], found
    assert found["lookcol:C9"]["fix"]["cells"] == {"C9": "=VLOOKUP(B9,Stammdaten!$A$2:$C$4,2,FALSE)"}


def test_a_range_on_another_sheet_is_pulled_with_its_neighbours():
    from test_haertetest import _cases, _hints

    report = {"A1": "Geschlecht", "B1": "Fälle", "C1": "Kosten"}
    for r, sex in enumerate(["w", "m", "w", "m"], start=2):
        end = 17 if r == 4 else 21
        report[f"A{r}"] = sex
        report[f"B{r}"] = ("=", f"COUNTIF(Fälle!$B$2:$B${end},A{r})", 0)
        report[f"C{r}"] = ("=", f"SUMIF(Fälle!$B$2:$B${end},A{r},Fälle!$C$2:$C${end})", 0)
    found = _hints(("Fälle", _cases(20)), ("Auswertung", report))
    here = {k: h for (s, k), h in found.items() if s == "Auswertung"}
    # Je Zelle ein Hinweis: der genaue mit „4 Zeilen fehlen“ und dem Knopf.
    assert [k for k in here if k.endswith(":B4")] == ["range:B4"], here
    assert here["range:B4"]["fix"]["cells"] == {"B4": "=COUNTIF(Fälle!$B$2:$B$21,A4)"}
    # SUMMEWENN zieht den Bereich daneben mit, sonst passten die beiden nicht zusammen.
    assert here["range:C4"]["fix"]["cells"] == {"C4": "=SUMIF(Fälle!$B$2:$B$21,A4,Fälle!$C$2:$C$21)"}


def test_a_range_on_another_sheet_is_not_pulled_over_a_subtotal():
    """Stehen zwischen dem Ende des Bereichs und den letzten Fällen eine Zwischensumme, gibt es den Hinweis, aber keinen
    Knopf: Bis ganz unten gezogen zählte sie mit."""
    from test_haertetest import _cases, _hints

    costs = [100 + 10 * i for i in range(20)]
    cases = _cases(20, {"A19": "Zwischensumme", "C19": ("=", "SUM(C2:C18)", sum(costs[:17]))})
    cases.pop("B19")
    report = {"A1": "Kennzahl", "B1": "Wert", "A2": "Mittelwert", "B2": ("=", "AVERAGE(Fälle!C2:C18)", sum(costs[:17]) / 17)}
    hint = _hints(("Fälle", cases), ("Auswertung", report))[("Auswertung", "range:B2")]
    assert "2 Zeilen fehlen (Fälle!C20, Fälle!C21)" in hint["text"] and "fix" not in hint


def test_a_short_range_in_one_row_gets_the_formula_of_its_column():
    """Der genaue Hinweis („3 Zeilen fehlen“) bleibt und bekommt den Knopf mit der Formel der Nachbarn."""
    cells = {(1, 1): Cell("Fall"), (1, 2): Cell("Station"), (1, 3): Cell("Fälle Station")}
    for r in range(2, 22):
        cells[(r, 1)], cells[(r, 2)] = Cell(f"F{r}"), Cell(["K1", "C1", "M1"][r % 3])
        cells[(r, 3)] = Cell("0", "COUNTIF($B$2:$B$18,B9)" if r == 9 else f"COUNTIF($B$2:$B$21,B{r})")
    cells = filled(cells)
    at_c9 = {k: f for k, f in found(cells).items() if k.endswith(":C9")}
    assert len(at_c9) == 1 and not [k for k in at_c9 if k.startswith("pattern:")], at_c9
    hint = next(iter(at_c9.values()))
    assert hint.fix["cells"] == {"C9": "=COUNTIF($B$2:$B$21,B9)"}
    assert not [k for k in found(applied(cells, hint.fix)) if k.endswith(":C9")]
