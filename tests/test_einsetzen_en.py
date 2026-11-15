"""Korrekturen zum Einsetzen aus englischen Tipps: dieselbe Korrektur wie aus dem deutschen Tipp."""

import pytest
from test_statistik import filled

from voiceagent import einsetzen, sprache
from voiceagent.analyse import Finding
from voiceagent.analyse import Sheet as Checked
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Sheet


def _sheet():
    cells = {(1, 1): Cell("Fall"), (1, 2): Cell("Punkte"), (1, 3): Cell("Doppelt")}
    for r in range(2, 8):
        cells[(r, 1)], cells[(r, 2)], cells[(r, 3)] = Cell(f"F{r}"), Cell(str(r)), Cell("0", f"B{r}*2")
    return Checked(filled(cells), layout(Sheet("T", cells)), "T", {})


@pytest.mark.parametrize("german, english", [
    ("Korrigiert: =B2*3, nachgerechnet 6.", "Corrected: =B2*3, recalculated 6."),
    ("Korrigiert: =B2*3, nachgerechnet ergibt das 6.", "Corrected: =B2*3, recalculated this gives 6."),
    ("Korrigiert in C2: =B2*3, dann bis C5 herunterziehen.", "Corrected in C2: =B2*3, then fill down to C5."),
    ("Korrigiert: =B2*3, dann bis unten herunterziehen.", "Corrected: =B2*3, then fill it down to the bottom."),
    ("Schreib in C2 =B2*3 und zieh es herunter bis C7.", "In C2, write =B2*3 and fill it down to C7."),
    ("Schreib in C2 =B2*3 und zieh es herunter bis C7.", "Write =B2*3 in C2 and fill it down to C7."),
    ('Schreib in C2 =B2*3 ("" statt 0) und zieh es bis C4.', 'In C2, write =B2*3 ("" instead of 0) and copy it to C4.'),
    ("Nimm =B2*3. Nachgerechnet ergibt das 6.", "Use =B2*3. Recalculated, this gives 6."),
    ("Sicherer ist =B2*3, nachgerechnet 6.", "Safer is =B2*3, recalculated 6."),
    ("Gemeint ist meist die Quote über alle Zeilen: =B2*3, nachgerechnet 6.",
     "What is usually meant is the rate across all rows: =B2*3, recalculated 6."),
])
def test_an_english_tip_gives_the_same_fix(german, english):
    sheet = _sheet()
    want = einsetzen.fix(sheet, Finding("x:C2", "", "fehler", german))
    assert want is not None
    with sprache.using("en"):
        assert einsetzen.fix(sheet, Finding("x:C2", "", "fehler", english)) == want


def test_an_english_tip_with_another_value_gives_no_fix():
    with sprache.using("en"):
        assert einsetzen.fix(_sheet(), Finding("x:C2", "", "fehler", "Corrected: =B2*3, recalculated 99.")) is None


def test_several_fixes_and_cleared_cells_are_named_in_english():
    one = {"range": "C2", "cells": {"C2": "=B2"}, "was": {"C2": ""}, "german": "=B2"}
    two = {"range": "C7", "cells": {"C7": "=B7"}, "was": {"C7": ""}, "german": "=B7"}
    with sprache.using("en"):
        both = einsetzen.merged([one, two])
    assert both["german"] == "2 fixes" and both["range"] == "C2 and C7"
    assert einsetzen.merged([one, two])["german"] == "2 Korrekturen"
