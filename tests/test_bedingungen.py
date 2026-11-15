"""Eine Bedingung mit Tippfehler („Nrod“) oder Leerzeichen („Nord “) in SUMMEWENNS findet nichts: Steht sie als Text in
der Formel, kommt die Korrektur zum Einsetzen mit, nachgerechnet."""

from test_workbook import xlsx

from voiceagent import pruefer

REGIONS = ["Nord", "Süd", "Ost", "West"]


def _hints(wanted: str) -> dict:
    cells = {"A1": "Region", "B1": "Umsatz", "D1": "Region", "E1": "Umsatz"}
    for i in range(12):
        cells.update({f"A{i + 2}": REGIONS[i % 4], f"B{i + 2}": 100 + 10 * i})
    cells.update({"D2": "Nord", "E2": ("=", f'SUMIFS(B2:B13,A2:A13,"{wanted}")', 0)})
    done = pruefer.Checker(None).check_file("mappe.xlsx", xlsx(("Daten", cells)))
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


def test_a_misspelt_condition_comes_with_the_formula_to_insert():
    hint = _hints("Nrod")["criterion:E2"]
    assert "Gemeint ist wohl „Nord“" in hint["text"]
    assert hint["fix"]["cells"] == {"E2": '=SUMIFS(B2:B13,A2:A13,"Nord")'}
    assert hint["tip"] == 'Korrigiert: =SUMMEWENNS(B2:B13;A2:A13;"Nord"), nachgerechnet 420.'


def test_a_condition_with_a_space_comes_with_the_formula_to_insert():
    hint = _hints("Nord ")["criterion:E2"]
    assert "mit Leerzeichen" in hint["text"]
    assert hint["fix"]["cells"] == {"E2": '=SUMIFS(B2:B13,A2:A13,"Nord")'}
