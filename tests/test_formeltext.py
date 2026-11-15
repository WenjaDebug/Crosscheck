"""Eine Formel als Text: In einer als Text formatierten Zelle bleibt =SUMME(B2:B6) stehen, wie getippt. Gemeldet mit
dem, was die Formel ergäbe und was die Zellen, die sie lesen, damit anders rechnen. Einsetzen stellt die Zelle auf
„Standard“ und schreibt die Formel. Formeln zum Nachlesen (Spalte „Formel“, neben der echten Formel) bleiben still."""

from test_workbook import xlsx

from voiceagent import addin, einsetzen, pruefer


def _kosten(**changes) -> dict:
    cells = {"A1": "Posten", "B1": "Betrag", "C1": "Doppelt"}
    for r, (name, amount) in enumerate([("Miete", 1200.0), ("Strom", 180.0), ("Wasser", 60.0), ("Internet", 45.0),
                                         ("Versicherung", 95.0)], start=2):
        cells.update({f"A{r}": name, f"B{r}": amount, f"C{r}": ("=", f"B{r}*2", amount * 2)})
    cells.update({"A7": "Summe", "B7": "=SUMME(B2:B6)", "A8": "Mit Reserve", "B8": ("=", "B7*1.1", 0.0)})
    cells.update(changes)
    return cells


def _hints(cells: dict) -> dict:
    done = pruefer.Checker(None).check_file("kosten.xlsx", xlsx(("Kosten", cells)))
    assert done["state"] == "ok", done
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


def test_a_german_sum_typed_as_text_is_an_error_with_einsetzen():
    found = _hints(_kosten())
    hint = found["formeltext:B7"]
    assert hint["level"] == "fehler"
    assert hint["text"].startswith("B7 (Betrag Summe) zeigt =SUMME(B2:B6) als Text, statt zu rechnen")
    assert "Darum ergibt B8 (Betrag Mit Reserve) #WERT! statt 1738" in hint["text"]
    assert hint["tip"] == "Korrigiert in B7: =SUMME(B2:B6), nachgerechnet 1580."
    fix = hint["fix"]
    assert fix["cells"] == {"B7": "=SUM(B2:B6)"} and fix["general"] == ["B7"] and fix["was"] == {"B7": "SUMME(B2:B6)"}
    assert fix["german"] == "=SUMME(B2:B6)"
    assert not [k for k in found if k.startswith("text:")]


def test_an_english_formula_in_a_formula_column():
    hint = _hints(_kosten(C4="=B4*2"))["formeltext:C4"]
    assert hint["level"] == "fehler" and hint["fix"]["cells"] == {"C4": "=B4*2"}


def test_notes_titles_and_documentation_stay_quiet():
    cells = _kosten(B7=("=", "SUM(B2:B6)", 1580.0), B8=("=", "B7*1.1", 1738.0))
    cells.update({"D1": "Formel", "D7": "=SUMME(B2:B6)", "E7": "=SUM(B2:B6)", "A10": "== Ergebnis ==", "A11": "=Netto",
                  "A12": "= Summe aus Miete und Strom", "F8": "=B7*1.1", "A13": "=MEINEFUNKTION(A2:A6)"})
    assert not [k for k in _hints(cells) if k.startswith("formeltext:")]


def test_the_pane_reads_a_text_formula_as_text_and_einsetzen_makes_it_a_formula():
    values = [["Posten", "Betrag"], ["Miete", 1200], ["Strom", 180], ["Summe", "=SUMME(B2:B3)"]]
    formulas = [["Posten", "Betrag"], ["Miete", 1200], ["Strom", 180], ["Summe", "=SUMME(B2:B3)"]]
    sheet = addin.sheets_from({"sheets": [{"name": "K", "row": 1, "col": 1, "values": values, "formulas": formulas}]})[2][0]
    assert sheet.cells[(4, 2)].formula is None and sheet.cells[(4, 2)].value == "=SUMME(B2:B3)"
    done = pruefer.Checker(None).check_payload({"book": "K.xlsx", "active": "K", "sheets": [
        {"name": "K", "row": 1, "col": 1, "values": values, "formulas": formulas}]})
    hint = next(h for h in done["hints"] if h["key"] == "formeltext:B4")
    after = einsetzen.applied(sheet.cells, [hint["fix"]])
    assert after[(4, 2)].formula == "SUM(B2:B3)"
    assert einsetzen.merged([hint["fix"], {"cells": {"Z1": 1.0}, "was": {"Z1": {"text": "x"}}, "german": "1"}])["general"] == ["B4"]


def _payload(rows: list[tuple[str, object, str | None]]) -> list[dict]:
    values = [["Posten", "Betrag"]] + [[name, value] for name, value, _ in rows]
    formulas = [["Posten", "Betrag"]] + [[name, formula or value] for name, value, formula in rows]
    done = pruefer.Checker(None).check_payload({"book": "K.xlsx", "active": "K", "sheets": [
        {"name": "K", "row": 1, "col": 1, "values": values, "formulas": formulas}]})
    return done["hints"]


def test_the_errors_it_causes_are_told_once_at_the_text():
    """B6 und B7 zeigen #WERT!, weil B5 Text ist: Das sagt der Hinweis zu B5, kein eigener Hinweis je Fehler."""
    hints = _payload([("Miete", 1200, None), ("Strom", 180, None), ("Wasser", 60, None), ("Summe", "=SUMME(B2:B4)", None),
                      ("MwSt", "#VALUE!", "=B5*0.19"), ("Gesamt", "#VALUE!", "=B5+B6")])
    keys = {h["key"]: h for h in hints}
    assert "formeltext:B5" in keys and not [k for k in keys if k.startswith("error:")]
    assert "Darum ergibt B6 (Betrag MwSt) #WERT! statt 273,6 und B7 (Betrag Gesamt) #WERT! statt 1713,6" in keys["formeltext:B5"]["text"]


def test_rows_computed_from_the_total_are_not_its_parts():
    """Summe, darunter MwSt = Summe * 0,19 und Gesamt: Die MwSt fehlt nicht in der Summe."""
    hints = _payload([("Miete", 1200, None), ("Strom", 180, None), ("Wasser", 60, None), ("Summe", 1440, "=SUM(B2:B4)"),
                      ("MwSt", 273.6, "=B5*0.19"), ("Gesamt", 1713.6, "=B5+B6")])
    assert hints == []
    hints = _payload([("Miete", 1200, None), ("Strom", 180, None), ("Wasser", 60, None), ("Summe", 1380, "=SUM(B2:B3)"),
                      ("MwSt", 262.2, "=B5*0.19")])
    assert any(h["key"] == "parts:B5" and "zählt aber B4 (Wasser, 60) nicht mit" in h["text"] for h in hints)
