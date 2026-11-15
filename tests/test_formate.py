"""Zahlen und Daten als Text in allen Schreibweisen, die aus Importen kommen: Buchhaltung (12,50) und 12,50-, Währung
vorne oder hinten mit Vorzeichen, schmale Leerzeichen als Tausendertrennung, Prozent, 1,5E+03, Daten mit Schrägstrich
(Tag oder Monat vorne, aus der Spalte erkannt), zweistellige Jahre und Monatsnamen."""

from test_workbook import xlsx

from voiceagent import alstext, pruefer

STYLES = ('<cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/>'
          '<xf numFmtId="10" applyNumberFormat="1"/></cellXfs>')


def test_numbers_written_every_way_mean_what_they_say():
    assert alstext.number("(12,50)") == -12.5 and alstext.number("12,50-") == -12.5
    assert alstext.number("-€12.00") == -12.0 and alstext.number("€-12.00") == -12.0
    assert alstext.number("15 000,50") == 15000.5 and alstext.number("CHF 1'200.00") == 1200.0
    assert alstext.number("78,5 %") == 0.785 and alstext.number("1,5E+03") == 1500.0 and alstext.number("+12") == 12.0
    assert alstext.number("−7") == -7.0
    # Telefonnummern und Bereiche sind keine Zahlen.
    assert alstext.number("+41 44 123 45 67") is None and alstext.number("1-2") is None
    # Ein Vorzeichen oder eine Klammer allein ist nichts, und die Prüfung läuft weiter.
    assert all(alstext.number(t) is None for t in ("(", "+", "-", "( - )", "€", "%", ""))


def test_dates_written_every_way_mean_the_same_day():
    feb23 = 46076.0
    for text in ["23.02.2026", "23.2.26", "2026-02-23", "2026/02/23", "23/02/2026", "23. Februar 2026", "23 Feb 2026",
                 "Feb 23, 2026", "23-Feb-26"]:
        assert alstext.day(text) == feb23, text
    assert alstext.day("02/23/2026", "mdy") == feb23 and alstext.day("02/23/2026") is None
    assert alstext.day("31.02.2026") is None
    assert alstext.slash_order(["23/02/2026", "01/03/2026"]) == "dmy" and alstext.slash_order(["02/23/2026"]) == "mdy"
    assert alstext.slash_order(["01/02/2026"]) is None


def _book(texts: dict, column: str = "Betrag", style: int | None = None, values=None) -> dict:
    cells = {"A1": "Posten", "B1": column}
    total = 0.0
    for r in range(2, 16):
        value = (values or (lambda r: 100.0 + 7 * r))(r)
        cells[f"A{r}"] = f"Posten {r - 1}"
        if r in texts:
            cells[f"B{r}"] = texts[r]
        else:
            cells[f"B{r}"] = (value, style) if style is not None else value
            total += value
    cells.update({"A16": "Summe", "B16": ("=", "SUM(B2:B15)", total)})
    done = pruefer.Checker(None).check_file("mappe.xlsx", xlsx(("Daten", cells), styles=STYLES))
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


def test_numbers_as_text_from_an_import_are_found_with_their_value():
    found = _book({3: "(12,50)", 5: "1,5E+03", 8: "-€12.00", 11: "15 000,50", 13: "+12"})
    hint = next(h for k, h in found.items() if k.startswith("astext:"))
    assert "In Betrag stehen 5 Zahlen als Text" in hint["text"]
    assert hint["fix"]["cells"] == {"B3": -12.5, "B5": 1500.0, "B8": -12.0, "B11": 15000.5, "B13": 12.0}


def test_a_percent_as_text_means_the_percent_of_its_column():
    percent = _book({4: "78,5 %"}, "Anteil", style=2, values=lambda r: 0.05 * r)
    assert next(h for k, h in percent.items() if k.startswith("astext:"))["fix"]["cells"] == {"B4": 0.785}
    points = _book({4: "78,5 %"}, "Anteil (%)", values=lambda r: 5.0 * r)
    assert next(h for k, h in points.items() if k.startswith("astext:"))["fix"]["cells"] == {"B4": 78.5}


def test_dates_with_slashes_follow_their_column_and_are_only_inserted_when_sure():
    dates = lambda r: 46020.0 + r  # noqa: E731
    sure = _book({3: "23/02/2026", 6: "03/04/2026"}, "Datum", style=1, values=dates)
    hint = next(h for k, h in sure.items() if k.startswith("astext:"))
    # 23/02 zeigt: Der Tag steht vorne, also ist 03/04 der 3. April.
    assert hint["fix"]["cells"] == {"B3": 46076.0, "B6": 46115.0}
    unsure = _book({3: "03/04/2026"}, "Datum", style=1, values=dates)
    hint = next(h for k, h in unsure.items() if k.startswith("astext:"))
    assert "fix" not in hint


def test_a_count_with_its_unit_word_is_a_number_as_text():
    """„12 Stk“ in einer Mengenspalte meint 12. Andere Einheiten nur, wenn der Kopf sie nennt."""
    assert alstext.with_unit("12 Stk", "Menge") == (12.0, "Stk") and alstext.with_unit("1'200 Stück", "Menge")[0] == 1200.0
    assert alstext.with_unit("72 kg", "Gewicht kg") == (72.0, "kg")
    assert alstext.with_unit("500 g", "Gewicht kg") is None and alstext.with_unit("12 A", "Zimmer") is None
    hint = next(h for k, h in _book({6: "12 Stk"}, "Menge").items() if k.startswith("astext:"))
    assert hint["text"].startswith("In Menge steht bei B6 („12 Stk“) eine Zahl mit Einheit als Text.")
    assert hint["fix"]["cells"] == {"B6": 12.0} and hint["tip"].startswith("Trag nur die Zahl ein")
    # Eine andere Einheit als im Kopf bleibt hier still (dafür gibt es den Hinweis zu zwei Einheiten).
    assert not any(k.startswith("astext:") for k in _book({6: "500 g"}, "Gewicht kg"))
