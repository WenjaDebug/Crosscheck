"""Mappe in Worten: was die ganze Mappe rechnet, von den Daten bis zu den Ergebnissen."""

from pathlib import Path

from test_addin import payload
from test_excel_live import FakeExcel, client_with
from test_herleitung import rechnung, tarife

from voiceagent import bericht, inworten
from voiceagent.xlsx import Cell, Sheet, read_book

DEMO = Path(__file__).resolve().parents[1] / "src" / "voiceagent" / "web" / "demo" / "pruefer-uebung.xlsx"


def describe(cells_of, hints=None):
    return inworten.describe("Test.xlsx", [Sheet(name, cells) for name, cells in cells_of.items()], hints)


def results(found):
    return {(r["sheet"], r["cell"]): r for r in found["results"]}


def sheets(found):
    return {s["name"]: s for s in found["sheets"]}


def werte(rows=(10, 20, 30, 40), formula=None):
    """Name und Wert, darunter die Summe in B6; formula: eine gerechnete Spalte C ({r}: die Zeile, {d}: die Zeile der
    Summe, wenn C2 auf sie zeigt und ohne $ heruntergezogen ist)."""
    cells = {(1, 1): Cell("Name"), (1, 2): Cell("Wert")}
    for r, value in enumerate(rows, start=2):
        cells[(r, 1)], cells[(r, 2)] = Cell(f"N{r}"), Cell(str(value))
    last = len(rows) + 2
    cells[(last, 1)], cells[(last, 2)] = Cell("Summe"), Cell(str(sum(rows)), f"SUM(B2:B{last - 1})")
    if formula:
        cells[(1, 3)] = Cell("Anteil")
        for r, value in enumerate(rows, start=2):
            cells[(r, 3)] = Cell(f"{value / sum(rows):g}", formula.format(r=r, d=r + len(rows)))
    return cells


def test_a_lookup_across_sheets_from_the_data_to_the_result():
    found = describe({"Rechnung": rechnung(["K", "A", "M"], wrong=4), "Tarife": tarife()})
    assert found["state"] == "ok"
    # Das Blatt mit den Daten zuerst, auch wenn es in der Mappe hinten steht.
    assert [s["name"] for s in found["sheets"]] == ["Tarife", "Rechnung"]
    assert found["summary"] == [
        "2 Blätter mit Inhalt, 7 Formeln.",
        "Daten stehen in „Tarife“ (3 Zeilen) und „Rechnung“ (3 Zeilen).",
        "Nachgerechnet ergibt 1 Formel etwas anderes als in Excel (Rechnung!C4), 6 Formeln ergeben dasselbe.",
        "„Rechnung“ liest aus „Tarife“ (SVERWEIS).",
        "1 Ergebnis, das keine andere Formel weiterverwendet, steht unten mit dem, woraus es entsteht.",
    ]
    assert found["flow"] == [{"from": "Tarife", "to": "Rechnung", "formulas": 3, "columns": ["Code", "Satz"],
                              "functions": ["SVERWEIS"]}]
    total = results(found)[("Rechnung", "D5")]
    assert total["name"] == "Total Summe" and total["value"] == "50" and total["check"] == "ok"
    assert total["text"] == [
        "Entsteht aus Code und Betrag sowie Code und Satz („Tarife“), gerechnet über Total und Satz.",
        "Unterwegs ergibt nachgerechnet C4 etwas anderes als in Excel: Dann stimmt womöglich auch dieses Ergebnis nicht.",
    ]
    rate = sheets(found)["Rechnung"]["columns"][0]
    assert rate["name"] == "Satz" and rate["words"] == "=SVERWEIS(Code; Tarife!$A$2:$B$4; 2; FALSCH)"
    assert rate["check"] == "anders" and rate["wrong"] == ["C4"]
    assert rate["text"] == ["In allen 3 Zeilen dieselbe Formel.", "Nachgerechnet anders als in Excel: C4."]
    assert sheets(found)["Tarife"]["read_by"] == ["Rechnung"]
    assert sheets(found)["Rechnung"]["reads"] == [{"sheet": "Tarife", "columns": ["Code", "Satz"], "functions": ["SVERWEIS"],
                                                  "formulas": 3}]


def test_without_mistakes_and_with_hints_from_the_checker():
    found = describe({"Tarife": tarife(), "Rechnung": rechnung(["K", "A", "M"])},
                     {"Tarife": {"fehler": 0, "pruefen": 0}, "Rechnung": {"fehler": 1, "pruefen": 2}})
    assert "Nachgerechnet ergeben alle 7 Formeln, die Gegenprobe nachrechnen kann, dasselbe wie in Excel." in found["summary"]
    assert "Fehler und Stellen zum Prüfen gibt es auf „Rechnung“ (1 Fehler, 2 zum Prüfen)." in found["summary"]
    assert "Unterwegs" not in " ".join(results(found)[("Rechnung", "D5")]["text"])
    assert sheets(found)["Rechnung"]["hints"] == {"fehler": 1, "pruefen": 2}
    quiet = describe({"Tarife": tarife(), "Rechnung": rechnung(["K", "A", "M"])},
                     {"Tarife": {"fehler": 0, "pruefen": 0}, "Rechnung": {"fehler": 0, "pruefen": 0}})
    assert "Fehler oder Stellen zum Prüfen hat Gegenprobe auf keinem Blatt gefunden." in quiet["summary"]
    # Noch nicht jedes Blatt geprüft: dann kein „auf keinem Blatt“.
    partly = describe({"Tarife": tarife(), "Rechnung": rechnung(["K", "A", "M"])}, {"Tarife": {"fehler": 0, "pruefen": 0}})
    assert not any("Prüfen" in line for line in partly["summary"])


def test_a_reference_without_dollar_moves_down_with_the_formula():
    moving = sheets(describe({"T": werte(formula="B{r}/B{d}")}))["T"]["columns"][0]
    assert moving["words"] == "=Wert / Wert Summe" and moving["same"] == 4
    assert moving["moving"] == ("Der Bezug auf B6 (Summe) hat kein $ vor der Zeile und wandert beim Herunterziehen "
                                "mit: C3 nimmt B7, C5 schon B9.")
    fixed = sheets(describe({"T": werte(formula="B{r}/B$6")}))["T"]["columns"][0]
    assert "moving" not in fixed and fixed["text"] == ["In allen 4 Zeilen dieselbe Formel."]
    # Von Hand in jede Zeile geschrieben (=B3/B6): dieselbe Formel, nichts wandert.
    typed = sheets(describe({"T": werte(formula="B{r}/B6")}))["T"]["columns"][0]
    assert "moving" not in typed and typed["same"] == 4 and typed["check"] == "ok"
    # Ein Bezug in die eigene Zeile oder eine andere Zeile der Tabelle wandert gewollt mit.
    own = sheets(describe({"T": werte(formula="B{r}*2")}))["T"]["columns"][0]
    assert "moving" not in own and own["constants"] == ["2"]


def test_results_are_what_no_other_formula_uses():
    cells = werte(formula="B{r}/B$6")
    found = describe({"T": cells})
    # Die Summe B6 braucht die Spalte Anteil, also ist sie kein Ergebnis; die Spalte Anteil schon.
    assert list(results(found)) == [("T", "C2:C5")]
    share = results(found)[("T", "C2:C5")]
    assert share["kind"] == "spalte" and share["text"] == ["Entsteht aus Wert, gerechnet über B6 (Summe)."]
    # Liest ein anderes Blatt die Spalte, ist das Ergebnis dort.
    other = {(1, 1): Cell("Größter Anteil"), (1, 2): Cell("0,4", "MAX(T!C2:C5)")}
    found = describe({"T": cells, "Auswertung": other})
    assert list(results(found)) == [("Auswertung", "B1")]
    top = results(found)[("Auswertung", "B1")]
    assert top["name"] == "Größter Anteil" and top["check"] == "ok"
    assert top["text"] == ["Entsteht aus Wert („T“), gerechnet über T: Anteil und T: B6 (Summe)."]
    assert found["summary"][-2:] == ["„Auswertung“ liest aus „T“ (MAX).",
                                     "1 Ergebnis, das keine andere Formel weiterverwendet, steht unten mit dem, woraus es entsteht."]


def test_inputs_outside_a_table_are_named_by_the_word_above():
    cells = {(1, 1): Cell("Monat"), (1, 2): Cell("Patienten")}
    for r in range(2, 14):
        cells[(r, 1)], cells[(r, 2)] = Cell(f"M{r}"), Cell("10")
    # Ein Zinssatz rechts neben der Tabelle, in der Zeile der Spaltenköpfe.
    cells[(1, 7)], cells[(1, 8)] = Cell("Zinssatz"), Cell("3 %")
    cells[(2, 7)], cells[(2, 8)] = Cell("Summe mit Zins"), Cell("123,6", "SUM(B2:B13)*(1+H1)")
    found = describe({"S": cells})
    assert sheets(found)["S"]["table"] == {"area": "A1:B13", "rows": 12, "inputs": ["Monat", "Patienten"], "computed": []}
    assert results(found)[("S", "H2")]["text"] == ["Entsteht aus Patienten und H1 (Zinssatz)."]
    # Zahlen unter einer zweiten Überschrift weiter unten: der Bereich mit dem Wort darüber.
    cells[(20, 1)], cells[(20, 2)] = Cell("Woche"), Cell("Betten")
    for r in range(21, 25):
        cells[(r, 1)], cells[(r, 2)] = Cell(f"W{r}"), Cell("5")
    cells[(26, 1)], cells[(26, 2)] = Cell("Betten gesamt"), Cell("20", "SUM(B21:B24)")
    assert results(describe({"S": cells}))[("S", "B26")]["text"] == ["Entsteht aus Betten (B21:B24)."]


def test_today_and_formulas_it_cannot_recompute():
    cells = {(1, 1): Cell("Geburt"), (1, 2): Cell("Alter")}
    for r in range(2, 5):
        cells[(r, 1)] = Cell("01.01.1980")
        cells[(r, 2)] = Cell("46", f'DATEDIF(A{r},TODAY(),"Y")')
    found = describe({"S": cells})
    age = results(found)[("S", "B2:B4")]
    assert age["check"] == "unbekannt" and age["today"]
    assert age["text"][-1].startswith("Hängt vom heutigen Datum ab (HEUTE())")
    assert found["summary"][2] == "3 Formeln rechnet Gegenprobe nicht nach (darunter solche mit HEUTE(), das jeden Tag etwas anderes ergibt)."
    column = sheets(found)["S"]["columns"][0]
    assert column["check"] == "unbekannt" and column["today"]


def test_empty_books_cycles_and_the_time_limit(monkeypatch):
    assert describe({"Leer": {}})["state"] == "empty"
    # Zwei Blätter, die gegenseitig voneinander lesen: Die Reihenfolge der Mappe gilt, nichts hängt.
    a = {(1, 1): Cell("Wert A"), (1, 2): Cell("2", "B!B1+1")}
    b = {(1, 1): Cell("Wert B"), (1, 2): Cell("1", "A!B1-1")}
    found = describe({"A": a, "B": b})
    assert [s["name"] for s in found["sheets"]] == ["A", "B"] and found["results"] == []
    assert "„A“ liest aus „B“ und „B“ liest aus „A“." in found["summary"]
    # Sehr viele Formeln: in den gerechneten Spalten nur eine Auswahl, und das steht dabei.
    monkeypatch.setattr(inworten, "MAX_CHECKS", 2)
    found = describe({"T": werte(formula="B{r}/B$6")})
    assert found["summary"][2] == ("Nachgerechnet habe ich 2 der 5 Formeln (in den gerechneten Spalten jede 4.): Alle, die "
                                   "Gegenprobe nachrechnen kann, ergeben dasselbe wie in Excel.")
    # Keine Zeit mehr: Was nicht nachgerechnet ist, heißt nicht „stimmt“.
    monkeypatch.setattr(inworten, "MAX_CHECKS", 5000)
    monkeypatch.setattr(inworten, "BUDGET", 0.0)
    found = describe({"T": werte(formula="B{r}/B$6")})
    assert found["summary"][-1] == "5 Formeln habe ich nicht mehr nachgerechnet, weil die Mappe so groß ist."
    assert not any(line.startswith("Nachgerechnet") for line in found["summary"])
    assert results(found)[("T", "C2:C5")]["check"] == "unbekannt"


def test_the_practice_workbook_in_words():
    book = read_book(DEMO.read_bytes())
    found = inworten.describe("pruefer-uebung.xlsx", book.sheets)
    assert found["summary"][0] == "6 Blätter mit Inhalt, 139 Formeln."
    assert "„Rechnung“ liest aus „Tarife“ (SVERWEIS)." in found["summary"]
    every = results(found)
    age = sheets(found)["Fälle"]["figures"][0]
    assert age["cell"] == "L2" and age["name"] == "Mittelwert Alter"
    assert age["words"] == "Nimmt den Mittelwert von Alter aus den 37 Zeilen E2:E38, also allen außer E39:E41."
    assert every[("Fälle", "L3")]["text"] == ["Entsteht aus Eintritt und Austritt, gerechnet über Liegedauer."]
    share = next(c for c in sheets(found)["Rechnung"]["columns"] if c["name"] == "Anteil")
    assert share["moving"].startswith("Der Bezug auf D9 (Total Summe) hat kein $ vor der Zeile")
    assert every[("Rechnung", "E2:E8")]["inputs"][1] == {"sheet": "Tarife", "names": ["Code", "Satz"], "more": 0}
    assert every[("Zeiträume", "H2:H17")]["today"]
    assert every[("Zeiträume", "E2:E17")]["constants"] == ["365"]
    assert every[("Pivot und Diagramm", "B33")]["text"] == ["Entsteht aus Patienten (B21:B32)."]
    parts = [p["text"] for p in sheets(found)["Pivot und Diagramm"]["parts"]]
    assert "Säulendiagramm „Patienten je Monat“ mit Patienten nach Monat" in parts


def test_the_pane_and_the_report_describe_the_open_book(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    headers = {"X-VoiceAgent": "1"}
    assert client.post("/api/addin/mappe", headers=headers).json()["state"] == "again"
    assert client.post("/api/addin/mappe").status_code == 403
    client.post("/api/addin/pruefen", json=payload(), headers=headers)
    found = client.post("/api/addin/mappe", headers=headers).json()
    assert found["state"] == "ok" and found["book"] == "Wahl.xlsx"
    assert found["summary"][0] == "1 Blatt mit Inhalt, 7 Formeln."
    # Die Hinweise des Prüfers stehen dabei, je Blatt.
    assert sheets(found)["Wahl"]["hints"]["fehler"] >= 1
    assert any(line.startswith("Fehler und Stellen zum Prüfen gibt es auf „Wahl“") for line in found["summary"])
    full = client.post("/api/addin/bericht", json={"daten": True}, headers=headers).json()
    page = client.get(full["url"]).text
    assert "Was die Mappe rechnet" in page and "Was dieses Blatt rechnet" in page
    plain = client.post("/api/addin/bericht", json={"daten": False}, headers=headers).json()
    page = client.get(plain["url"]).text
    assert "Was die Mappe rechnet" not in page and "Wahl" not in page


def test_the_report_of_a_dropped_file_describes_it_too():
    found = describe({"Tarife": tarife(), "Rechnung": rechnung(["K", "A", "M"])})
    page = bericht.render("Rechnung.xlsx", [("Tarife", {"hints": []}), ("Rechnung", {"hints": []})], True, mappe=found)
    assert "Was die Mappe rechnet" in page and "Entsteht aus Code und Betrag sowie Code und Satz („Tarife“)" in page
    assert "=Betrag * Satz" in page
