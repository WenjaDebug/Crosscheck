"""Herleitung einer Zahl: woraus sie entsteht (auch über Blätter und SVERWEIS) und wohin sie fließt."""

from test_addin import payload
from test_excel_live import FakeExcel, client_with

from voiceagent.herleitung import explain
from voiceagent.xlsx import Cell

RATES = {"K": 0.1, "A": 0.2, "M": 0.3}


def tarife():
    cells = {(1, 1): Cell("Code"), (1, 2): Cell("Satz")}
    for r, (code, rate) in enumerate(RATES.items(), start=2):
        cells[(r, 1)], cells[(r, 2)] = Cell(code), Cell(f"{rate * 100:g} %")
    return cells


def rechnung(codes, wrong=None):
    """Code, Betrag, Satz (SVERWEIS auf Tarife), Total, darunter die Summe. wrong: diese Zeile zeigt Excel falsch an."""
    cells = {(1, 1): Cell("Code"), (1, 2): Cell("Betrag"), (1, 3): Cell("Satz"), (1, 4): Cell("Total")}
    total = 0.0
    for r, code in enumerate(codes, start=2):
        rate = 0.2 if r == wrong else RATES[code]
        cells[(r, 1)], cells[(r, 2)] = Cell(code), Cell("100")
        cells[(r, 3)] = Cell(f"{rate * 100:g} %", f"VLOOKUP(A{r},Tarife!$A$2:$B$4,2,FALSE)")
        cells[(r, 4)] = Cell(f"{rate * 100:g}", f"B{r}*C{r}")
        total += rate * 100
    last = len(codes) + 2
    cells[(last, 1)], cells[(last, 4)] = Cell("Summe"), Cell(f"{total:g}", f"SUM(D2:D{last - 1})")
    return cells


def flat(node, depth=0):
    yield depth, node
    for child in node.get("children", []):
        yield from flat(child, depth + 1)


def test_the_path_to_a_wrong_lookup():
    found = explain({"Tarife": tarife(), "Rechnung": rechnung(["K", "A", "M"], wrong=4)}, "Rechnung", "D5")
    assert found["state"] == "ok" and found["problems"] == ["C4"]
    top = found["tree"]
    assert top["check"] == "ok" and top["words"] == "=SUMME(Total K bis M)" and top["uses"] == []
    nodes = {(n["sheet"], n["cell"]): n for _, n in flat(top)}
    assert nodes[("Rechnung", "D2:D4")]["value"] == "3 Zahlen, Summe 50"
    wrong = nodes[("Rechnung", "C4")]
    assert wrong["check"] == "anders" and wrong["value"] == "20 %" and wrong["ours"] == "30 %"
    assert nodes[("Tarife", "A2:B4")]["value"] == "3 Zeilen, 2 Spalten"
    hit = nodes[("Tarife", "B4")]
    assert hit["note"] == "gefunden von SVERWEIS" and hit["value"] == "30 %" and hit["kind"] == "eingabe"
    # Aus der Tabelle nur die Fundstelle, nicht jede Zelle.
    assert ("Tarife", "A4") not in nodes


def test_large_ranges_show_only_what_does_not_add_up():
    codes = ["K", "A", "M"] * 100
    found = explain({"Tarife": tarife(), "Rechnung": rechnung(codes, wrong=50)}, "Rechnung", "D302")
    area = found["tree"]["children"][0]
    assert area["cell"] == "D2:D301" and "300 Zahlen" in area["value"] and "300 davon Formeln" in area["value"]
    assert [c["cell"] for c in area["children"]] == ["D50"]
    assert found["problems"] == ["C50"]


def test_an_input_shows_where_it_flows():
    found = explain({"Tarife": tarife(), "Rechnung": rechnung(["K", "A", "A"])}, "Tarife", "B3")
    top = found["tree"]
    assert top["kind"] == "eingabe" and top["value"] == "20 %"
    assert [u["cell"] for u in top["uses"]] == ["C2", "C3", "C4"] and top["uses"][0]["label"] == "Rechnung!C2 (Satz K)"


def test_circles_unknown_and_empty_cells_do_not_break_it():
    cells = {(1, 1): Cell("2", "B1+1"), (1, 2): Cell("1", "A1+1"), (2, 1): Cell("5", "TODAY()+1"), (3, 1): Cell("7", "C9+7")}
    top = explain({"S": cells}, "S", "A1")["tree"]
    again = [n for _, n in flat(top) if n.get("again")]
    assert again and again[0]["cell"] == "A1"
    assert explain({"S": cells}, "S", "A2")["tree"]["check"] == "unbekannt"
    leaf = explain({"S": cells}, "S", "A3")["tree"]["children"][0]
    assert leaf["kind"] == "leer" and leaf["cell"] == "C9"
    assert explain({"S": cells}, "S", "Z9")["state"] == "empty"
    assert explain({"S": cells}, "Fehlt", "A1")["state"] == "none"
    assert explain({"S": cells}, "S", "nicht")["state"] == "none"


def test_addin_explains_the_selected_cell(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    headers = {"X-VoiceAgent": "1"}
    assert client.post("/api/addin/herleitung", json={"sheet": "Wahl", "cell": "C7"}, headers=headers).json()["state"] == "again"
    client.post("/api/addin/pruefen", json=payload(), headers=headers)
    found = client.post("/api/addin/herleitung", json={"sheet": "Wahl", "cell": "$C$7"}, headers=headers).json()
    assert found["state"] == "ok" and found["cell"] == "C7" and found["tree"]["formula"] == "=SUM(C3:C5)"
    assert client.post("/api/addin/herleitung", json={"sheet": "Wahl", "cell": "C7"}).status_code == 403


def test_a_formula_says_in_words_which_rows_it_takes():
    from test_bereiche import table

    def words(formula):
        cells = table()
        cells[(2, 7)], cells[(2, 8)] = Cell("Kennzahl"), Cell("0", formula)
        return explain({"T": cells}, "T", "H2")["tree"]["rows"]

    assert words('SUMIF(B:B,"Nord",D:D)') == "Summiert Umsatz (€) aus den 10 Zeilen mit Region „Nord“ (D2:D11)."
    # Ein Bereich, der eine Zeile zu früh aufhört: man sieht es, auch ohne Namen daneben.
    assert words("SUM(D2:D10)") == "Summiert Umsatz (€) aus den 9 Zeilen mit Region „Nord“ außer D11 (D2:D10)."
    assert words('SUMIFS(D:D,A:A,">="&DATE(2026,2,1),A:A,"<"&DATE(2026,3,1))') == (
        "Summiert Umsatz (€) aus den 4 Zeilen mit Datum im Februar 2026 (D4:D5 und D12:D13).")
    assert words('COUNTIF(B:B,"Nord")') == "Zählt die 10 Zeilen mit Region „Nord“ (B2:B11)."
    assert words("SUM(D2:D11)-D11") == "Summiert Umsatz (€) aus den 9 Zeilen mit Region „Nord“ außer D11 (D2:D10)."
    assert words("SUM(D2:D10)+D5") == "Rechnet mit Umsatz (€) aus den 9 Zeilen mit Region „Nord“ außer D11 (D2:D10): D5 zählt doppelt."
    assert words("SUM(D2:D11)/9") == ("Teilt die Summe von Umsatz (€) aus den 10 Zeilen mit Region „Nord“ (D2:D11) durch die "
                                      "feste Zahl 9, es sind aber 10 Zeilen.")
    assert words("SUM(D:D)") == "Summiert Umsatz (€) aus allen 30 Zeilen der Tabelle (D2:D31)."


def test_rows_in_words_stay_quiet_where_the_formula_says_it_already():
    found = explain({"Tarife": tarife(), "Rechnung": rechnung(["K", "A", "M"])}, "Rechnung", "D5")
    assert found["tree"]["rows"] == "Summiert Total aus allen 3 Zeilen der Tabelle (D2:D4)."
    # Eine Formel in ihrer Zeile (=B2*C2) und eine einzelne Zelle brauchen keinen Satz.
    assert explain({"Tarife": tarife(), "Rechnung": rechnung(["K", "A", "M"])}, "Rechnung", "D2")["tree"]["rows"] == ""
