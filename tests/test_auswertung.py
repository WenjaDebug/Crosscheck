"""Prüfungen für Datenauswertungen: fehlende Werte, Codes, unmögliche Werte, Dubletten, Einheiten, Bereiche."""

import random
from datetime import date, timedelta

from voiceagent.analyse import Sheet as AnalyseSheet, analyse
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Sheet, column_number

HEADERS = ["Fall-Nr", "Geburtsdatum", "Eintritt", "Austritt", "Alter", "Gewicht kg", "Größe", "Liegedauer", "Station"]
N = 40


def day(d: date) -> str:
    return d.strftime("%d.%m.%Y")


def faelle(**planted):
    """Eine erfundene Fallliste wie aus dem Spital, ohne echte Daten. planted: welche Fehler drin sein sollen."""
    rnd = random.Random(7)
    rows = []
    for i in range(N):
        start = date(2024, 1, 1) + timedelta(days=rnd.randint(0, 300))
        stay = rnd.randint(1, 14)
        age = rnd.randint(18, 90)
        rows.append({
            "Fall-Nr": str(240100 + i), "Geburtsdatum": day(date(2024 - age, rnd.randint(1, 12), rnd.randint(1, 28))),
            "Eintritt": day(start), "Austritt": day(start + timedelta(days=stay)), "Alter": str(age),
            "Gewicht kg": str(rnd.randint(50, 110)), "Größe": str(rnd.randint(150, 195)), "Station": rnd.choice("ABC"),
        })
    if planted.get("dupe_id"):
        rows[20]["Fall-Nr"] = rows[5]["Fall-Nr"]
    if planted.get("same_row"):
        rows[30] = dict(rows[12])
    if planted.get("code"):
        rows[3]["Alter"] = rows[17]["Alter"] = "999"
    if planted.get("impossible"):
        rows[8]["Alter"] = "-4"
    if planted.get("order"):
        rows[9]["Austritt"] = day(date(2023, 12, 1))
    if planted.get("units"):
        for i in (2, 11, 25):
            rows[i]["Größe"] = f"1,{rnd.randint(55, 90)}"
    if planted.get("missing"):
        for i in (4, 14, 24):
            rows[i]["Gewicht kg"] = ""
        rows[33]["Gewicht kg"] = "k.A."
    cells = {(1, c): Cell(h) for c, h in enumerate(HEADERS, start=1)}
    for r, row in enumerate(rows, start=2):
        for c, h in enumerate(HEADERS, start=1):
            if h == "Liegedauer":
                stay = (date(*map(int, reversed(row["Austritt"].split(".")))) - date(*map(int, reversed(row["Eintritt"].split("."))))).days
                cells[(r, c)] = Cell(str(stay), f"D{r}-C{r}")
            elif row[h]:
                cells[(r, c)] = Cell(row[h])
    return cells


def report(cells, filtered=None):
    return analyse(cells, layout(Sheet("Fälle", cells)), name="Fälle", filtered=filtered)


def kinds(rep):
    return {f.key.split(":")[0]: f for f in rep.findings}


def texts(rep, kind):
    return " | ".join(f.text for f in rep.findings if f.key.startswith(f"{kind}:"))


def test_a_clean_case_list_raises_no_analysis_alarm():
    found = kinds(report(faelle()))
    assert not {"missing", "placeholder", "code", "impossible", "order", "future", "id", "same", "units", "range",
                "total", "slide", "filter", "outlier"} & set(found), found


def test_planted_mistakes_are_found():
    rep = report(faelle(dupe_id=True, same_row=True, code=True, impossible=True, order=True, units=True, missing=True))
    found = kinds(rep)
    assert "In Alter steht 2-mal 999 (E5, E19)" in found["code"].text and found["code"].level == "fehler"
    impossible = texts(rep, "impossible")
    assert "-4 in E10: Ein Alter unter 0 oder über 120 Jahren gibt es nicht" in impossible
    # Die Liegedauer aus Austritt minus Eintritt wird dort negativ: Das sagt der Hinweis zur Ursache, nicht ein zweiter.
    assert "Liegedauer" not in impossible
    assert "Austritt liegt vor Eintritt (Zeile 11" in found["order"].text
    assert found["order"].text.endswith("Darum zeigt H11 (Liegedauer, Fall-Nr 240109) -72.")
    assert "Tag und Monat vertauscht" in found["order"].tip
    # 240105 steht in Zeile 22, wo 240120 fehlt: kopiert. Zeile 32 ist ganz gleich wie Zeile 14: doppelt eingefügt.
    ids = texts(rep, "id")
    assert "A22 (Fall-Nr) „240105“ steht schon in A7" in ids and "„240120“ fehlt" in ids
    assert "„240112“ in Zeilen 14, 32" in ids
    assert "Zeile 32 ist genau gleich wie Zeile 14" in found["same"].text
    assert "Faktor 100" in found["units"].text and "(m und cm)" in found["units"].text and "G4" in found["units"].text
    assert "Gewicht kg 3 von 40 (8 %, F6, F16, F26)" in found["missing"].text
    assert "„k.A.“" in found["placeholder"].text and "F35" in found["placeholder"].text
    # Die 999 ist ein Code, kein Ausreißer, und keine unmögliche Zahl.
    assert "outlier" not in found and "999" not in impossible and "text" not in found
    assert "recalc" not in found


def test_average_over_a_range_that_stops_early():
    cells = faelle()
    last = N + 1
    cells[(last + 2, 1)] = Cell("Mittelwert")
    cells[(last + 2, 5)] = Cell("50", f"AVERAGE(E2:E{last - 3})")
    cells[(last + 3, 1)] = Cell("Median")
    cells[(last + 3, 5)] = Cell("50", f"MEDIAN(E2:E{last})")
    found = [f for f in report(cells).findings if f.key.startswith("range:")]
    assert len(found) == 1 and "MITTELWERT in E43" in found[0].text and "3 Zeilen fehlen (E39, E40, E41)" in found[0].text


def test_sliding_denominator_and_filtered_sums():
    cells = {(1, 1): Cell("Station"), (1, 2): Cell("Fälle"), (1, 3): Cell("Anteil")}
    for r, (name, n) in enumerate([("A", 10), ("B", 20), ("C", 30), ("D", 40)], start=2):
        cells[(r, 1)], cells[(r, 2)] = Cell(name), Cell(str(n))
        cells[(r, 3)] = Cell("0", f"B{r}/B{r + 4}")
    cells[(6, 1)], cells[(6, 2)] = Cell("Summe"), Cell("100", "SUM(B2:B5)")
    found = kinds(report(cells))
    assert "C2 teilt durch B6 (die Summe), C3 aber schon durch B7" in found["slide"].text and "B$6" in found["slide"].tip
    found = kinds(report(cells, filtered=[(1, 1, 5, 3)]))
    assert "ist gefiltert" in found["filter"].text and "TEILERGEBNIS(109" in found["filter"].tip
    cells[(6, 2)] = Cell("100", "SUBTOTAL(109,B2:B5)")
    assert "filter" not in kinds(report(cells, filtered=[(1, 1, 5, 3)]))


def test_codes_need_repeats_and_distance():
    cells = faelle()
    col = HEADERS.index("Alter") + 1
    cells[(5, col)] = Cell("99")  # einmal 99: ein echtes Alter
    assert "code" not in kinds(report(cells))
    assert column_number("E") == col


def _clean_table(seed: int):
    """Erfundene, aber stimmige Daten: schiefe Laborwerte und Kosten, Prozente, Veränderungen, Umfragen."""
    rnd = random.Random(seed)
    n = rnd.choice([12, 30, 80, 200])
    kind = seed % 4
    if kind == 0:
        headers = ["Patienten-ID", "Station", "Alter", "CRP mg/l", "Kosten CHF", "Liegedauer", "Geschlecht"]

        def gen(i):
            return [f"P{10000 + i}", rnd.choice(["A", "B", "C", "IPS"]), rnd.randint(0, 99), round(rnd.lognormvariate(2, 1.2), 1),
                    round(rnd.lognormvariate(8, 1.5)), rnd.randint(0, 40), rnd.choice("mw")]
    elif kind == 1:
        headers = ["Monat", "Umsatz", "Kosten", "Anteil %", "Veränderung %"]

        def gen(i):
            return [day(date(2020, 1, 1) + timedelta(days=31 * i)), rnd.randint(1000, 90000), rnd.randint(500, 50000),
                    round(rnd.uniform(5, 95), 1), round(rnd.gauss(0, 40), 1)]
    elif kind == 2:
        headers = ["Frage 1", "Frage 2", "Frage 3", "Frage 4", "Gruppe"]

        def gen(i):
            return [rnd.randint(1, 5), rnd.randint(1, 5), rnd.randint(1, 5), rnd.randint(1, 5), rnd.choice("AB")]
    else:
        headers = ["Fall", "Eintritt", "Austritt", "Gewicht kg", "Größe cm", "BMI", "Glukose"]

        def gen(i):
            s, w, h = date(2024, 1, 1) + timedelta(days=rnd.randint(0, 300)), rnd.randint(40, 140), rnd.randint(145, 200)
            return [str(1000 + i), day(s), day(s + timedelta(days=rnd.randint(0, 30))), w, h, round(w / (h / 100) ** 2, 1),
                    round(rnd.lognormvariate(1.7, 0.3), 1)]
    cells = {(1, c): Cell(h) for c, h in enumerate(headers, start=1)}
    for r in range(2, n + 2):
        for c, v in enumerate(gen(r), start=1):
            cells[(r, c)] = Cell(f"{v:.2f}".replace(".", ",") if isinstance(v, float) else str(v))
    return cells


def test_no_false_alarms_on_plausible_data():
    for seed in range(40):
        found = [f for f in report(_clean_table(seed)).findings]
        assert not found, (seed, [f.text for f in found])


def test_filtered_table_from_the_addin(tmp_path):
    from test_excel_live import FakeExcel, client_with

    _, client = client_with(tmp_path, FakeExcel(None))
    values = [["Station", "Fälle"]] + [[s, n] for s, n in zip("ABCDEFGHIJ", range(10, 110, 10))] + [["Summe", 550]]
    formulas = [[None, None]] * 11 + [[None, "=SUM(B2:B11)"]]
    sheet = {"name": "Fälle", "row": 1, "col": 1, "values": values, "formulas": formulas, "filtered": ["Fälle!$A$1:$B$11"]}
    found = client.post("/api/addin/pruefen", json={"book": "F.xlsx", "active": "Fälle", "sheets": [sheet]},
                        headers={"X-VoiceAgent": "1"}).json()
    assert any(h["key"] == "filter:B12" and "gefiltert" in h["text"] for h in found["hints"])
    sheet["filtered"] = []
    found = client.post("/api/addin/pruefen", json={"book": "F.xlsx", "active": "Fälle", "sheets": [sheet]},
                        headers={"X-VoiceAgent": "1"}).json()
    assert not any(h["key"].startswith("filter:") for h in found["hints"])


def test_practice_workbook_shows_every_planted_mistake():
    """web/demo/pruefer-uebung.xlsx: erfundene Fälle, Tarife, eine Rechnung, Statistik und Zeiträume mit eingebauten
    Fehlern."""
    from pathlib import Path

    from voiceagent.xlsx import read_book

    book = read_book((Path(__file__).parent.parent / "src/voiceagent/web/demo/pruefer-uebung.xlsx").read_bytes())
    cells_of = {s.name: s.cells for s in book.sheets}
    found, texts = {}, {}
    for sheet in book.sheets:
        rep = analyse(sheet.cells, layout(sheet), name=sheet.name, book=cells_of)
        found[sheet.name] = {f.key.split(":")[0] for f in rep.findings}
        texts |= {f.key: f.text for f in rep.findings}
    # Der zu kurze Bereich steckt im Hinweis zu den Codes im selben MITTELWERT, mit einer Korrektur für beides.
    assert {"code", "order", "statcode", "placeholder", "units", "id", "same", "missing"} <= found["Fälle"]
    # Die negative Liegedauer steht beim Austritt vor dem Eintritt, nicht als eigener unmöglicher Wert.
    assert "impossible" not in found["Fälle"] and "Darum zeigt H11 (Liegedauer" in texts["order:D11"]
    assert "Zahlen aber bis E41" in texts["statcode:L2"]
    assert {"slide", "hidden", "twice"} <= found["Rechnung"] and "emptyref" not in found["Rechnung"]
    assert "outlier" not in found["Fälle"] and "round" not in found["Rechnung"]
    # Tests und Intervalle: gepaart über zwei Gruppen, ungepaart über vorher und nachher, KONFIDENZ mit 0,95, ein
    # Intervall ohne Wurzel aus n und Prozentpunkte unter „Veränderung %“.
    assert {"ttest:H4", "ttest:H5", "confidence:H6", "interval:H7", "points:J12"} <= {
        k for k in texts if k.split(":")[0] in ("ttest", "confidence", "interval", "points")}
    # Zeiträume: Alter durch 365, Tage „inklusive“ ohne +1, Arbeitstage mit Wochenenden, Alter bei Eintritt bis HEUTE().
    assert {"age", "days", "today"} <= found["Zeiträume"] and "constant" not in found["Zeiträume"]
    assert {"age:E2", "days:F2", "days:G2", "today:H2"} <= set(texts)


def test_drivers_name_the_figure_and_skip_codes():
    cells = faelle(code=True)
    ages = [float(cells[(r, 5)].value) for r in range(2, N + 2)]
    cells[(2, 11)] = Cell("Mittelwert Alter")
    cells[(2, 12)] = Cell(f"{sum(ages) / len(ages):.10g}", f"AVERAGE(E2:E{N + 1})")
    rep = report(cells)
    drivers = [d for d in rep.drivers if d.output.startswith("L2")]
    assert drivers and all(d.output == "L2 (Mittelwert Alter)" for d in drivers)
    # Eine Formel, die nur in ihrer Zeile rechnet, heißt weiter nach Spalte und Fall.
    cells[(3, 7)] = Cell("gross")
    sheet = AnalyseSheet(cells, layout(Sheet("Fälle", cells)))
    assert sheet.figure_label((2, 12)) == "L2 (Mittelwert Alter)"
    assert sheet.figure_label((3, 8)) == "H3 (Liegedauer, Fall-Nr 240101)"
    assert all(d.input_value != "999" for d in drivers)


def test_the_code_999_can_be_cleared_with_one_click():
    """Einsetzen macht die Zellen mit 999 leer: MITTELWERT und SUMME lassen leere Zellen aus. Danach ist der Hinweis
    weg, und Rückgängig weiß, was dastand."""
    from voiceagent import einsetzen

    cells = faelle(code=True)
    cells[(1, 11)], cells[(2, 11)] = Cell("Mittel Alter"), Cell("0", "AVERAGE(E2:E41)")
    fix = kinds(report(cells))["code"].fix
    assert fix["clear"] and fix["cells"] == {"E5": "", "E19": ""} and fix["was"] == {"E5": {"text": "999"}, "E19": {"text": "999"}}
    assert fix["range"] == "E5 und E19"
    after = einsetzen.applied(cells, [fix])
    assert (5, 5) not in after and (19, 5) not in after
    assert "code" not in kinds(report(after))


def test_no_clearing_when_a_formula_reads_the_cell_by_itself():
    """=E5*12 rechnete mit der leeren Zelle als 0: wieder falsch, nur leiser. Dann bleibt es beim Tipp."""
    cells = faelle(code=True)
    cells[(1, 11)], cells[(5, 11)] = Cell("Monate"), Cell("11988", "E5*12")
    assert kinds(report(cells))["code"].fix is None
