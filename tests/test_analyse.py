"""Tiefe Prüfung eines Blattes: eigener Formelrechner, Rechenkette, Sensitivität und Daten."""

import time

import pytest

from voiceagent import formel
from voiceagent.analyse import analyse, value_of
from voiceagent.formel import Err, Evaluator, parse
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Sheet, number_text


def calc(formula, values=None):
    values = values or {}
    return Evaluator(lambda p: values.get(p)).run(parse(formula))


def test_formulas_compute_like_excel():
    v = {(6, 3): 33.2, (7, 3): 28.0, (2, 4): 577.0, (6, 2): "RN", (7, 2): "NFP"}
    assert calc("C6/100*$D$2", v) == pytest.approx(191.564)
    assert calc("-2^2") == 4 and calc("2^3^2") == 64 and calc("1+2*3") == 7 and calc("10%") == 0.1
    assert calc('IF(C6>30,"gross","klein")', v) == "gross"
    assert calc("ROUND(2.5,0)") == 3 and calc("ROUND(-2.5,0)") == -3 and calc("ROUNDDOWN(1.99,1)") == 1.9
    assert calc("SUM(C6:C8)", v) == pytest.approx(61.2) and calc("AVERAGE(C6:C8)", v) == pytest.approx(30.6)
    assert calc('SUMIF(B6:B7,"NFP",C6:C7)', v) == 28 and calc('COUNTIF(C6:C7,">30")', v) == 1
    assert calc("1/0") == Err("#DIV/0!") and calc('IFERROR(1/0,"x")') == "x"
    assert calc('C6&" %"', v) == "33.2 %" and calc("SUMPRODUCT(C6:C7,C6:C7)", v) == pytest.approx(33.2**2 + 28**2)
    assert calc("STDEV.S(C6:C7)", v) == pytest.approx(3.676955, rel=1e-6)
    # Ohne die anderen Blätter und ohne Namen rechnet er nicht nach.
    for unknown in ("Tabelle2!A1+1", "Zinssatz*2", "[1]Daten!A1", "TODAY()+1"):
        node = parse(unknown)
        with pytest.raises(formel.Unknown):
            Evaluator(lambda p: None).eval(node)


def sheet(rows, extra=None, formula="C{r}/100*$D$2"):
    """Ein Blatt wie die Sitzverteilung der Übungsmappe, mit Werten, wie Excel sie zeigen würde."""
    cells = {(2, 3): Cell("Total Sitze:"), (2, 4): Cell("577"),
             (4, 2): Cell("Partei"), (4, 3): Cell("Wähleranteil (%)"), (4, 4): Cell("Sitze")}
    for r, (name, share) in enumerate(rows, start=6):
        cells[(r, 2)], cells[(r, 3)] = Cell(name), Cell(share)
        cells[(r, 4)] = Cell("", formula.format(r=r))
    cells.update(extra or {})
    for _ in range(4):
        ev = Evaluator(lambda p: value_of(cells.get(p)))
        for p, c in list(cells.items()):
            if c.formula:
                v = ev.run(parse(c.formula))
                cells[p] = Cell(number_text(v) if isinstance(v, float) else getattr(v, "code", str(v)), c.formula)
    return cells, layout(Sheet("Tabelle1", cells))


PARTIES = [("RN", "33,2"), ("NFP", "28"), ("Ensemble", "20,8"), ("LR", "6,6"), ("Rest", "11,4")]


def texts(report, level=None):
    return [f.text for f in report.findings if level is None or f.level == level]


def test_clean_sheet_has_no_findings_and_shows_what_drives_the_result():
    cells, plan = sheet(PARTIES, {(11, 4): Cell("", "SUM(D6:D10)")})
    report = analyse(cells, plan)
    assert texts(report) == []
    total = [d for d in report.drivers if d.output.startswith("D11")]
    # Alle Sitze hängen 1:1 an der Gesamtzahl, die Anteile je nach Größe.
    assert total[0].input == "D2 (Total Sitze)" and total[0].elasticity == pytest.approx(1.0)
    assert total[1].input.startswith("C6") and total[1].elasticity == pytest.approx(0.332, abs=0.01)
    assert report.stats == {"formeln": 6, "eingaben": 6, "ergebnisse": 1, "tiefe": 2, "nicht_nachrechenbar": 0}


def test_formula_out_of_pattern_is_found_with_what_it_should_be():
    cells, plan = sheet(PARTIES, {(9, 4): Cell("", "C9/100*577"), (11, 4): Cell("", "SUM(D6:D10)")})
    report = analyse(cells, plan)
    assert texts(report, "fehler") == [
        "D9 (Sitze LR) rechnet anders als die übrigen 4 Formeln in D6:D10: =Wähleranteil (%) / 100 * 577 "
        "statt =Wähleranteil (%) / 100 * Total Sitze ($D$2)."]
    # Und D9 hängt deshalb nicht an D2: Die Gesamtzahl wirkt nur noch zu 93 %.
    d2 = next(d for d in report.drivers if d.output.startswith("D11") and d.input.startswith("D2"))
    assert d2.elasticity == pytest.approx(0.93, abs=0.01)


def test_double_count_circle_empty_reference_and_recalculation():
    cells, plan = sheet(PARTIES, {
        (11, 4): Cell("", "SUM(D6:D7)"), (12, 4): Cell("", "SUM(D6:D11)"),
        (14, 4): Cell("", "D15+1"), (15, 4): Cell("", "D14*2"),
        (16, 4): Cell("", "C30*2"),
    })
    cells[(8, 4)] = Cell("999", "C8/100*$D$2")  # Excel zeigt etwas anderes (manuelle Berechnung)
    found = " ".join(texts(analyse(cells, plan)))
    assert "Die Summe in D12 (Sitze) zählt D11 mit, und D11 ist selbst schon die Summe von D6:D7" in found
    assert "Kreisbezug: D14, D15 hängen über Umwege von sich gegenseitig selbst ab." in found
    assert "D16 (Sitze) rechnet mit C30, die ist aber leer" in found
    assert "D8 (Sitze Ensemble) zeigt 999, nachgerechnet ergibt =Wähleranteil (%) / 100 * Total Sitze ($D$2) aber 120,016." in found


def test_data_checks_shares_mixed_scale_outliers_and_duplicates():
    cells, plan = sheet([("RN", "33,2"), ("NFP", "28"), ("Ensemble", "20,8"), ("LR", "6,6"), ("Rest", "11,3")])
    assert texts(analyse(cells, plan)) == ["Die Anteile in Wähleranteil (%) ergeben zusammen 99,9, nicht 100."]
    cells, plan = sheet([("RN", "33,2"), ("NFP", "0,28"), ("Ensemble", "20,8"), ("RN", "6,6"), ("Rest", "11,4")])
    found = texts(analyse(cells, plan))
    assert any("bei C7 aber als Bruch (0,28)" in t for t in found)
    assert any("„RN“ steht 2-mal in Partei (Zeilen 6, 9)" in t for t in found)
    many = [(f"P{i}", str(10 + i % 3)) for i in range(12)] + [("Tippfehler", "1100")]
    cells, plan = sheet(many, formula="C{r}*2")
    assert any(t.startswith("C18 (Wähleranteil (%) Tippfehler) = 1100 liegt weit weg") for t in texts(analyse(cells, plan)))
    assert not any("Code" in t for t in texts(analyse(cells, plan)))
    # Ein einzelnes 999 ist meist ein Code für „fehlt“: Das steht gleich dabei.
    cells = {(1, 1): Cell("Probe"), (1, 2): Cell("Wert"), (1, 3): Cell("Doppelt")}
    for r in range(2, 16):
        value = "999" if r == 9 else str(10 + r % 3)
        cells.update({(r, 1): Cell(f"L{r}"), (r, 2): Cell(value), (r, 3): Cell(str(2 * float(value)), f"B{r}*2")})
    code = next(f for f in analyse(cells, layout(Sheet("Labor", cells))).findings if f.key == "outlier:B9")
    assert code.text.endswith("999 sieht nach einem Code für „fehlt“ oder „unbekannt“ aus, aber MITTELWERT und SUMME rechnen damit.")
    assert code.tip.startswith("Ist es ein Code, lass die Zelle leer")


def test_code_among_widely_spread_counts():
    # Stückzahlen von 4 bis 122: Auf der Log-Skala ist 999 nur gut 8-mal das Größte, als Code für „fehlt“ aber
    # weit weg von allem. Eine echte große Menge (1100 unter 400 bis 900) bleibt still.
    cells = {(1, 1): Cell("Artikel"), (1, 2): Cell("Stück")}
    for r, value in enumerate(["4", "7", "27", "999", "49", "65", "999", "122"], start=2):
        cells.update({(r, 1): Cell(f"A{r}"), (r, 2): Cell(value)})
    found = {f.key: f for f in analyse(cells, layout(Sheet("Lager", cells))).findings}
    assert {"outlier:B5", "outlier:B8"} <= set(found)
    assert "999 sieht nach einem Code für „fehlt“" in found["outlier:B5"].text
    for r, value in enumerate(["400", "700", "270", "1100", "490", "650", "900", "520"], start=2):
        cells[(r, 2)] = Cell(value)
    assert not any(k.startswith(("outlier:", "scale:")) for k in
                   {f.key for f in analyse(cells, layout(Sheet("Lager", cells))).findings})


def test_large_sheet_stays_fast():
    rows = [(f"Fall {i}", str(i % 17 + 1)) for i in range(2000)]
    cells, plan = sheet(rows, {(2010, 4): Cell("", "SUM(D6:D2005)")})
    started = time.monotonic()
    report = analyse(cells, plan)
    assert time.monotonic() - started < 6
    assert report.stats["formeln"] == 2001


def test_running_total_over_thousands_of_rows_does_not_crash():
    cells, plan = sheet([(f"Tag {i}", "1") for i in range(3000)], formula="C{r}")
    for r in range(7, 3006):
        cells[(r, 4)] = Cell(str(r - 5), f"D{r - 1}+C{r}")
    report = analyse(cells, plan)
    assert report.stats["tiefe"] == 3000 and texts(report, "fehler") == []


def test_no_false_alarms_for_subtotals_guarded_blanks_and_rounded_file_values():
    # Eine Zwischensumme mitten in der Spalte ist kein Kopierfehler, und WENN(C30="";…) fragt bewusst nach Leerem.
    cells, plan = sheet(PARTIES, {(9, 4): Cell("", "SUM(D6:D8)"), (13, 4): Cell("", 'IF(C30="","",C30*2)')})
    assert texts(analyse(cells, plan)) == []
    # Aus der Datei sind Werte gerundet wie angezeigt: dann kein „Excel zeigt … nachgerechnet …“.
    cells, plan = sheet(PARTIES)
    cells[(6, 4)] = Cell("191,56", "C6/100*$D$2")
    assert any("zeigt 191,56" in t for t in texts(analyse(cells, plan)))
    assert not any("zeigt 191,56" in t for t in texts(analyse(cells, plan, exact=False)))
