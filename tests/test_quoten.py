"""Quoten, Mittelwerte und Wachstum, die falsch gewichtet oder auf die falsche Basis gerechnet sind."""

from test_bereiche import ROWS, figure, hints, table

from voiceagent.xlsx import Cell

# Erfundene Kliniken: Fälle, Komplikationen, Quote je Zeile.
KLINIKEN = [("A", 12, 3), ("B", 900, 45), ("C", 150, 12), ("D", 40, 6), ("E", 600, 30)]


def kliniken(rows=KLINIKEN):
    cells = {(1, 1): Cell("Klinik"), (1, 2): Cell("Fälle"), (1, 3): Cell("Komplikationen"), (1, 4): Cell("Quote")}
    for r, (name, cases, events) in enumerate(rows, start=2):
        cells[(r, 1)], cells[(r, 2)], cells[(r, 3)] = Cell(name), Cell(str(cases)), Cell(str(events))
        cells[(r, 4)] = Cell(f"{events / cases * 100:.6f} %".replace(".", ","), f"C{r}/B{r}")
    return cells


def test_a_mean_of_rates_is_not_the_rate_over_all_rows():
    cells = kliniken()
    cells[(8, 3)], cells[(8, 4)] = Cell("Komplikationsrate"), Cell("0", "AVERAGE(D2:D6)")
    found = hints(cells)["ratio"]
    assert ("D8 (Komplikationsrate) mittelt die Quoten D2:D6 (Quote) ungewichtet: Jede Zeile zählt gleich viel, ob "
            "Fälle 12 ist (Zeile 2) oder 900 (Zeile 3). Über alle Zeilen zusammen, Summe Komplikationen durch Summe "
            "Fälle, ist die Quote 5,64 % statt 11,6 %.") == found.text
    assert found.level == "pruefen" and "=SUMME(C2:C6)/SUMME(B2:B6), nachgerechnet 5,64 %" in found.tip
    # Richtig über alle Zeilen, ausdrücklich ungewichtet, oder alle Nenner gleich: nichts.
    cells = kliniken()
    cells[(8, 3)], cells[(8, 4)] = Cell("Komplikationsrate"), Cell("0", "SUM(C2:C6)/SUM(B2:B6)")
    assert "ratio" not in hints(cells)
    cells = kliniken()
    cells[(8, 3)], cells[(8, 4)] = Cell("Quote ungewichtet"), Cell("0", "AVERAGE(D2:D6)")
    assert "ratio" not in hints(cells)
    cells = kliniken([(name, 100, events) for name, _, events in KLINIKEN])
    cells[(8, 3)], cells[(8, 4)] = Cell("Komplikationsrate"), Cell("0", "AVERAGE(D2:D6)")
    assert "ratio" not in hints(cells)


def test_a_mean_of_group_means_with_different_sizes():
    nord = [r[3] for r in ROWS if r[1] == "Nord"]
    ost = [r[3] for r in ROWS if r[1] == "Ost"]
    for formula in ["AVERAGE(H2,H3)", "(H2+H3)/2"]:
        cells = figure(table(), "Schnitt Nord", sum(nord) / 10, 'AVERAGEIF(B:B,"Nord",D:D)')
        figure(cells, "Schnitt Ost", sum(ost) / 8, 'AVERAGEIF(B:B,"Ost",D:D)', row=3)
        figure(cells, "Schnitt Nord und Ost", 0, formula, row=4)
        found = hints(cells)
        assert ("H4 (Schnitt Nord und Ost) mittelt die Mittelwerte H2 (Schnitt Nord, 10 Zeilen) und H3 (Schnitt Ost, "
                "8 Zeilen) zu gleichen Teilen. Über alle 18 Zeilen ist der Mittelwert 2502,935 statt 2482,1898") in found["means"].text
        assert found["means"].level == "pruefen" and "nachgerechnet 2502,935" in found["means"].tip
        assert "label" not in found, formula


def growth(formula, heads=("Fälle 2025", "Fälle 2026", "Wachstum")):
    cells = {(1, 1): Cell("Station"), (1, 2): Cell(heads[0]), (1, 3): Cell(heads[1]), (1, 4): Cell(heads[2])}
    for r, (name, old, new) in enumerate([("A", 100, 125), ("B", 80, 60), ("C", 200, 210), ("D", 50, 75)], start=2):
        cells[(r, 1)], cells[(r, 2)], cells[(r, 3)] = Cell(name), Cell(str(old)), Cell(str(new))
        cells[(r, 4)] = Cell("0", formula.replace("2", str(r)))
    return hints(cells)


def test_growth_on_the_wrong_base_or_the_wrong_way_round():
    found = growth("(C2-B2)/C2")["growth"]
    assert ("D2 (Wachstum) teilt durch C2 (Fälle 2026), den neuen Wert. Wachstum misst man gegenüber dem alten Wert B2 "
            "(Fälle 2025). Richtig wäre 0,25 statt 0,2. Das gilt für alle 4 Zellen D2:D5.") == found.text
    assert found.level == "fehler" and "Schreib in D2 =(C2-B2)/B2 und zieh es herunter, nachgerechnet 0,25." == found.tip
    found = growth("(B2-C2)/B2", ("Vorjahr", "Aktuell", "Veränderung"))["growth"]
    assert "rechnet B2 − C2, also alt minus neu: Das Vorzeichen ist umgedreht. Richtig wäre 0,25 statt -0,25" in found.text
    # Richtig, als Verhältnis minus 1, ein Rückgang, oder eine Marge (Preis − Kosten) / Preis: nichts.
    assert "growth" not in growth("(C2-B2)/B2") and "growth" not in growth("C2/B2-1")
    assert "growth" not in growth("(B2-C2)/B2", ("Fälle 2025", "Fälle 2026", "Rückgang"))
    assert "growth" not in growth("(C2-B2)/C2", ("Kosten", "Preis", "Marge"))
