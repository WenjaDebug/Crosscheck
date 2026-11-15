"""Bereiche, die nicht zu ihrer Kennzahl passen: „Umsatz alle Nord“, SUMMEWENN verschoben, vertippte Bedingungen."""

import random

from voiceagent.analyse import analyse
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Sheet

# Wie die Übungsmappe (erfundene Zahlen): Datum, Region, Produkt, Umsatz, nach Region sortiert.
ROWS = [
    ("12.01.2026", "Nord", "Tastatur", 2630.64), ("19.01.2026", "Nord", "Dock", 3029.22),
    ("09.02.2026", "Nord", "Maus", 958.29), ("23.02.2026", "Nord", "Dock", 4855.45),
    ("02.03.2026", "Nord", "Tastatur", 4787.37), ("09.03.2026", "Nord", "Tastatur", 2841.50),
    ("13.04.2026", "Nord", "Dock", 1345.51), ("15.06.2026", "Nord", "Maus", 832.74),
    ("13.07.2026", "Nord", "Monitor", 1915.95), ("20.07.2026", "Nord", "Monitor", 3492.30),
    ("02.02.2026", "Ost", "Laptop", 2593.05), ("16.02.2026", "Ost", "Tastatur", 1045.76),
    ("16.03.2026", "Ost", "Monitor", 1017.86), ("30.03.2026", "Ost", "Monitor", 937.24),
    ("27.04.2026", "Ost", "Tastatur", 3842.67), ("04.05.2026", "Ost", "Laptop", 1274.59),
    ("06.07.2026", "Ost", "Maus", 2861.81), ("27.07.2026", "Ost", "Monitor", 4790.88),
    ("05.01.2026", "Süd", "Tastatur", 754.28), ("26.01.2026", "Süd", "Monitor", 3950.02),
    ("11.05.2026", "Süd", "Laptop", 428.72), ("18.05.2026", "Süd", "Maus", 3305.08),
    ("01.06.2026", "Süd", "Tastatur", 4224.47), ("08.06.2026", "Süd", "Tastatur", 1467.49),
    ("22.06.2026", "Süd", "Maus", 4320.21), ("23.03.2026", "West", "Tastatur", 1316.44),
    ("06.04.2026", "West", "Monitor", 1880.27), ("20.04.2026", "West", "Tastatur", 1332.98),
    ("25.05.2026", "West", "Monitor", 2303.66), ("29.06.2026", "West", "Dock", 4329.90),
]
NORD = 26688.97


def table(rows=ROWS):
    cells = {(1, 1): Cell("Datum"), (1, 2): Cell("Region"), (1, 3): Cell("Produkt"), (1, 4): Cell("Umsatz (€)")}
    for r, (day, region, product, amount) in enumerate(rows, start=2):
        # Leerzeichen am Rand behält das Lesen eigens, wie aus Excel.
        exact = region if region != region.strip() else None
        cells[(r, 1)], cells[(r, 2)] = Cell(day), Cell(region.strip(), text=exact)
        cells[(r, 3)], cells[(r, 4)] = Cell(product), Cell(str(amount))
    return cells


def figure(cells, label, value, formula=None, row=2):
    cells[(row, 7)] = Cell(label)
    cells[(row, 8)] = Cell(str(value), formula)
    return cells


def hints(cells):
    return {f.key.split(":")[0]: f for f in analyse(cells, layout(Sheet("Umsatz", cells)), name="Umsatz").findings}


def test_a_named_sum_that_misses_a_row_of_its_group():
    found = hints(figure(table(), "Umsatz alle Nord:", 23196.67, "SUM(D2:D10)"))
    text = found["label"].text
    assert "„Umsatz alle Nord“ in H2 summiert D2:D10, aber der Bereich D2:D10 hört eine Zeile zu früh auf: D11 (20.07.2026) gehört noch dazu." in text
    assert "Richtig wäre 26688,97 statt 23196,67" in text
    assert 'SUMMEWENN(B:B;"Nord";D:D)' in found["label"].tip
    # Derselbe Fehler, nicht noch einmal als Gruppe oder Bereich.
    assert "group" not in found and "range" not in found


def test_also_in_an_unsorted_table_and_with_a_row_too_many():
    rows = ROWS[:]
    random.Random(3).shuffle(rows)
    cells = table(rows)
    nord = [r for r, row in enumerate(rows, start=2) if row[1] == "Nord"]
    text = hints(figure(cells, "Umsatz Nord", 1, f"SUM(D2:D{nord[-1]})"))["label"].text
    assert "gehören nicht dazu" in text and "fehlt" not in text
    cells = table()
    text = hints(figure(cells, "Summe Nord", 1, "SUM(D2:D12)"))["label"].text
    assert "der Bereich D2:D12 reicht eine Zeile zu weit: D12 (Ost) gehört nicht mehr dazu." in text


def test_right_ranges_and_unknown_words_stay_quiet():
    for label, formula, value in [("Umsatz alle Nord", "SUM(D2:D11)", NORD), ("Umsatz Nord", 'SUMIF(B2:B31,"Nord",D2:D31)', NORD),
                                  ("Ziel Umsatz Nord", "SUM(D2:D10)", 1), ("Provision Nord", "SUM(D2:D10)*0.1", 1),
                                  ("Umsatz Ost", "SUM(D12:D19)", 1), ("Mittelwert Süd", "AVERAGE(D20:D26)", 1),
                                  ("Umsatz Nord in %", 'SUMIF(B2:B31,"Nord",D2:D31)/74666.4', 0.357)]:
        found = hints(figure(table(), label, value, formula))
        assert not {"label", "shift", "partial", "criterion"} & set(found), (label, found)


def test_a_typed_number_that_no_longer_fits_its_rows():
    found = hints(figure(table(), "Umsatz alle Nord", 23196.67))
    assert "steht „Umsatz alle Nord“ als feste Zahl 23196,67" in found["label"].text
    assert "Mit Region „Nord“ gibt es 10 Zeilen, nachgerechnet ist die Summe 26688,97. Es fehlt genau D11 (20.07.2026)." in found["label"].text
    assert found["label"].level == "pruefen"
    # Richtig, oder auf ganze Franken gerundet: nichts zu sagen. Und eine Zahl ohne Spaltenkopf im Namen ist eher ein Satz.
    for value, label in [(NORD, "Umsatz alle Nord"), (26689, "Umsatz Nord"), (0.15, "Nord")]:
        assert "label" not in hints(figure(table(), label, value)), (value, label)


def test_months_quarters_and_totals():
    text = hints(figure(table(), "Umsatz März 2026", 1, "SUM(D6:D7)"))["label"].text
    assert "Datum im März 2026" in text and "fehlen" in text and "D15" in text
    found = hints(figure(table(), "Umsatz März 2026", 1, "SUM(D6:D7)"))
    assert 'SUMMEWENNS(D:D;A:A;">="&DATUM(2026;3;1);A:A;"<"&DATUM(2026;4;1))' in found["label"].tip
    text = hints(figure(table(), "Total Umsatz", 1, "SUM(D2:D20)"))["label"].text
    assert "summiert D2:D20, aber der Bereich D2:D20 hört 11 Zeilen zu früh auf: D21 (26.01.2026), D22 (11.05.2026), D23 (18.05.2026) und 8 weitere gehören noch dazu." in text
    assert "label" not in hints(figure(table(), "Total Umsatz", 1, "SUM(D2:D31)"))


def test_sumif_with_shifted_or_partial_ranges():
    text = hints(figure(table(), "Umsatz Nord", 1, 'SUMIF(B2:B31,"Nord",D3:D32)'))["shift"].text
    assert "B2:B31 und D3:D32 fangen nicht in derselben Zeile an" in text and "1 Zeile darunter" in text
    # Ganze Spalte für die Zahlen: Excel nimmt dann D1:D30, eine Zeile verschoben.
    assert "shift" in hints(figure(table(), "Umsatz Nord", 1, 'SUMIF(B2:B31,"Nord",D:D)'))
    found = hints(figure(table(), "Umsatz", 1, 'SUMIF(B2:B9,"Nord",D2:D9)'))
    assert "prüft nur B2:B9, die Tabelle geht aber bis Zeile 31: 2 Zeilen erfüllen" in found["partial"].text
    assert "partial" not in hints(figure(table(), "Umsatz", 1, 'SUMIF(B2:B11,"Nord",D2:D11)'))


def test_conditions_that_miss_what_the_table_says():
    rows = [(d, f"{region} " if i in (3, 7) else region, p, a) for i, (d, region, p, a) in enumerate(ROWS)]
    text = hints(figure(table(rows), "Umsatz", 1, 'SUMIF(B2:B31,"Nord",D2:D31)'))["criterion"].text
    assert "sucht „Nord“, in B5, B9 steht aber „Nord “ mit einem Leerzeichen am Ende: 2 Zeilen zählen nicht mit" in text
    found = hints(figure(table(), "Umsatz", 1, 'SUMIFS(D2:D31,B2:B31,"Nrd")'))
    assert "„Nrd“, das kommt in B2:B31 nicht vor" in found["criterion"].text and "„Nord“" in found["criterion"].text
    cells = figure(table(), "Umsatz", 1, "SUMIF(B2:B31,G5,D2:D31)")
    cells[(5, 7)] = Cell("Nord", text="Nord ")
    assert "mit Leerzeichen" in hints(cells)["criterion"].text
    assert "criterion" not in hints(figure(table(), "Umsatz", 1, 'SUMIF(B2:B31,"N*",D2:D31)'))


def test_a_range_that_stops_one_row_before_the_end_of_its_group_without_a_name():
    cells = table()
    cells[(2, 8)] = Cell("23196.67", "SUM(D2:D10)")
    found = hints(cells)["group"]
    assert "SUMME in H2 nimmt D2:D10" in found.text and "„Nord“ aber in den Zeilen 2 bis 11" in found.text
    assert "1 Zeile mit „Nord“ fehlt (D11)" in found.text and 'SUMMEWENN(B:B;"Nord";D:D)' in found.tip
    cells[(2, 8)] = Cell("1", "SUM(D2:D12)")
    assert "D12 gehört schon nicht mehr zu „Nord“" in hints(cells)["group"].text
    for formula in ("SUM(D2:D11)", "SUM(D12:D19)", "SUM(D5:D8)"):
        cells[(2, 8)] = Cell("1", formula)
        assert "group" not in hints(cells), formula


def test_an_average_that_divides_by_a_fixed_count():
    # Wie in der Übungsmappe: H2 richtig mit SUMMEWENN, H3 teilt durch 9, Nord hat aber 10 Zeilen.
    cells = figure(table(), "Umsatz alle Nord:", NORD, 'SUMIF(B:B,"Nord",D:D)')
    figure(cells, "Durchschnitt Umsatz Nord", NORD / 9, "H2/9", row=3)
    found = hints(cells)
    assert ("„Durchschnitt Umsatz Nord“ in H3 teilt durch 9. Mit Region „Nord“ gibt es aber 10 Zeilen. "
            "Richtig wäre 2668,897 statt 2965,4411.") == found["label"].text
    assert found["label"].level == "fehler" and "Korrigiert: =H2/10, nachgerechnet 2668,897." in found["label"].tip
    assert 'Sicherer, auch wenn Zeilen dazukommen oder neu sortiert wird: =MITTELWERTWENN(B:B;"Nord";D:D).' in found["label"].tip
    assert "count" not in found and "shift" not in found
    # Ohne Namen sagt es die feste Anzahl, direkt mit SUMME, dann nur zum Prüfen.
    cells = table()
    cells[(3, 8)] = Cell("1", "SUM(D2:D11)/9")
    assert hints(cells)["count"].level == "pruefen" and "MITTELWERT(D2:D11)" in hints(cells)["count"].tip
    # Richtig geteilt, durch ZÄHLENWENN oder mit MITTELWERTWENN: nichts.
    for formula, value in [("H2/10", NORD / 10), ('H2/COUNTIF(B:B,"Nord")', NORD / 10),
                           ('AVERAGEIF(B2:B31,"Nord",D2:D31)', NORD / 10)]:
        cells = figure(table(), "Umsatz alle Nord", NORD, 'SUMIF(B:B,"Nord",D:D)')
        figure(cells, "Durchschnitt Umsatz Nord", value, formula, row=3)
        assert not {"count", "label"} & set(hints(cells)), formula
    # Durch 100 ist kein Zählfehler; dass das kein Durchschnitt ist, sagt der Name.
    cells = figure(table(), "Umsatz alle Nord", NORD, 'SUMIF(B:B,"Nord",D:D)')
    figure(cells, "Durchschnitt Umsatz Nord", NORD / 100, "H2/100", row=3)
    assert "count" not in hints(cells) and "teilt durch 100. Mit Region „Nord“ gibt es aber 10 Zeilen." in hints(cells)["label"].text


def test_any_formula_is_checked_against_its_name():
    # Mittelwert über einen Bereich mit einer Zeile zu wenig, und eine Formel, die etwas anderes zählt.
    cells = figure(table(), "Durchschnitt Umsatz Nord", 2577.408, "AVERAGE(D2:D10)", row=3)
    assert "mittelt D2:D10, aber der Bereich D2:D10 hört eine Zeile zu früh auf: D11 (20.07.2026) gehört noch dazu." in hints(cells)["label"].text
    cells = figure(table(), "Umsatz Nord", 1, 'SUMIFS(D:D,B:B,"Nord",C:C,"Maus")')
    cells[(2, 8)] = Cell("1791.03", 'SUMIFS(D:D,B:B,"Nord",C:C,"Maus")')
    found = hints(cells)["label"]
    # Die Ursache ist die Bedingung, die nicht im Namen steht, nicht „8 Zeilen fehlen“.
    assert ("„Umsatz Nord“ in H2 filtert zusätzlich nach Produkt „Maus“ (C:C;\"Maus\"). Das steht nicht im Namen, der "
            "nur Region „Nord“ sagt. Dadurch fehlen 8 der 10 Zeilen: D2 (Tastatur), D3 (Dock), D5 (Dock) und 5 weitere. "
            "Richtig wäre 26688,97 statt 1791,03.") == found.text
    assert found.level == "fehler" and 'Korrigiert: =SUMMEWENNS(D:D;B:B;"Nord"), nachgerechnet 26688,97.' in found.tip
    assert "nenn die Kennzahl „Umsatz Nord Maus“" in found.tip
    # Folgt ein Fehler nur aus dem darüber, steht nur der erste da.
    cells = figure(table(), "Umsatz alle Nord", 23196.67, "SUM(D2:D10)")
    figure(cells, "Durchschnitt Umsatz Nord", 2319.667, "H2/10", row=3)
    assert [f.key for f in analyse(cells, layout(Sheet("U", cells))).findings if f.key.startswith(("label", "count"))] == ["label:H2"]


def test_months_that_are_not_a_run_and_the_row_that_is_missing():
    # Wie H8 der Übungsmappe: nach Datum sortiert, Januar, Februar und Juli, und die letzte Juli-Zeile fehlt.
    rows = sorted(ROWS, key=lambda row: row[0][3:5] + row[0][:2])
    cells = figure(table(rows), "Umsatz Januar, Februar und Juli", 28086.77, "SUM(D2:D9)+SUM(D28:D30)", row=8)
    found = hints(cells)["label"]
    assert ("„Umsatz Januar, Februar und Juli“ in H8 summiert D2:D9 und D28:D30, aber der Bereich D28:D30 hört eine "
            "Zeile zu früh auf: D31 (27.07.2026) gehört noch dazu. Richtig wäre 32877,65 statt 28086,77.") == found.text
    assert "Korrigiert: =SUMME(D2:D9)+SUMME(D28:D31), nachgerechnet 32877,65." in found.tip
    assert found.level == "fehler" and "bis" not in found.text
    # Die Formel ist nachgerechnet und nimmt die Monate einzeln, nicht Januar bis Juli.
    assert ('Sicherer, auch wenn Zeilen dazukommen oder neu sortiert wird: =SUMMEWENNS(D:D;A:A;">="&DATUM(2026;1;1);'
            'A:A;"<"&DATUM(2026;3;1))+SUMMEWENNS(D:D;A:A;">="&DATUM(2026;7;1);A:A;"<"&DATUM(2026;8;1)).') in found.tip
    # Eine Zeile doppelt, eine zu viel, eine Gruppe dazu: Er sagt, welche und warum.
    text = hints(figure(table(rows), "Umsatz Januar und Februar", 1, "SUM(D2:D9)+D5", row=8))["label"].text
    assert "rechnet =SUMME(D2:D9)+D5, aber D5 steckt schon in D2:D9 und wird mit +D5 ein zweites Mal gezählt." in text
    assert "19816,71 statt 23766,73" in text
    text = hints(figure(table(rows), "Umsatz Juli", 1, "SUM(D27:D31)", row=8))["label"].text
    assert "D27 (29.06.2026) gehört nicht dazu" in text
    text = hints(figure(table(rows), "Umsatz Ost Juli", 1, "SUM(D28:D31)", row=8))["label"].text
    assert "D29 (13.07.2026, Nord) und D30 (20.07.2026, Nord) gehören nicht dazu" in text
    # Die feste Zahl: genau eine Zeile erklärt den Unterschied, das steht da, aber nur zum Prüfen.
    found = hints(figure(table(rows), "Umsatz Januar, Februar und Juli", 28086.77, row=8))["label"]
    assert "Es fehlt genau D31 (27.07.2026)." in found.text and found.level == "pruefen"
    # Richtig über Monate, die keine Reihe sind: nichts zu sagen.
    right = 'SUMIFS(D:D,A:A,">="&DATE(2026,1,1),A:A,"<"&DATE(2026,3,1))+SUMIFS(D:D,A:A,">="&DATE(2026,7,1),A:A,"<"&DATE(2026,8,1))'
    assert "label" not in hints(figure(table(rows), "Umsatz Januar, Februar und Juli", 32877.65, right, row=8))


def test_the_month_names_say_what_the_name_says():
    from voiceagent.bereiche import _months_text

    assert _months_text({1, 2, 7}) == "Januar, Februar und Juli"
    assert _months_text({1, 2, 3, 7}) == "Januar bis März und Juli"
    assert _months_text({3}) == "März" and _months_text({11, 12}) == "November und Dezember"


def fill(cells, col, top, bottom, formula):
    from voiceagent.xlsx import shift

    for r in range(top, bottom + 1):
        cells[(r, col)] = Cell("0", shift(formula, r - top, 0))
    return cells


def test_ranges_that_slide_when_copied_down():
    # Anteil je Zeile: Der Nenner SUMME(D2:D31) wird beim Herunterziehen zu D3:D32, D4:D33 …
    cells = table()
    cells[(1, 5)] = Cell("Anteil")
    found = hints(fill(cells, 5, 2, 31, "D2/SUM(D2:D31)"))
    slide = found["slide"]
    assert ("E3:E31 sind aus E2 heruntergezogen, und der Bereich D2:D31 in SUMME rutscht dabei mit, weil vor den "
            "Zeilen die $-Zeichen fehlen: E3 nimmt D3:D32, E31 schon D31:D60. Ganz unten in E31 fehlen so schon 29 "
            "Zeilen vom Anfang der Tabelle. 29 von 29 Ergebnissen sind dadurch falsch") in slide.text
    assert slide.level == "fehler" and "Schreib in E2 =D2/SUMME(D$2:D$31) und zieh es wieder herunter" in slide.tip
    assert "range" not in found
    # SUMMEWENN je Region daneben: Heute zufällig richtig, weil die Tabelle nach Region sortiert ist.
    cells = table()
    for k, region in enumerate(["Nord", "Ost", "Süd", "West"]):
        cells[(2 + k, 7)] = Cell(region)
    slide = hints(fill(cells, 8, 2, 5, "SUMIF(B2:B31,G2,D2:D31)"))["slide"]
    assert "Heute ändert das zufällig nichts" in slide.text and slide.level == "pruefen"
    assert "=SUMMEWENN(B$2:B$31;G2;D$2:D$31)" in slide.tip
    # Rang ohne $: falsch, sobald die Zeilen oben fehlen.
    cells = table()
    cells[(1, 5)] = Cell("Rang")
    assert hints(fill(cells, 5, 2, 31, "RANK(D2,D2:D31)"))["slide"].level == "fehler"


def test_copied_ranges_that_are_meant_to_move_stay_quiet():
    for formula, top in [("SUM(D$2:D2)", 2), ("D2/SUM(D$2:D$31)", 2), ("RANK(D2,$D$2:$D$31)", 2),
                         ("SUM(D2:D$31)", 2), ("AVERAGE(D2:D4)", 4), ("D2/SUM(D:D)", 2)]:
        cells = table()
        cells[(1, 5)] = Cell("Wert")
        if formula == "AVERAGE(D2:D4)":
            for r in range(4, 32):
                cells[(r, 5)] = Cell("0", f"AVERAGE(D{r - 2}:D{r})")
        else:
            fill(cells, 5, top, 31, formula)
        assert "slide" not in hints(cells), formula


def ursache(label, formula, rows=ROWS):
    cells = table(rows)
    cells[(4, 6)], cells[(4, 7)] = Cell(label), Cell("0", formula)
    return hints(cells).get("label")


def test_the_cause_is_the_part_of_the_formula_that_is_wrong():
    # Wie G4 der Übungsmappe: eine Bedingung, die der Name nicht sagt. Nicht „5 Zeilen fehlen“, sondern warum.
    found = ursache("Umsatz Ost:", 'SUMIFS(D:D,B:B,"Ost",C:C,"Monitor")')
    assert ("„Umsatz Ost“ in G4 filtert zusätzlich nach Produkt „Monitor“ (C:C;\"Monitor\"). Das steht nicht im Namen, "
            "der nur Region „Ost“ sagt. Dadurch fehlen 5 der 8 Zeilen: D12 (Laptop), D13 (Tastatur), D16 (Tastatur) und "
            "2 weitere. Richtig wäre 18363,86 statt 6745,98.") == found.text
    assert found.tip.startswith('Korrigiert: =SUMMEWENNS(D:D;B:B;"Ost"), nachgerechnet 18363,86.')
    assert "nenn die Kennzahl „Umsatz Ost Monitor“" in found.tip
    assert "label" not in hints(figure(table(), "Umsatz Ost Monitor", 0, 'SUMIFS(D:D,B:B,"Ost",C:C,"Monitor")'))
    # Je Art von Fehler die Stelle in der Formel und die Korrektur, nachgerechnet.
    for label, formula, cause, fixed in [
        ("Umsatz Nord", 'SUMIF(B:B,"Ost",D:D)', "prüft Region „Ost“, der Name sagt aber „Nord“.", 'SUMMEWENN(B:B;"Nord";D:D)'),
        ("Umsatz Nord Maus", 'SUMIF(B:B,"Nord",D:D)', "filtert nicht nach Produkt „Maus“, obwohl das im Namen steht.",
         'SUMMEWENNS(D:D;B:B;"Nord";C:C;"Maus")'),
        ("Durchschnitt Umsatz Nord", 'SUMIF(B:B,"Nord",D:D)',
         "rechnet SUMMEWENN, der Name verlangt aber den Mittelwert (MITTELWERTWENN).", 'MITTELWERTWENN(B:B;"Nord";D:D)'),
        ("Durchschnitt Umsatz Nord", "SUM(D2:D11)/9", "teilt durch 9. Mit Region „Nord“ gibt es aber 10 Zeilen.",
         "SUMME(D2:D11)/10"),
        ("Umsatz Nord", "SUM(D3:D11)", "der Bereich D3:D11 fängt eine Zeile zu spät an: D2 (12.01.2026) fehlt.", "SUMME(D2:D11)"),
        ("Umsatz Nord", "SUM(D2:D11)+D5", "D5 steckt schon in D2:D11 und wird mit +D5 ein zweites Mal gezählt.", "SUMME(D2:D11)"),
        ("Umsatz Nord", "SUM(D2:D11)+D12", "+D12 nimmt D12 (Ost) mit, und das gehört nicht zu Region „Nord“.", "SUMME(D2:D11)"),
        ("Umsatz Nord", "SUM(D2:D10)-D11", "D11 wird abgezogen statt dazugezählt.", "SUMME(D2:D10)+D11"),
        ("Umsatz Nord", "SUM(D2,D3,D4,D5,D6,D7,D8,D9,D11)", "D10 (13.07.2026) fehlt in der Formel.",
         "SUMME(D2;D3;D4;D5;D6;D7;D8;D9;D11;D10)"),
        ("Summe Nord", "AVERAGE(D2:D11)", "rechnet MITTELWERT, der Name verlangt aber die Summe (SUMME).", "SUMME(D2:D11)"),
    ]:
        found = ursache(label, formula)
        assert cause in found.text and found.level == "fehler", (formula, found.text)
        assert found.tip.startswith(f"Korrigiert: ={fixed}, nachgerechnet"), (formula, found.tip)


def test_date_limits_in_the_formula():
    rows = sorted(ROWS, key=lambda row: row[0][3:5] + row[0][:2])
    found = ursache("Umsatz März 2026", 'SUMIFS(D:D,A:A,">="&DATE(2026,3,1),A:A,"<"&DATE(2026,3,30))', rows)
    assert ('die Grenze "<"&DATUM(2026;3;30) müsste "<"&DATUM(2026;4;1) heißen. Dadurch fehlt D14 (30.03.2026).') in found.text
    assert 'Korrigiert: =SUMMEWENNS(D:D;A:A;">="&DATUM(2026;3;1);A:A;"<"&DATUM(2026;4;1)), nachgerechnet 10900,41.' in found.tip
    found = ursache("Umsatz März 2026", 'SUMIFS(D:D,A:A,">="&DATE(2025,3,1),A:A,"<"&DATE(2026,4,1))', rows)
    assert 'die Grenze ">="&DATUM(2025;3;1) müsste ">="&DATUM(2026;3;1) heißen. Dadurch zählen D2 (05.01.2026)' in found.text
    # Zeilen, die nicht beieinander liegen, passen in keinen einzelnen Bereich: Das steht als Ursache da.
    found = ursache("Umsatz Ost Juli", "SUM(D28:D31)", rows)
    assert "liegen nicht beieinander (D28 und D31), ein einzelner Bereich kann sie nicht fassen." in found.text


def test_two_errors_in_one_formula():
    # Keine einzelne Änderung macht die Formel richtig, aber zwei an verschiedenen Stellen: beide nennen.
    for label, formula, causes, fixed in [
        ("Umsatz Nord", 'SUMIFS(D:D,B:B,"Ost",C:C,"Monitor")',
         ["filtert zusätzlich nach Produkt „Monitor“", "prüft Region „Ost“, der Name sagt aber „Nord“."],
         'SUMMEWENNS(D:D;B:B;"Nord")'),
        ("Umsatz Nord Maus", 'SUMIF(B:B,"Ost",D:D)',
         ["prüft Region „Ost“", "filtert nicht nach Produkt „Maus“"], 'SUMMEWENNS(D:D;B:B;"Nord";C:C;"Maus")'),
        ("Durchschnitt Umsatz Nord", 'SUMIFS(D:D,B:B,"Nord",C:C,"Monitor")',
         ["der Name verlangt aber den Mittelwert (MITTELWERTWENNS)", "filtert zusätzlich nach Produkt „Monitor“"],
         'MITTELWERTWENNS(D:D;B:B;"Nord")'),
        ("Durchschnitt Umsatz Nord", "SUM(D2:D10)/9",
         ["teilt durch 9", "Der Bereich D2:D10 hört eine Zeile zu früh auf: D11 (20.07.2026) gehört noch dazu."],
         "SUMME(D2:D11)/10"),
        ("Umsatz Nord", "SUM(D2:D10)+D5",
         ["D5 steckt schon in D2:D10 und wird mit +D5 ein zweites Mal gezählt.", "hört eine Zeile zu früh auf"],
         "SUMME(D2:D11)"),
        ("Umsatz Nord", "SUM(D2:D9)-D11", ["D2:D9 hört eine Zeile zu früh auf", "D11 wird abgezogen statt dazugezählt."],
         "SUMME(D2:D10)+D11"),
        ("Umsatz Nord", "SUM(D2:D10)-D12", ["-D12 zieht D12 (Ost) ab, und das gehört nicht zu Region „Nord“.",
                                             "hört eine Zeile zu früh auf"], "SUMME(D2:D11)"),
    ]:
        found = ursache(label, formula)
        assert "hat zwei Fehler. Erstens: " in found.text and " Zweitens: " in found.text, (formula, found.text)
        assert all(cause in found.text for cause in causes), (formula, found.text)
        assert found.tip.startswith(f"Korrigiert: ={fixed}, nachgerechnet"), (formula, found.tip)
    # Derselbe Bereich an beiden Enden falsch ist ein Satz, kein „zwei Fehler“.
    found = ursache("Umsatz Nord", "SUM(D3:D10)")
    assert ("„Umsatz Nord“ in G4 nimmt D3:D10, aber der Bereich fängt eine Zeile zu spät an und hört eine Zeile zu früh "
            "auf. Dadurch fehlen D2 (12.01.2026) und D11 (20.07.2026).") in found.text
    assert found.tip.startswith("Korrigiert: =SUMME(D2:D11), nachgerechnet 26688,97.")
    # Richtig gerechnet, aber verschlungen: =SUMME(D2:D12)-D12 ist nicht die Korrektur.
    assert ursache("Umsatz Nord", "SUM(D2:D11)-D11").tip.startswith("Korrigiert: =SUMME(D2:D11),")


def kreuz(bad=None, corner="Umsatz"):
    """Eine Übersicht neben der Tabelle: Regionen untereinander, Produkte nebeneinander, mit SUMMEWENNS."""
    cells = table()
    if corner:
        cells[(3, 6)] = Cell(corner)
    for j, product in enumerate(["Maus", "Tastatur", "Monitor"]):
        cells[(3, 7 + j)] = Cell(product)
    for i, region in enumerate(["Nord", "Ost", "Süd", "West"]):
        cells[(4 + i, 6)] = Cell(region)
        for j in range(3):
            cells[(4 + i, 7 + j)] = Cell("0", f'SUMIFS($D:$D,$B:$B,$F{4 + i},$C:$C,{"GHI"[j]}$3)')
    if bad:
        cells[bad[0]] = Cell(*bad[1:])
    return {f.key: f for f in analyse(cells, layout(Sheet("T", cells)), name="T").findings if f.key.startswith("label")}


def test_cross_tables_take_the_name_from_row_and_column():
    # Richtig gerechnet bleibt still: G5 heißt „Ost“ und „Maus“, nicht nur „Ost“.
    assert kreuz() == {} and kreuz(corner="") == {} and kreuz(corner="Region / Produkt") == {}
    # Die Bedingung aus der falschen Kopfzelle.
    found = kreuz(((5, 8), "0", "SUMIFS($D:$D,$B:$B,$F6,$C:$C,H$3)"))["label:H5"]
    assert found.text == ("„Umsatz / Ost / Tastatur“ in H5 nimmt Region aus $F6 („Süd“), gemeint ist aber „Ost“ aus $F5. "
                          "Richtig wäre 4888,43 statt 6446,24.")
    assert found.tip == "Korrigiert: =SUMMEWENNS($D:$D;$B:$B;$F5;$C:$C;H$3), nachgerechnet 4888,43."
    found = kreuz(((5, 8), "0", "SUMIFS($D:$D,$B:$B,$F5,$C:$C,G$3)"))["label:H5"]
    assert "nimmt Produkt aus G$3 („Maus“), gemeint ist aber „Tastatur“ aus H$3." in found.text
    # Eine fehlende Bedingung kommt aus dem Kopf, geschrieben wie die anderen.
    found = kreuz(((5, 8), "0", "SUMIFS($D:$D,$B:$B,$F5)"))["label:H5"]
    assert found.tip.startswith("Korrigiert: =SUMMEWENNS($D:$D;$B:$B;$F5;$C:$C;H$3), nachgerechnet 4888,43.")
    # Ein getippter Wert statt der Kopfzelle.
    found = kreuz(((5, 8), "0", 'SUMIFS($D:$D,$B:$B,"Nord",$C:$C,H$3)'))["label:H5"]
    assert "prüft Region „Nord“, der Name sagt aber „Ost“." in found.text
    assert found.tip.startswith("Korrigiert: =SUMMEWENNS($D:$D;$B:$B;$F5;$C:$C;H$3),")
    # Eine feste Zahl mitten in der Übersicht.
    found = kreuz(((6, 7), "1234"))["label:G6"]
    assert found.text.startswith("In G6 steht „Umsatz / Süd / Maus“ als feste Zahl 1234.") and found.level == "pruefen"


def test_cross_tables_with_shares_stay_quiet():
    # „Anteil“ in der Ecke oder Anteile als Formel: Die Köpfe sagen keine Summe.
    cells = table()
    cells[(3, 6)] = Cell("Anteil")
    for j, product in enumerate(["Maus", "Tastatur", "Monitor"]):
        cells[(3, 7 + j)] = Cell(product)
    for i, region in enumerate(["Nord", "Ost", "Süd", "West"]):
        cells[(4 + i, 6)] = Cell(region)
        for j in range(3):
            cells[(4 + i, 7 + j)] = Cell("0", f'SUMIFS($D:$D,$B:$B,$F{4 + i},$C:$C,{"GHI"[j]}$3)/SUMIFS($D:$D,$B:$B,$F{4 + i})')
    assert "label" not in hints(cells)
    cells[(3, 6)] = Cell("")
    assert "label" not in hints(cells)


def test_a_summary_beside_the_table_does_not_move_its_header():
    # Die Köpfe einer Übersicht in Zeile 2 machen die erste Datenzeile nicht zur Kopfzeile.
    cells = table()
    for j, product in enumerate(["Maus", "Tastatur", "Monitor"]):
        cells[(2, 7 + j)] = Cell(product)
    cells[(3, 6)] = Cell("Nord")
    for j in range(3):
        cells[(3, 7 + j)] = Cell("0", f'SUMIFS($D:$D,$B:$B,$F3,$C:$C,{"GHI"[j]}$2)')
    assert layout(Sheet("T", cells)).header_row == 1
