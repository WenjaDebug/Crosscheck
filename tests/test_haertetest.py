"""Befunde aus dem Härtetest mit fremden Excel-Dateien: Abstürze, falsch nachgerechnete Werte und Fehlalarme."""

import random

import pytest
from test_workbook import xlsx

from voiceagent import formel, pruefer
from voiceagent.analyse import analyse
from voiceagent.formel import Err, Evaluator, parse
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Sheet, kept_text, number_text, read_xlsx


def calc(formula, values=None):
    values = values or {}
    return Evaluator(lambda p: values.get(p)).run(parse(formula))


def keys(cells, name="T"):
    report = analyse(cells, layout(Sheet(name, cells)), exact=True, name=name)
    return {f.key: f for f in report.findings}


def test_powers_and_dates_compute_like_excel():
    assert calc("(-8)^0.5") == Err("#NUM!") and calc("0^-1") == Err("#DIV/0!") and calc("0^0") == Err("#NUM!")
    assert calc("(-2)^3") == -8 and calc("POWER(4,0.5)") == 2
    with pytest.raises(formel.Unknown):
        calc("(-8)^(1/3)")  # Was Excel hier rechnet, wissen wir nicht sicher: lieber gar nicht nachrechnen.
    # Jahre unter 1900 zählen ab 1900, Tage und Monate laufen über, Excel kennt den 29.2.1900.
    assert calc("DATE(2009,1,400)") == 40213 and calc("DATE(1,1,1)") == 367 and calc("DATE(2024,14,1)") == 45689
    assert calc("DATE(1900,2,29)") == 60 and calc("DATE(1900,3,1)") == 61 and calc("DATE(1900,1,0)") == 0
    assert calc("DATE(1900,1,-1)") == Err("#NUM!") and calc("DATE(10000,1,1)") == Err("#NUM!")
    assert calc("NA()") == Err("#N/A") and calc("ASIN(2)") == Err("#NUM!") and calc("FACT(5)") == 120
    assert calc("ATAN2(1,1)") == pytest.approx(0.785398, rel=1e-5) and calc("DEGREES(PI())") == pytest.approx(180)
    assert calc("CEILING(2.5,1)") == 3 and calc("CEILING(-2.5,2)") == -2 and calc("CEILING(2,-2)") == Err("#NUM!")


def test_sheets_without_header_and_negative_powers_do_not_crash():
    # Wie die Testmappen von Formelrechnern: Formeln ohne Kopfzeile, eine Potenz einer negativen Zahl.
    cells = {(1, 1): Cell("3", "SUM(1,2)"), (2, 1): Cell("-10"), (2, 2): Cell("-1000", "A2^3"), (3, 1): Cell("7", "MAX(1,7)")}
    assert analyse(cells, layout(Sheet("T", cells)), exact=True, name="T") is not None


def test_a_time_is_compared_as_precisely_as_excel_shows_it():
    cells = {(1, 1): Cell("Beginn"), (1, 2): Cell("Dauer"), (2, 1): Cell("08:00"), (2, 2): Cell("02:10", "1/11")}
    assert not [k for k in keys(cells) if k.startswith("recalc:")]
    cells[(2, 2)] = Cell("03:10", "1/11")
    assert "recalc:B2" in keys(cells)


def test_text_that_looks_like_an_error_value_is_text():
    data = xlsx(("Fehler", {"A1": "Beschreibung", "B1": "Wert", "A2": "#div/0!", "B2": ("=", "1/0", "#DIV/0!")}))
    sheet = read_xlsx(data)[0]
    assert sheet.cells[(2, 1)].text == "#div/0!"
    done = pruefer.Checker(None).check_file("x.xlsx", data)
    cells = {h["cell"] for s in done["sheets"] for h in s["hints"] if h["key"].startswith("error:")}
    assert cells == {"B2"}


def test_password_old_and_empty_files_say_what_to_do():
    ole = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 64
    checker = pruefer.Checker(None)
    locked = checker.check_file("x.xlsx", ole + "EncryptionInfo".encode("utf-16-le"))["text"]
    assert "Kennwort" in locked and "Nimm eine" not in locked
    assert ".xls" in checker.check_file("x.xls", ole)["text"]
    assert checker.check_file("x.xlsx", b"")["text"].startswith("Die Datei ist leer.")


def test_doubled_spaces_are_kept_for_recalculation():
    assert kept_text("foo  bar") == "foo  bar" and kept_text("foo bar") is None and kept_text(" A") == " A"


def _column(header, values):
    cells = {(1, 1): Cell("Fall"), (1, 2): Cell(header)}
    for i, v in enumerate(values):
        cells[(i + 2, 1)] = Cell(f"F{i}")
        cells[(i + 2, 2)] = Cell(number_text(v))
    return cells


def test_an_outlier_needs_a_gap_and_a_measured_value():
    rng = random.Random(3)
    # Steuerbeträge: viele kleine, ein langes oberes Ende ohne Lücke. Der größte ist nur der größte.
    tail = [0.0] * 30 + [round(rng.lognormvariate(2.5, 1.6), 2) for _ in range(400)]
    assert not [k for k in keys(_column("Betrag", tail)) if k.startswith(("outlier:", "scale:"))]
    # Ein Tippfehler steht allein: 720 statt 72. Mal 10 daneben, darum mit der Zahl zum Einsetzen.
    weights = [round(rng.uniform(55, 95), 1) for _ in range(40)] + [720.0]
    found = keys(_column("Gewicht", weights))
    assert found["scale:B42"].fix["cells"] == {"B42": 72.0}
    # Postleitzahlen messen nichts.
    plz = keys(_column("PLZ", [8001.0] * 30 + [8002.0] * 20 + [3011.0]))
    assert not [k for k in plz if k.startswith(("outlier:", "scale:"))]


def test_a_dollar_that_changes_nothing_is_no_pattern_break():
    cells = {(1, 1): Cell("Jahr"), (1, 2): Cell("Wert"), (1, 3): Cell("Doppelt")}
    for r in range(2, 10):
        cells[(r, 1)] = Cell(str(2000 + r))
        cells[(r, 2)] = Cell(str(r * 3))
        cells[(r, 3)] = Cell(str(r * 6), f"$B{r}*2" if r == 5 else f"B{r}*2")
    assert not [k for k in keys(cells) if k.startswith("pattern:")]
    cells[(5, 3)] = Cell(str(5 * 6), "B4*2")
    assert "pattern:C5" in keys(cells)


def test_two_units_only_when_both_cover_the_same_values():
    months = [1.0, 1.0, 1.0, 12.0, 13.0, 14.0, 12.0, 13.0, 12.0, 14.0]
    assert not [k for k in keys(_column("Month", months)) if k.startswith("units:")]
    assert not [k for k in keys(_column("Wert", months)) if k.startswith("units:")]
    heights = [172.0, 165.0, 181.0, 158.0, 190.0, 177.0, 169.0, 1.74, 1.68, 175.0, 183.0]
    found = [f for k, f in keys(_column("Größe", heights)).items() if k.startswith("units:")]
    assert found and "genau" not in found[0].text


def test_no_rounding_question_for_long_data_columns():
    found = keys(_column("Kosten", [0.5] * 300))
    assert not [k for k in found if k.startswith("round:")]


def test_a_repeated_header_starts_a_second_table():
    cells = {(1, 1): Cell("Item"), (1, 2): Cell("Units")}
    for r in range(2, 6):
        cells[(r, 1)], cells[(r, 2)] = Cell(f"A{r}"), Cell(str(r))
    cells[(7, 1)], cells[(7, 2)] = Cell("Item"), Cell("Units")
    for r in range(8, 11):
        cells[(r, 1)], cells[(r, 2)] = Cell(f"B{r}"), Cell(str(r))
    assert not [k for k in keys(cells) if k.startswith("text:")]


def test_an_empty_cell_inside_text_is_no_zero():
    cells = {(1, 1): Cell("Name"), (1, 2): Cell("Zusatz"), (1, 3): Cell("Anzeige"), (1, 4): Cell("Fehlt"),
             (2, 1): Cell("Huber"), (2, 3): Cell("Huber", 'A2&" "&B2'), (2, 4): Cell("WAHR", "ISNA(B2)"),
             (3, 1): Cell("Meier"), (3, 3): Cell("Meier", "CONCAT(A3,B3)"), (3, 4): Cell("0", "B3*2")}
    found = keys(cells)
    assert "emptyref:C2" not in found and "emptyref:C3" not in found and "emptyref:D2" not in found
    assert "emptyref:D3" in found


def test_false_and_true_written_as_functions():
    # So speichert LibreOffice WAHR und FALSCH: als Funktion. Vorher war jede solche Formel „nicht nachrechenbar“.
    assert calc("FALSE()") is False and calc("TRUE()") is True and calc("IF(FALSE(),1,2)") == 2


def _cases(n, extra=()):
    """Ein Blatt „Fälle“ mit n Fällen (Kosten 100, 110, …) und darunter einer Summenzeile."""
    cells = {"A1": "Fall-Nr", "B1": "Geschlecht", "C1": "Kosten"}
    for i in range(n):
        r = i + 2
        cells[f"A{r}"], cells[f"B{r}"], cells[f"C{r}"] = f"F{100 + i}", "w" if i % 3 else "m", 100 + 10 * i
    total = sum(100 + 10 * i for i in range(n))
    cells[f"A{n + 2}"], cells[f"C{n + 2}"] = "Summe", ("=", f"SUM(C2:C{n + 1})", total)
    cells.update(extra)
    return cells


def _hints(*sheets):
    done = pruefer.Checker(None).check_file("x.xlsx", xlsx(*sheets))
    return {(s["name"], h["key"]): h for s in done["sheets"] for h in s["hints"]}


def test_ranges_on_another_sheet_that_miss_rows_or_take_the_total():
    n, costs = 20, [100 + 10 * i for i in range(20)]
    women = sum(1 for i in range(n) if i % 3)
    report = {"A1": "Kennzahl", "B1": "Wert",
              "A2": "Mittelwert Kosten", "B2": ("=", "AVERAGE(Fälle!C2:C19)", sum(costs[:18]) / 18),
              "A3": "Mittel mit Summe", "B3": ("=", "AVERAGE(Fälle!C2:C22)", (sum(costs) * 2) / 21),
              "A4": "Mittel richtig", "B4": ("=", "AVERAGE(Fälle!C2:C21)", sum(costs) / 20),
              "A5": "Anteil Frauen", "B5": ("=", 'COUNTIF(Fälle!B2:B21,"w")/22', women / 22)}
    found = _hints(("Fälle", _cases(n)), ("Auswertung", report))
    short = found[("Auswertung", "range:B2")]
    assert "Fälle!C20" in short["text"] and "Fälle!C21" in short["text"] and short["level"] == "fehler"
    assert "Fälle!C22" in found[("Auswertung", "total:B3")]["text"]
    assert not [k for s, k in found if s == "Auswertung" and k.endswith(":B4")]
    share = found[("Auswertung", "count:B5")]
    assert "22" in share["text"] and "20 Einträge" in share["text"] and share["level"] == "pruefen"


def test_a_fixed_count_that_matches_is_no_hint():
    women = sum(1 for i in range(20) if i % 3)
    report = {"A1": "Kennzahl", "B1": "Wert", "A2": "Anteil Frauen", "B2": ("=", 'COUNTIF(Fälle!B2:B21,"w")/20', women / 20),
              "A3": "Anteil pro 100", "B3": ("=", 'COUNTIF(Fälle!B2:B21,"w")/100', women / 100)}
    found = _hints(("Fälle", _cases(20)), ("Auswertung", report))
    assert not [k for s, k in found if k.startswith("count:")]


def test_a_dropped_dollar_in_a_lookup_into_another_sheet_is_a_pattern_break():
    codes = ["K1", "C1", "M1", "N1"]
    master = {"A1": "Station", "B1": "Klinik"}
    for i, c in enumerate(codes):
        master[f"A{i + 2}"], master[f"B{i + 2}"] = c, f"Klinik {c}"
    cases = {"A1": "Fall-Nr", "B1": "Station", "C1": "Klinik"}
    for r in range(2, 14):
        code = codes[r % 4]
        table = "Stammdaten!A7:B10" if r == 7 else "Stammdaten!$A$2:$B$5"
        cases[f"A{r}"], cases[f"B{r}"] = f"F{r}", code
        cases[f"C{r}"] = ("=", f"VLOOKUP(B{r},{table},2,FALSE)", "#N/A" if r == 7 else f"Klinik {code}")
    found = _hints(("Stammdaten", master), ("Fälle", cases))
    assert "$A$2:$B$5" in found[("Fälle", "pattern:C7")]["text"]


def test_a_rough_lookup_is_only_the_cause_when_an_exact_one_finds_the_value():
    from voiceagent.fehlerwerte import _exactly

    cells = {(1, 1): Cell("Code"), (1, 2): Cell("Betten")}
    for r, (code, beds) in enumerate([("K2", 20), ("C1", 22), ("M1", 18), ("B9", 24)], start=2):
        cells[(r, 1)], cells[(r, 2)] = Cell(code), Cell(str(beds))
    sheet = formel_sheet(cells)
    calc, _ = sheet.evaluator()
    assert _exactly(calc, parse('VLOOKUP("M1",A2:B5,2)')) == 18
    assert _exactly(calc, parse('VLOOKUP("Z9",A2:B5,2)')) is None
    assert _exactly(calc, parse('VLOOKUP("M1",A2:B5,2,FALSE)')) is None
    assert _exactly(calc, parse('MATCH("M1",A2:A5)')) == 3


def formel_sheet(cells):
    from voiceagent.analyse import Sheet as Checked

    return Checked(cells, layout(Sheet("T", cells)), "T")


def test_missing_rows_are_named_by_their_case_number_not_a_birth_date():
    cells = {"A1": "Fall-Nr", "B1": "Geburtsdatum", "C1": "Kosten"}
    for i in range(12):
        r = i + 2
        cells[f"A{r}"], cells[f"B{r}"], cells[f"C{r}"] = f"F{100 + i}", (20000 + 37 * i, 1), 100 + 10 * i
    cells["A14"], cells["C14"] = "Summe", ("=", "SUM(C2:C11)", sum(100 + 10 * i for i in range(10)))
    styles = '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs>'
    done = pruefer.Checker(None).check_file("x.xlsx", xlsx(("Fälle", cells), styles=styles))
    hints = {h["key"]: h for s in done["sheets"] for h in s["hints"]}
    # Mit dem Spaltenkopf: F110 allein läse man als Zelle.
    assert "C12 (Fall-Nr F110) und C13 (Fall-Nr F111)" in hints["label:C14"]["text"]
    # Die genauere Prüfung bringt die Korrektur mit: kein zweiter Hinweis „lässt Zeilen aus“ zur selben Zelle.
    assert "gap:C14" not in hints and hints["label:C14"]["fix"]


def test_the_long_tail_of_costs_is_no_outlier_but_a_typo_is():
    rng = random.Random(7)
    costs = [round(rng.lognormvariate(9.2, 0.6), 2) for _ in range(36)]
    costs[5] = max(costs) * 1.6  # Ein großer, aber echter Fall am langen Ende.
    assert not [k for k in keys(_column("Kosten CHF", costs)) if k.startswith(("outlier:", "scale:"))]
    weights = [round(rng.gauss(72, 9), 1) for _ in range(36)]
    right = weights[4]
    weights[4] = right * 10
    assert keys(_column("Gewicht kg", weights))["scale:B6"].fix["cells"] == {"B6": right}
    weights[4] = right / 10
    assert keys(_column("Gewicht kg", weights))["scale:B6"].fix["cells"] == {"B6": right}


def test_a_text_key_makes_the_column_pattern_count_only_where_it_finds_something():
    cases = {"A1": "Code", "B1": "Klinik", "C1": "Kosten"}
    for r, (code, clinic, cost) in enumerate([("K1", "Kardio", 10), ("C1", "Chir", 20), ("K2", "Kardio", 30),
                                              ("M1", "Med", 40), ("C2", "Chir", 50), ("N1", "Neuro", 60)], start=2):
        cases[f"A{r}"], cases[f"B{r}"], cases[f"C{r}"] = code, clinic, cost
    sums = {"Kardio": 40, "Chir": 70, "Med": 40, "Neuro": 60}

    def report(wrong_row=None, average=False):
        cells = {"A1": "Klinik", "B1": "Kosten"}
        for r, clinic in enumerate(sums, start=2):
            # Die Bedingung sucht in der Spalte Code statt Klinik: Dort steht „Chir“ nicht.
            column = "Fälle!$A$2:$A$7" if r == wrong_row else "Fälle!$B$2:$B$7"
            cells[f"A{r}"], cells[f"B{r}"] = clinic, ("=", f"SUMIF({column},A{r},Fälle!$C$2:$C$7)",
                                                      0 if r == wrong_row else sums[clinic])
        if average:
            # Eine Zeile darunter, deren Name nirgends vorkommt, rechnet zu Recht anders.
            cells["A6"], cells["B6"] = "Durchschnitt", ("=", "AVERAGE(B2:B5)", 52.5)
        return [k for s, k in _hints(("Fälle", cases), ("Auswertung", cells)) if s == "Auswertung"]

    assert "pattern:B3" in report(wrong_row=3)
    assert not [k for k in report(average=True) if k.startswith("pattern:")]


def test_parts_named_with_backslashes_are_read():
    # So schreiben manche Programme die Teile (xl\worksheets\sheet1.xml); Excel öffnet die Datei trotzdem.
    import io
    import zipfile

    plain = zipfile.ZipFile(io.BytesIO(xlsx(("Fälle", {"A1": "Fall", "B1": "Kosten", "A2": "F1", "B2": 120}))))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for info in plain.infolist():
            archive.writestr(info.filename.replace("/", "\\").lower(), plain.read(info))
    assert read_xlsx(buffer.getvalue())[0].cells[(2, 2)].value == "120"
