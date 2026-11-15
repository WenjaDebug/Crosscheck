"""Die dritte selbst gebaute Testreihe mit Themen, die es noch nicht gab: Noten, Projektplan, Budget gegen Ist, Umfrage,
Fahrtenbuch, Nachtschichten und Kennzahlen auf einem zweiten Blatt. Was sie zuerst übersah oder falsch meldete: Jede
Mappe ohne Fehler bleibt still, jeder eingebaute Fehler kommt mit der Korrektur zum Einsetzen."""

import random
from datetime import date, timedelta

import pytest
from test_workbook import xlsx

from voiceagent import formel, pruefer
from voiceagent.xlsx import column_letters

STYLES = ('<cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="20" applyNumberFormat="1"/>'
          '<xf numFmtId="14" applyNumberFormat="1"/><xf numFmtId="10" applyNumberFormat="1"/></cellXfs>')
TIME, DAY, PERCENT = 1, 2, 3


def _serial(day: date) -> float:
    return float((day - date(1899, 12, 30)).days)


def _book(*sheets, cached: dict | None = None) -> bytes:
    """Die Mappe mit den Ergebnissen, die Excel speichern würde: Jede Formel ist hier nachgerechnet.
    Je Blatt (Name, {Adresse: Zahl, Text, Datum, "=Formel" oder (Wert, Stil)}). cached: {(Blatt, Adresse): Wert} für
    Formeln, die der Formelrechner nicht nachrechnet (ungefähre Suche in unsortierten Werten)."""
    data = {name.casefold(): cells for name, cells in sheets}
    cached = {(k[0].casefold(), k[1]): v for k, v in (cached or {}).items()}
    memo: dict = {}

    def raw(sheet: str, pos):
        entry = data[sheet].get(f"{column_letters(pos[1])}{pos[0]}")
        return entry[0] if isinstance(entry, tuple) else entry

    def value(sheet: str, pos):
        if (sheet, pos) not in memo:
            entry = raw(sheet, pos)
            if isinstance(entry, str) and entry.startswith("="):
                calc = formel.Evaluator(lambda p: value(sheet, p), other=lambda s, p: value(s.casefold(), p),
                                        extent=lambda s: (200, 30))
                try:
                    memo[(sheet, pos)] = calc.run(formel.parse(entry[1:]))
                except formel.Unknown:
                    memo[(sheet, pos)] = cached[(sheet, f"{column_letters(pos[1])}{pos[0]}")]
            elif isinstance(entry, date):
                memo[(sheet, pos)] = _serial(entry)
            elif isinstance(entry, (int, float)):
                memo[(sheet, pos)] = float(entry)
            else:
                memo[(sheet, pos)] = entry
        return memo[(sheet, pos)]

    out = []
    for name, cells in sheets:
        written = {}
        for address, entry in cells.items():
            style = entry[1] if isinstance(entry, tuple) else None
            plain = entry[0] if isinstance(entry, tuple) else entry
            if isinstance(plain, str) and plain.startswith("="):
                result = value(name.casefold(), formel.position(address))
                written[address] = ("=", plain[1:], repr(result)) + ((style,) if style else ())
            elif isinstance(plain, date):
                written[address] = (_serial(plain), style or DAY)
            elif style:
                written[address] = (plain, style)
            else:
                written[address] = plain
        out.append((name, written))
    return xlsx(*out, styles=STYLES)


def _hints(*sheets) -> dict:
    done = pruefer.Checker(None).check_file("mappe.xlsx", _book(*sheets))
    assert done["state"] == "ok", done
    return {(s["name"], h["key"]): h for s in done["sheets"] for h in s["hints"]}


def _put(cells: dict, row: int, values: list, style: dict | None = None) -> None:
    for i, v in enumerate(values):
        col = chr(ord("A") + i)
        cells[f"{col}{row}"] = (v, style[col]) if style and col in style and row > 1 else v


def _rnd(name: str) -> random.Random:
    return random.Random(sum(map(ord, name)))


# --- Die Mappen ---


def _noten(planted: bool = False) -> tuple:
    rnd = _rnd("noten")
    names = ["Aline", "Ben", "Chiara", "David", "Elif", "Finn", "Gioia", "Hugo", "Ilaria", "Jan", "Kim", "Leon",
             "Mara", "Nils", "Olivia", "Pius", "Rahel", "Simon", "Tara", "Viktor", "Yara", "Zeno"]
    cells = {}
    _put(cells, 1, ["Name", "Klasse", "Prüfung 1", "Prüfung 2", "Prüfung 3", "Schnitt", "Gerundet"])
    for i, name in enumerate(names):
        r = i + 2
        grades = [rnd.choice([3.5, 4, 4.5, 5, 5, 5.5, 6]) for _ in range(3)]
        _put(cells, r, [name, "3a" if i < 11 else "3b", *grades, f"=AVERAGE(C{r}:E{r})", f"=ROUND(F{r}*2,0)/2"])
    cells.update({"I2": "Klassenschnitt", "J2": "=AVERAGE(F2:F23)", "I3": "Schnitt 3a",
                  "J3": '=AVERAGEIF(B2:B23,"3a",F2:F23)', "I4": "Ungenügend", "J4": '=COUNTIF(G2:G23,"<4")'})
    if planted:
        cells["F9"] = "=AVERAGE(C9:D9)"
    return ("Noten", cells),


def _projekt(planted: bool = False) -> tuple:
    rnd = _rnd("projekt")
    phases = {"Konzept": "R. Amsler", "Umsetzung": "T. Brändli", "Test": "C. Rossi", "Einführung": "M. Favre"}
    cells = {}
    _put(cells, 1, ["Nr", "Aufgabe", "Phase", "Verantwortlich", "Start", "Dauer (Tage)", "Ende"])
    start = date(2026, 3, 2)
    for i in range(20):
        r = i + 2
        phase = list(phases)[i * 4 // 20]
        _put(cells, r, [i + 1, f"Aufgabe {i + 1}", phase, phases[phase], start, rnd.choice([2, 3, 5, 5, 8, 10]),
                        f"=WORKDAY(E{r},F{r})"], {"G": DAY})
        start += timedelta(days=rnd.choice([3, 5, 7]))
    cells.update({"I2": "Projektstart", "J2": "=MIN(E2:E21)", "I3": "Projektende", "J3": "=MAX(G2:G21)"})
    if planted:
        cells["G8"] = date(2026, 4, 1)
        cells["G13"] = ("=E13+F13", DAY)
    return ("Projekt", cells),


def _budget(planted: bool = False) -> tuple:
    rnd = _rnd("budget")
    accounts = ["Personal", "Miete", "Energie", "Material", "Reisen", "Weiterbildung", "Informatik", "Versicherungen",
                "Marketing", "Beratung", "Unterhalt", "Telefon"]
    cells = {}
    _put(cells, 1, ["Konto", "Bezeichnung", "Budget", "Ist", "Abweichung", "Abweichung %"])
    for i, name in enumerate(accounts):
        r = i + 2
        b = round(rnd.uniform(8000, 250000), -2)
        _put(cells, r, [4000 + 10 * i, name, b, round(b * rnd.uniform(0.85, 1.15), -1), f"=D{r}-C{r}", f"=E{r}/C{r}"],
             {"F": PERCENT})
    _put(cells, 14, ["", "Total", "=SUM(C2:C13)", "=SUM(D2:D13)", "=D14-C14", "=E14/C14"], {"F": PERCENT})
    if planted:
        cells["F6"] = ("=E6/D6", PERCENT)
        cells["E10"] = 4200
        cells["D14"] = "=SUM(D2:D12)"
    return ("Budget", cells),


def _umfrage(planted: bool = False) -> tuple:
    rnd = _rnd("umfrage")
    cells = {}
    _put(cells, 1, ["Antwort-Nr", "Alter", "Frage 1", "Frage 2", "Frage 3", "Frage 4"])
    for i in range(40):
        _put(cells, i + 2, [i + 1, rnd.randint(19, 78), *[rnd.choice([2, 3, 4, 4, 5]) for _ in range(4)]])
    _put(cells, 43, ["Mittelwert", "=AVERAGE(B2:B41)", '=AVERAGEIF(C2:C41,"<>9")', '=AVERAGEIF(D2:D41,"<>9")',
                     '=AVERAGEIF(E2:E41,"<>9")', '=AVERAGEIF(F2:F41,"<>9")'])
    if planted:
        cells["F12"] = cells["F25"] = 9
        cells["F43"] = "=AVERAGE(F2:F41)"
    return ("Umfrage", cells),


def _fahrtenbuch(planted: bool = False) -> tuple:
    rnd = _rnd("fahrtenbuch")
    cells = {}
    _put(cells, 1, ["Datum", "Zweck", "km Anfang", "km Ende", "Strecke km", "Spesen CHF"])
    km = 48210
    for i in range(26):
        r = i + 2
        trip = rnd.randint(12, 180)
        _put(cells, r, [date(2026, 5, 4) + timedelta(days=i), rnd.choice(["Kunde", "Lieferung", "Schulung"]),
                        km if i == 0 else f"=D{r - 1}", km + trip, f"=D{r}-C{r}", f"=E{r}*$I$2"])
        km += trip
    cells.update({"H2": "Satz pro km", "I2": 0.7, "H3": "km gesamt", "I3": "=SUM(E2:E27)", "H4": "Spesen gesamt",
                  "I4": "=SUM(F2:F27)"})
    if planted:
        cells["C18"] = 49100
        cells["I3"] = "=SUM(E3:E27)"
    return ("Fahrten", cells),


def _nachtschicht(planted: bool = False) -> tuple:
    cells = {}
    _put(cells, 1, ["Datum", "Name", "Schicht", "Beginn", "Ende", "Pause (h)", "Stunden"])
    shifts = {"Früh": (6 / 24, 14.5 / 24), "Spät": (14 / 24, 22.5 / 24), "Nacht": (22 / 24, 6.5 / 24)}
    for i in range(30):
        r = i + 2
        s = list(shifts)[i % 3]
        _put(cells, r, [date(2026, 6, 1) + timedelta(days=i // 3), ["Ana", "Bo", "Cem"][i % 3], s, round(shifts[s][0], 6),
                        round(shifts[s][1], 6), 0.5, f"=MOD(E{r}-D{r},1)*24-F{r}"], {"D": TIME, "E": TIME})
    cells.update({"I2": "Stunden gesamt", "J2": "=SUM(G2:G31)", "I3": "Stunden Nacht",
                  "J3": '=SUMIF(C2:C31,"Nacht",G2:G31)'})
    if planted:
        cells["G10"] = "=(E10-D10)*24-F10"
        cells["G17"] = 8
        cells["J3"] = '=SUMIF(C2:C31,"Nacht",F2:F31)'
    return ("Schichten", cells),


def _kennzahlen(planted: bool = False) -> tuple:
    rnd = _rnd("kennzahlen")
    daten = {"A1": "Monat", "B1": "Umsatz", "C1": "Kosten"}
    for i, m in enumerate(["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]):
        u = round(rnd.uniform(80000, 140000), -2)
        _put(daten, i + 2, [m, u, round(u * rnd.uniform(0.6, 0.8), -2)])
    kpi = {}
    _put(kpi, 1, ["Kennzahl", "Q1", "Q2", "Q3", "Q4", "Jahr"])
    _put(kpi, 2, ["Umsatz", "=SUM(Daten!B2:B4)", "=SUM(Daten!B5:B7)", "=SUM(Daten!B8:B10)", "=SUM(Daten!B11:B13)",
                  "=SUM(B2:E2)"])
    _put(kpi, 3, ["Kosten", "=SUM(Daten!C2:C4)", "=SUM(Daten!C5:C7)", "=SUM(Daten!C8:C10)", "=SUM(Daten!C11:C13)",
                  "=SUM(B3:E3)"])
    _put(kpi, 4, ["Marge", "=(B2-B3)/B2", "=(C2-C3)/C2", "=(D2-D3)/D2", "=(E2-E3)/E2", "=(F2-F3)/F2"],
         dict.fromkeys("BCDEF", PERCENT))
    if planted:
        kpi["C2"] = "=SUM(Daten!B5:B6)"
        kpi["F3"] = "=SUM(B3:D3)"
    return ("Daten", daten), ("KPI", kpi)


BOOKS = [_noten, _projekt, _budget, _umfrage, _fahrtenbuch, _nachtschicht, _kennzahlen]


@pytest.mark.parametrize("book", BOOKS, ids=lambda b: b.__name__.strip("_"))
def test_a_clean_book_is_quiet(book):
    """Dazu gehören: Nachtschichten enden am Morgen danach (kein Ende vor Beginn), RUNDEN(F2*2;0)/2 rundet auf halbe
    Noten (die 2 ist keine Annahme), der erste km-Stand ist der Anfang der Kette, ein Verhältnis der Totale ist keine
    Summe der Prozente, kleine Konten sind keine Ausreißer, und MITTELWERT über das Alter ist kein Musterbruch neben
    MITTELWERTWENN über die Fragen."""
    assert _hints(*book()) == {}


def _fixed(hint: dict) -> dict:
    return (hint.get("fix") or {}).get("cells") or {}


def test_a_sumif_over_the_wrong_column_names_the_column_the_label_says():
    found = _hints(*_nachtschicht(planted=True))
    hint = found[("Schichten", "label:J3")]
    assert hint["level"] == "fehler"
    assert "rechnet mit der Spalte Pause (h) (F2:F31), der Name sagt aber Stunden (G2:G31)" in hint["text"]
    assert _fixed(hint) == {"J3": '=SUMIF(C2:C31,"Nacht",G2:G31)'}


def test_a_night_shift_without_mod_is_one_hint_with_its_negative_hours():
    found = _hints(*_nachtschicht(planted=True))
    hint = found[("Schichten", "pattern:G10")]
    assert hint["text"].endswith("Darum zeigt G10 -16, und eine negative Zahl passt hier nicht.")
    assert _fixed(hint) == {"G10": "=MOD(E10-D10,1)*24-F10"}
    assert ("Schichten", "impossible:G10") not in found


def test_typed_hours_equal_to_the_formula_are_a_question_with_the_formula():
    hint = _hints(*_nachtschicht(planted=True))[("Schichten", "typedcell:G17")]
    assert hint["level"] == "pruefen" and "mit demselben Ergebnis" in hint["text"]
    assert _fixed(hint) == {"G17": "=MOD(E17-D17,1)*24-F17"}


def test_a_quarter_that_leaves_out_a_month_and_a_year_without_q4_are_found():
    found = _hints(*_kennzahlen(planted=True))
    block = found[("KPI", "blocks:C2")]
    assert block["level"] == "fehler"
    assert block["text"].startswith("C2 (Q2 Umsatz) nimmt Daten!B5:B6. Die Formeln daneben nehmen je 3 Zeilen ohne Lücke")
    assert "Daten!B7 (Jun, " in block["text"] and "zählt dadurch in keinem Block" in block["text"]
    assert _fixed(block) == {"C2": "=SUM(Daten!B5:B7)"}
    year = found[("KPI", "parts:F3")]
    assert "zählt aber E3 („Q4“, " in year["text"] and _fixed(year) == {"F3": "=SUM(B3:E3)"}


def test_a_typed_end_date_and_calendar_days_get_workday():
    found = _hints(*_projekt(planted=True))
    typed = found[("Projekt", "typedcell:G8")]
    assert typed["level"] == "fehler" and _fixed(typed) == {"G8": "=WORKDAY(E8,F8)"}
    assert "ist ein festes Datum, 01.04.2026." in typed["text"] and typed["tip"].endswith("nachgerechnet 09.04.2026.")
    calendar = found[("Projekt", "pattern:G13")]
    assert _fixed(calendar) == {"G13": "=WORKDAY(E13,F13)"} and "nachgerechnet 06.05.2026" in calendar["tip"]


def test_planted_mistakes_in_budget_logbook_grades_and_survey_are_found():
    budget = _hints(*_budget(planted=True))
    assert {"pattern:F6", "typedcell:E10", "parts:D14"} <= {k for _, k in budget}
    assert not any(k.startswith(("outlier:", "parts:F14")) for _, k in budget)
    trips = _hints(*_fahrtenbuch(planted=True))
    assert _fixed(trips[("Fahrten", "typedcell:C18")]) == {"C18": "=D17"}
    assert _fixed(trips[("Fahrten", "label:I3")]) == {"I3": "=SUM(E2:E27)"}
    assert not any(k.startswith("typed") and k != "typedcell:C18" for _, k in trips)
    grades = _hints(*_noten(planted=True))
    assert _fixed(grades[("Noten", "pattern:F9")]) == {"F9": "=AVERAGE(C9:E9)"}
    assert not any(k.startswith("constant:") for _, k in grades)
    survey = _hints(*_umfrage(planted=True))
    assert _fixed(survey[("Umfrage", "pattern:F43")]) == {"F43": '=AVERAGEIF(F2:F41,"<>9")'}
    assert ("Umfrage", "pattern:B43") not in survey


def test_blocks_that_tile_with_different_sizes_stay_quiet():
    """Monate mit 31, 28, 31 und 30 Tagen: verschieden groß, aber ohne Lücke hintereinander. Und ein Block, der eine
    Leerzeile auslässt, lässt nichts aus."""
    daten = {"A1": "Tag", "B1": "Menge"}
    for r in range(2, 122):
        daten[f"A{r}"], daten[f"B{r}"] = r - 1, float(r % 7 + 3)
    kpi = {"A1": "Kennzahl", "B1": "Jan", "C1": "Feb", "D1": "Mär", "E1": "Apr",
           "A2": "Menge", "B2": "=SUM(Daten!B2:B32)", "C2": "=SUM(Daten!B33:B60)", "D2": "=SUM(Daten!B61:B91)",
           "E2": "=SUM(Daten!B92:B121)"}
    assert _hints(("Daten", daten), ("KPI", kpi)) == {}
    del daten["A7"], daten["B7"]
    kpi.update({"B2": "=SUM(Daten!B2:B6)", "C2": "=SUM(Daten!B8:B13)", "D2": "=SUM(Daten!B14:B19)",
                "E2": "=SUM(Daten!B20:B25)"})
    assert not any(k.startswith("blocks:") for _, k in _hints(("Daten", daten), ("KPI", kpi)))


def test_a_block_that_takes_a_row_of_the_next_one_counts_it_twice():
    """Waagrecht: Quartale aus Monaten in einer Zeile, Q3 nimmt den Juni mit."""
    cells = {"A1": "Monat", "A2": "Umsatz"}
    for c in range(2, 14):
        cells[f"{column_letters(c)}1"], cells[f"{column_letters(c)}2"] = f"M{c - 1}", float(100 + 7 * c)
    cells.update({"A4": "Quartal", "B4": "Q1", "C4": "Q2", "D4": "Q3", "E4": "Q4", "A5": "Umsatz",
                  "B5": "=SUM(B2:D2)", "C5": "=SUM(E2:G2)", "D5": "=SUM(G2:J2)", "E5": "=SUM(K2:M2)"})
    hint = _hints(("Plan", cells))[("Plan", "blocks:D5")]
    assert "G2 zählt dadurch zweimal, auch im Block daneben" in hint["text"]
    assert _fixed(hint) == {"D5": "=SUM(H2:J2)"}


def test_a_year_column_that_is_not_a_sum_stays_quiet():
    """„Jahr“ neben den Quartalen als Mittelwert (Auslastung) oder als Jahreszahl ist keine Summe der Quartale."""
    cells = {"A1": "Kennzahl", "B1": "Q1", "C1": "Q2", "D1": "Q3", "E1": "Q4", "F1": "Jahr",
             "A2": "Auslastung", "B2": 0.81, "C2": 0.77, "D2": 0.85, "E2": 0.79, "F2": "=AVERAGE(B2:E2)",
             "A3": "Plan", "B3": 10.0, "C3": 12.0, "D3": 11.0, "E3": 9.0, "F3": 2026}
    assert not any(k.startswith("parts:") for _, k in _hints(("KPI", cells)))


def test_a_year_column_with_the_closing_balance_stays_quiet():
    """Im Liquiditätsplan zählt „Jahr“ Einnahmen und Ausgaben zusammen, beim Kontostand steht der Stand im Dezember und
    bei der Veränderung Dezember minus Januar: keine vergessenen Monate."""
    months = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]
    cells = {"A1": "Plan", "N1": "Jahr", "A2": "Einnahmen", "A3": "Ausgaben", "A4": "Kontostand", "A5": "Veränderung"}
    for i, m in enumerate(months):
        col = column_letters(i + 2)
        cells[f"{col}1"], cells[f"{col}2"], cells[f"{col}3"] = m, float(1000 + 50 * i), float(900 + 30 * i)
        cells[f"{col}4"] = f"={'B2-B3' if i == 0 else f'{column_letters(i + 1)}4+{col}2-{col}3'}"
        cells[f"{col}5"] = f"={col}4-B4"
    cells.update({"N2": "=SUM(B2:M2)", "N3": "=SUM(B3:M3)", "N4": "=M4", "N5": "=M4-B4"})
    assert not any(k.startswith("parts:") for _, k in _hints(("Plan", cells)))
