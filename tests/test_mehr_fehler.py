"""Fehler, die eine selbst gebaute Testreihe (Labor, Lager, Rechnungen) zuerst übersah: ein Referenzwert, der nicht zu
seinem Parameter passt, ein Lieferant, der nicht zur Kategorie passt, eine kopierte Rechnungsnummer und ein alter
MwSt-Satz. Jeder mit der Korrektur zum Einsetzen, und keiner, wo die Mappe stimmt."""

from test_workbook import xlsx

from voiceagent import pruefer

STYLES = '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="10" applyNumberFormat="1"/></cellXfs>'
PARAMS = {"Kreatinin": ("µmol/l", 45, 110), "Hämoglobin": ("g/l", 120, 170), "Kalium": ("mmol/l", 3.5, 5.1),
          "Glukose": ("mmol/l", 3.9, 6.1), "CRP": ("mg/l", 0, 5)}


def _hints(name: str, cells: dict) -> dict:
    done = pruefer.Checker(None).check_file("mappe.xlsx", xlsx((name, cells), styles=STYLES))
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


def _labor(**changes) -> dict:
    cells = {f"{c}1": h for c, h in zip("ABCDEFGH", ["Probe-Nr", "Patient", "Parameter", "Wert", "Einheit", "Ref. min",
                                                     "Ref. max", "Status"])}
    r = 2
    for i in range(10):
        for p, (unit, lo, hi) in PARAMS.items():
            value = round(lo + (hi - lo) * ((i * 7 + r) % 10) / 9, 1)
            cells.update({f"A{r}": f"L{50000 + r}", f"B{r}": f"P-{300 + i}", f"C{r}": p, f"D{r}": value, f"E{r}": unit,
                          f"F{r}": float(lo), f"G{r}": float(hi)})
            r += 1
    cells.update(changes)
    for row in range(2, r):
        value, lo, hi = cells[f"D{row}"], cells[f"F{row}"], cells[f"G{row}"]
        status = "tief" if value < lo else "hoch" if value > hi else "normal"
        cells[f"H{row}"] = ("=", f'IF(D{row}<F{row},"tief",IF(D{row}>G{row},"hoch","normal"))', status)
    return cells


def test_a_reference_value_that_does_not_fit_its_parameter_gets_the_right_number():
    """Zu Kalium gehört 5,1 als oberer Referenzwert, und zu 3,5 unten auch: 5,5 in einer Kalium-Zeile ist falsch."""
    row = next(r for r in range(2, 52) if _labor()[f"C{r}"] == "Kalium")
    hint = _hints("Labor", _labor(**{f"D{row}": 5.3, f"G{row}": 5.5}))[f"mapping:G{row}"]
    assert hint["level"] == "fehler"
    assert f"G{row} (Ref. max) 5,5 passt nicht zum Rest der Zeile" in hint["text"]
    assert "Zu „Kalium“ (Parameter) gehört sonst immer 5,1 (9-mal), zu 3,5 (Ref. min) ebenso" in hint["text"]
    # Der Status rechnet mit dem falschen Wert: Das steht als Folge dabei.
    assert f"H{row} (Status" in hint["text"] and "„normal“ statt „hoch“" in hint["text"]
    assert hint["text"].endswith("Richtig wäre 5,1 statt 5,5.")
    assert hint["fix"]["cells"] == {f"G{row}": 5.1} and "text" not in hint["fix"]


def test_a_clean_lab_list_has_nothing_to_say():
    assert _hints("Labor", _labor()) == {}


def _lager(**changes) -> dict:
    cats = {"Elektronik": "Techno AG", "Büro": "Papeterie Muster", "Werkzeug": "Hobel GmbH"}
    cells = {f"{c}1": h for c, h in zip("ABCDE", ["Artikel-Nr", "Bezeichnung", "Kategorie", "Lieferant", "Menge"])}
    r = 2
    for k in range(8):
        for cat, supplier in cats.items():
            cells.update({f"A{r}": f"A-{2000 + r}", f"B{r}": f"{cat} {k + 1}", f"C{r}": cat, f"D{r}": supplier,
                          f"E{r}": float(3 + (r * 37) % 200)})
            r += 1
    cells.update(changes)
    return cells


def test_two_odd_rows_in_a_short_list_are_still_found():
    """24 Artikel, drei Kategorien, je ein Lieferant: Zwei Zeilen passen nicht dazu. Welche Zelle falsch ist, sagt
    die Zeile nicht, darum stehen beide Korrekturen zur Wahl."""
    found = _hints("Lager", _lager(D23="Papeterie Muster", D12="Techno AG"))
    hint = found["pair:C23"]
    assert hint["cell"] == "C23,D23" and hint["level"] == "fehler"
    assert "Zu „Elektronik“ gehört sonst immer „Techno AG“ (7-mal)" in hint["text"]
    assert [c["fix"]["cells"] for c in hint["choices"]] == [{"D23": "Techno AG"}, {"C23": "Büro"}]
    assert "pair:C12" in found


def test_a_copied_number_in_a_running_sequence_gets_the_missing_number():
    found = _hints("Lager", _lager(A14="A-2009"))
    hint = found["id:A14"]
    assert hint["level"] == "fehler"
    assert "A14 (Artikel-Nr) „A-2009“ steht schon in A9" in hint["text"] and "„A-2014“ fehlt" in hint["text"]
    assert hint["fix"]["cells"] == {"A14": "A-2014"} and hint["fix"]["text"]
    assert not any(k.startswith("id:") and k != "id:A14" for k in found)


def test_plain_numbers_and_leading_zeros_are_kept():
    numbers = _lager(**({f"A{r}": float(1000 + r) for r in range(2, 26)} | {"A14": 1009.0}))
    assert _hints("Lager", numbers)["id:A14"]["fix"]["cells"] == {"A14": 1014.0}
    padded = _lager(**({f"A{r}": f"F-{r:04d}" for r in range(2, 26)} | {"A14": "F-0009"}))
    assert _hints("Lager", padded)["id:A14"]["fix"]["cells"] == {"A14": "F-0014"}


def test_a_duplicate_without_a_sequence_is_only_a_question():
    """Nummern ohne Reihenfolge: Welche die richtige wäre, ist nicht zu sagen, also nur die Frage, ohne Einsetzen."""
    shuffled = _lager(**{f"A{r}": f"A-{(r * 7919) % 9000 + 1000}" for r in range(2, 26)})
    shuffled["A14"] = shuffled["A9"]
    hint = _hints("Lager", shuffled)["id:A9"]
    assert hint["level"] == "pruefen" and "fix" not in hint and "in Zeilen 9, 14" in hint["text"]


def _rechnungen(rates: dict) -> dict:
    cells = {f"{c}1": h for c, h in zip("ABCDE", ["Rechnungs-Nr", "Kunde", "Netto", "MwSt-Satz", "MwSt"])}
    total = 0.0
    for r in range(2, 32):
        net, rate = 100.0 + 10 * r, rates.get(r, 0.081)
        tax = net * rate
        total += tax
        cells.update({f"A{r}": f"R-{26100 + r}", f"B{r}": f"Kunde {r % 6}", f"C{r}": net, f"D{r}": (rate, 1),
                      f"E{r}": ("=", f"C{r}*D{r}", tax)})
    cells["G2"], cells["H2"] = "MwSt gesamt", ("=", "SUM(E2:E31)", total)
    return cells


def test_an_old_rate_among_the_current_one_is_a_question_with_the_current_rate():
    hint = _hints("Rechnungen", _rechnungen({9: 0.077}))["rate:D9"]
    assert hint["level"] == "pruefen"
    assert "D9 (MwSt-Satz) 7,7 % weicht ab: In den übrigen 29 Zeilen steht 8,1 %" in hint["text"]
    assert "H2 (MwSt gesamt) 643,19 statt 643,95" in hint["text"]
    assert hint["text"].endswith("Richtig wäre 8,1 % statt 7,7 %.")
    assert hint["fix"]["cells"] == {"D9": 0.081} and hint["fix"]["format"] == {"D9": "D2"}


def test_a_clearly_different_or_zero_rate_is_left_alone():
    """7 % neben 19 % (ermäßigt) und 0 % (steuerfrei) sind meist Absicht."""
    assert not any(k.startswith("rate:") for k in _hints("Rechnungen", _rechnungen({9: 0.0, 15: 0.026})))
    assert _hints("Rechnungen", _rechnungen({})) == {}


def test_an_error_value_in_a_sentence_is_written_as_excel_shows_it():
    """„6'480“ als Text: SUMME lässt es weg, =G19*2 zeigt #WERT!. Im Satz steht #WERT!, nicht Err(...)."""
    cells = {"A1": "Name", "B1": "Lohn", "C1": "Doppelt"}
    for r in range(2, 14):
        cells.update({f"A{r}": f"Person {r}", f"B{r}": 5000.0 + r, f"C{r}": ("=", f"B{r}*2", 2 * (5000.0 + r))})
    cells["B7"], cells["C7"] = "6'480", ("=", "B7*2", "#VALUE!")
    total = sum(2 * (5000.0 + r) for r in range(2, 14) if r != 7)
    cells["E2"], cells["F2"] = "Summe", ("=", "SUM(C2:C13)", "#VALUE!")
    hint = _hints("Lohn", cells)["astext:B7"]
    assert "Err(" not in hint["text"] and "#WERT! statt" in hint["text"]
    assert total > 0
