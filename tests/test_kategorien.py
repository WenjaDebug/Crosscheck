"""Textspalten, deren Werte zusammenpassen müssen (Testmappe, Tabelle4): Behandlungen zweier Praxen mit Arzt,
Fachbereich, ICD-10-Code und Diagnose. Ein Tippfehler im Namen, eine klein geschriebene Abteilung, die auch nicht zum
Arzt passt, und ein Code, der nicht zur Diagnose passt."""

import random

from test_workbook import xlsx

from voiceagent import ansicht, einsetzen, pruefer
from voiceagent.kategorien import _rule, _typo
from voiceagent.xlsx import Cell

DOCS = {"Dr. Schmidt": ("Kardiologie", {"I48.9": "Vorhofflimmern", "I50.9": "Herzinsuffizienz", "I25.1": "Koronare Herzkrankheit"}),
        "Dr. Weber": ("Dermatologie", {"L20.9": "Neurodermitis", "B35.1": "Nagelmykose", "L40.0": "Psoriasis vulgaris"}),
        "Dr. Müller": ("Innere Medizin", {"E11.9": "Diabetes mellitus Typ 2", "I10": "Essentielle Hypertonie",
                                          "J06.9": "Akute Infektion der oberen Atemwege"})}
STYLES = '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs>'


def _book(**changes) -> dict:
    rnd = random.Random(7)
    cells: dict = {}
    for c, h in zip("ABCDEFGHI", ["Datum", "Praxis", "Arzt", "Fachbereich", "Patienten-ID", "ICD-10", "Diagnose",
                                  "Kostenträger", "Betrag (EUR)"]):
        cells[f"{c}1"] = h
    total = 0.0
    for r in range(2, 62):
        doc = list(DOCS)[r % 3]
        dept, codes = DOCS[doc]
        code = sorted(codes)[(r // 3) % 3]
        amount = round(rnd.uniform(20, 180), 2)
        total += amount
        cells.update({f"A{r}": (46000 + 3 * r, 1), f"B{r}": rnd.choice(["Praxis Nord", "Praxis Mitte"]), f"C{r}": doc,
                      f"D{r}": dept, f"E{r}": f"P-{rnd.randint(1000, 1400)}", f"F{r}": code, f"G{r}": codes[code],
                      f"H{r}": rnd.choice(["GKV", "GKV", "PKV", "Selbstzahler"]), f"I{r}": amount})
    cells["K2"], cells["L2"] = "Behandlungen Dr. Schmidt", ("=", 'COUNTIF(C2:C61,"Dr. Schmidt")', 20.0)
    cells["K3"], cells["L3"] = "Umsatz gesamt", ("=", "SUM(I2:I61)", round(total, 2))
    cells.update(changes)
    return cells


def _hints(cells: dict) -> dict:
    done = pruefer.Checker(None).check_file("behandlungen.xlsx", xlsx(("Tabelle4", cells), styles=STYLES))
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


def test_a_clean_list_of_treatments_has_nothing_to_say():
    """Dieselbe Person kommt in einer Liste von Behandlungen mit Datum mehrmals vor: Das ist kein Fehler."""
    assert _hints(_book(E9="P-1001", E30="P-1001")) == {}


def test_a_typo_in_a_name_is_found_and_the_rest_of_the_row_says_which_name():
    row = next(r for r in range(2, 62) if list(DOCS)[r % 3] == "Dr. Schmidt")
    found = _hints(_book(**{f"C{row}": "Dr. Shmidt"}))
    hint = found[f"typo:C{row}"]
    assert hint["level"] == "fehler"
    assert "„Dr. Shmidt“ ist wohl ein Tippfehler für „Dr. Schmidt“" in hint["text"]
    assert "Zu „Kardiologie“ (Fachbereich) gehört sonst immer „Dr. Schmidt“" in hint["text"]
    # ZÄHLENWENN in L2 zählt die Zeile nicht mit: Das steht als Folge dabei.
    assert "Darum ergibt L2 (Behandlungen Dr. Schmidt) 19 statt 20" in hint["text"]
    assert hint["fix"]["cells"] == {f"C{row}": "Dr. Schmidt"} and hint["fix"]["text"]
    assert hint["compare"]["now"] == "„Dr. Shmidt“" and hint["compare"]["want"] == "„Dr. Schmidt“"


def test_a_department_that_fits_neither_the_doctor_nor_the_diagnosis():
    row = next(r for r in range(2, 62) if list(DOCS)[r % 3] == "Dr. Müller")
    found = _hints(_book(**{f"D{row}": "kardiologie"}))
    hint = found[f"mapping:D{row}"]
    assert hint["level"] == "fehler"
    assert "„kardiologie“ ist klein geschrieben und passt auch sonst nicht zum Rest der Zeile" in hint["text"]
    assert "Zu „Dr. Müller“ (Arzt) gehört sonst immer „Innere Medizin“" in hint["text"]
    assert "zählt die Zeile darum bei „Kardiologie“ statt bei „Innere Medizin“" in hint["text"]
    assert hint["fix"]["cells"] == {f"D{row}": "Innere Medizin"}


def test_a_code_that_does_not_fit_its_diagnosis_offers_both_corrections():
    row = next(r for r in range(2, 62) if list(DOCS)[r % 3] == "Dr. Müller" and _book()[f"F{r}"] == "E11.9")
    found = _hints(_book(**{f"G{row}": "Essentielle Hypertonie"}))
    hint = found[f"pair:F{row}"]
    assert hint["level"] == "fehler" and "fix" not in hint and hint["cell"] == f"F{row},G{row}"
    assert "Zu „E11.9“ gehört sonst immer „Diabetes mellitus Typ 2“" in hint["text"]
    assert "zu „Essentielle Hypertonie“ sonst immer „I10“" in hint["text"]
    assert [c["fix"]["cells"] for c in hint["choices"]] == [{f"G{row}": "Diabetes mellitus Typ 2"}, {f"F{row}": "I10"}]
    # Zwei Zellen: Nichts davon gehört als Ort unter den Titel.
    assert "where" not in hint and hint["title"] == "Zwei Angaben widersprechen sich"


def test_a_name_with_a_trailing_space_or_other_case_is_one_and_the_same():
    rows = [r for r in range(2, 62) if list(DOCS)[r % 3] == "Dr. Weber"]
    found = _hints(_book(**{f"C{rows[0]}": "Dr. Weber ", f"H{rows[1]}": "gkv"}))
    space = found[f"spelling:C{rows[0]}"]
    assert "„Dr. Weber “ hat ein Leerzeichen am Ende" in space["text"] and space["fix"]["cells"] == {f"C{rows[0]}": "Dr. Weber"}
    case = found[f"spelling:H{rows[1]}"]
    assert "„gkv“ ist klein geschrieben" in case["text"] and "R, Python und SPSS" in case["text"]


def test_an_amount_with_a_currency_sign_is_a_number_as_text():
    found = _hints(_book(I23="85,50 €"))
    hint = found["astext:I23"]
    assert hint["level"] == "fehler" and hint["fix"]["cells"] == {"I23": 85.5}
    assert "L3 (Umsatz gesamt)" in hint["text"]


def test_two_columns_that_only_happen_to_agree_are_no_rule():
    """Zehn Zeilen, in denen Geschlecht und Raucher fast immer zusammengehen: Zufall, keine Regel."""
    sex = {r: "m" if r < 5 else "w" for r in range(10)}
    smoker = {r: ("ja" if r < 5 else "nein") for r in range(10)}
    smoker[9] = "ja"
    assert _rule(sex, smoker) is None
    codes = {r: f"C{r % 4}" for r in range(40)}
    names = {r: f"Name {r % 4}" for r in range(40)}
    names[7] = "Name 0"
    assert _rule(codes, names) is not None


def test_what_counts_as_a_typo():
    assert _typo("Dr. Shmidt", "Dr. Schmidt") and _typo("Kardiolgie", "Kardiologie")
    assert not _typo("Station 3", "Station 4") and not _typo("Gruppe A", "Gruppe B") and not _typo("GKV", "PKV")


def test_a_text_correction_is_applied_as_text():
    cells = einsetzen.applied({(1, 1): Cell("Dr. Shmidt")}, [{"cells": {"A1": "Dr. Schmidt"}, "was": {}}])
    assert cells[(1, 1)].value == "Dr. Schmidt" and cells[(1, 1)].formula is None


def test_the_place_of_a_quoted_name_goes_under_the_title():
    view = ansicht.view({"key": "typo:C17", "cell": "C17", "level": "fehler",
                         "text": "C17 (Arzt) „Dr. Shmidt“ ist wohl ein Tippfehler für „Dr. Schmidt“: … "
                                 "Richtig wäre „Dr. Schmidt“ statt „Dr. Shmidt“."})
    assert view["where"] == "Arzt" and view["lead"].startswith("„Dr. Shmidt“ ist wohl")


def test_times_keep_their_seconds_and_are_rounded_not_cut():
    """0,395833 ist 9:29:59,97: als 09:30 gelesen, nicht als 09:29. Eine Zeit mit Sekunden behält sie."""
    from voiceagent import formel
    from voiceagent.xlsx import _date_text

    assert _date_text(0.395833, False) == "09:30"
    assert _date_text((8 * 3600 + 15 * 60 + 42) / 86400, False) == "08:15:42"
    assert abs(formel.serial_of("08:15:42") - (8 * 3600 + 15 * 60 + 42) / 86400) < 1e-12
    assert formel.serial_of("01.10.2024 08:30:15") == formel.serial_of("01.10.2024") + (8 * 3600 + 30 * 60 + 15) / 86400


def test_a_time_sheet_read_to_the_second_has_no_false_recalc_alarm():
    """Beginn und Ende mit Rundungsresten unter einer Sekunde: Die Stunden rechnen nach wie in Excel."""
    styles = '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="20" applyNumberFormat="1"/></cellXfs>'
    cells = {"A1": "Mitarbeiter", "B1": "Beginn", "C1": "Ende", "D1": "Stunden"}
    for r in range(2, 14):
        start, end = round(8 / 24, 6), round((8 + 7.5 + r % 3 * 0.25) / 24, 6)
        cells.update({f"A{r}": ["Lea", "Max", "Noah"][r % 3], f"B{r}": (start, 1), f"C{r}": (end, 1),
                      f"D{r}": ("=", f"(C{r}-B{r})*24", (end - start) * 24)})
    done = pruefer.Checker(None).check_file("zeiten.xlsx", xlsx(("Zeiten", cells), styles=styles))
    assert not [h for s in done["sheets"] for h in s["hints"] if h["key"].startswith("recalc:")]
