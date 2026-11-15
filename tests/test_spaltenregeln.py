"""Eigene Regeln je Spalte: anlegen, prüfen, mit Zelle und Grund melden, speichern, im Seitenfenster und im Bericht."""

import json
import re
from datetime import date

import pytest
from test_auswertung import faelle
from test_excel_live import FakeExcel, client_with

from voiceagent import bericht, spaltenregeln
from voiceagent.analyse import Sheet as Analysed
from voiceagent.analyse import analyse
from voiceagent.pruefer import Checker
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Sheet

HEADERS = {"X-VoiceAgent": "1"}


def rule(**fields):
    return spaltenregeln.clean(fields)


def run(cells, *rules, name="Fälle"):
    return analyse(cells, layout(Sheet(name, cells)), name=name, rules=list(rules))


def ruled(report):
    return {f.key: f for f in report.findings if f.key.startswith("regel:")}


def one(report):
    found = ruled(report)
    assert len(found) == 1, found
    return next(iter(found.values()))


def table(header, values, others=None):
    """Eine Tabelle mit Fall-Nr und einer Spalte header (values als Cell oder Text), darunter nichts."""
    cells = {(1, 1): Cell("Fall-Nr"), (1, 2): Cell(header)}
    for i, value in enumerate(values, start=2):
        cells[(i, 1)] = Cell(str(1000 + i))
        if value is not None:
            cells[(i, 2)] = value if isinstance(value, Cell) else Cell(value)
    for col, (name, column) in enumerate((others or {}).items(), start=3):
        cells[(1, col)] = Cell(name)
        for i, value in enumerate(column, start=2):
            cells[(i, col)] = value if isinstance(value, Cell) else Cell(value)
    return cells


# --- Regeln lesen ---


def test_numbers_and_dates_are_read_like_people_type_them():
    assert spaltenregeln.number("1,5") == 1.5 and spaltenregeln.number("37.5") == 37.5
    assert spaltenregeln.number("1'000") == 1000 and spaltenregeln.number("1.234,5") == 1234.5
    assert spaltenregeln.number("1,234.5") == 1234.5 and spaltenregeln.number("−3") == -3
    assert spaltenregeln.number("abc") is None and spaltenregeln.number("") is None
    assert spaltenregeln.day("2020-01-31") == "2020-01-31" and spaltenregeln.day("1.2.2020") == "2020-02-01"
    assert spaltenregeln.day("31.02.2020") is None


def test_a_rule_is_checked_before_it_is_saved():
    found = rule(spalte="  Alter ", pflicht=1, art="ganz", min="0", max="120,5", werte="0,5; 1; 1,5")
    assert found["spalte"] == "Alter" and found["max"] == 120.5 and found["werte"] == ["0,5", "1", "1,5"]
    assert rule(spalte="G", werte="m, w, d")["werte"] == ["m", "w", "d"]
    for bad, says in [
        ({"spalte": ""}, "Zu welcher Spalte"),
        ({"spalte": "Alter"}, "mindestens eine Bedingung"),
        ({"spalte": "Alter", "min": "200", "max": "100"}, "Von (200) ist größer als bis (100)"),
        ({"spalte": "Alter", "min": "zwölf"}, "„zwölf“ unter Von ist keine Zahl"),
        ({"spalte": "Alter", "art": "datum", "min": "1"}, "passen nicht zur Art Datum"),
        ({"spalte": "Alter", "art": "zahl", "keine_zukunft": True}, "passen nicht zur Art Zahl"),
        ({"spalte": "Austritt", "spalte2": "austritt"}, "mit sich selbst"),
        ({"spalte": "Eintritt", "nicht_vor": "2025-01-01", "nicht_nach": "2024-01-01"}, "liegt nach"),
        ({"spalte": "Eintritt", "nicht_vor": "32.1.2024"}, "ist kein Datum"),
    ]:
        with pytest.raises(ValueError, match=re.escape(says)):
            spaltenregeln.clean(bad)


def test_a_rule_says_in_words_what_it_wants():
    assert spaltenregeln.describe(rule(spalte="Alter", pflicht=True, art="ganz", min=0, max=120)) == "Pflichtfeld, ganze Zahl von 0 bis 120"
    assert spaltenregeln.describe(rule(spalte="Dosis", min="0,5")) == "Zahl ab 0,5"
    assert spaltenregeln.describe(rule(spalte="G", werte="m, w, d")) == "nur „m“, „w“ oder „d“"
    assert spaltenregeln.describe(rule(spalte="Austritt", art="datum", spalte2="Eintritt", keine_zukunft=True,
                                       nicht_vor="2020-01-01")) == "Datum, nicht vor dem 01.01.2020, nicht in der Zukunft, nicht vor „Eintritt“"
    assert spaltenregeln.describe(rule(spalte="Dia", art="zahl", spalte2="Sys", beziehung="le")) == "Zahl, nicht größer als „Sys“"
    assert spaltenregeln.describe(rule(spalte="Nr", einmal=True, muster="PAT-####")) == "jeder Wert nur einmal, Muster „PAT-####“"


# --- Prüfen ---


def test_a_clean_case_list_keeps_its_own_rules():
    rules = [rule(spalte="Alter", pflicht=True, art="ganz", min=0, max=120), rule(spalte="Fall-Nr", einmal=True, muster="######"),
             rule(spalte="Station", werte="A, B, C"), rule(spalte="Austritt", art="datum", spalte2="Eintritt", keine_zukunft=True),
             rule(spalte="Gewicht kg", pflicht=True, art="zahl", min=30, max=250), rule(spalte="Eintritt", nicht_vor="2024-01-01")]
    assert ruled(run(faelle(), *rules)) == {}


def test_each_broken_rule_says_where_and_why():
    cells = faelle(dupe_id=True, code=True, impossible=True, order=True, missing=True)
    cells[(7, 9)], cells[(8, 9)] = Cell("a"), Cell("B ", text="B ")
    cells[(9, 1)] = Cell("24-0108")
    report = run(cells, rule(spalte="Alter", pflicht=True, art="ganz", min=0, max=120), rule(spalte="fall-nr", einmal=True, muster="######"),
                 rule(spalte="Gewicht kg", pflicht=True, art="zahl", min=30, max=250), rule(spalte="Station", werte="A, B, C"),
                 rule(spalte="Austritt", art="datum", spalte2="Eintritt", keine_zukunft=True))
    found = ruled(report)
    assert found["regel:E1"].text == (
        "Deine Regel für „Alter“ (Pflichtfeld, ganze Zahl von 0 bis 120): 3 Zellen passen nicht (2 über 120, 1 unter 0): "
        "999 in E5 und E19 ist über 120; -4 in E10 ist unter 0. 999 steht 2-mal da, weit weg von den übrigen Werten: wohl "
        "ein Code für „fehlt“, und MITTELWERT und SUMME rechnen damit.")
    assert found["regel:E1"].cell == "E5,E10,E19" and found["regel:E1"].level == "fehler"
    assert found["regel:E1"].action == {"kind": "rule", "spalte": "Alter", "blatt": ""}
    assert found["regel:F1"].text.endswith("4 Zellen passen nicht: F6, F16 und F26 sind leer; „k.A.“ in F35 steht statt eines Werts da.")
    assert found["regel:I1"].text.endswith(
        "„a“ in I7 steht nicht in der Liste (anders geschrieben als „A“); „B “ in I8 steht nicht in der Liste (Leerzeichen am Rand).")
    # Die negative Liegedauer daraus sagt der Hinweis zur Regel mit (die Liegedauer hat keine eigene Regel).
    assert found["regel:D1"].text.endswith("Austritt 01.12.2023 in D11 liegt vor Eintritt 11.02.2024 (C11). "
                                           "Darum zeigt H11 (Liegedauer, Fall-Nr 240109) -72.")
    assert "„24-0108“ in A9 passt nicht zum Muster „######“: zu lang, 7 statt 6 Zeichen." in found["regel:A1"].text
    assert "Ein Wert kommt mehrmals vor: 240105 in A7 und A22." in found["regel:A1"].text
    assert found["regel:A1"].tip.startswith("Ist derselbe Fall doppelt erfasst?")
    # Was die Regeln genauer sagen, steht nicht noch einmal allgemein da.
    kinds = {f.key.split(":")[0] for f in report.findings}
    assert not {"code", "id", "order", "placeholder", "impossible"} & kinds


def test_without_rules_the_general_hints_stay():
    report = run(faelle(dupe_id=True, code=True, impossible=True, order=True, missing=True))
    kinds = {f.key.split(":")[0] for f in report.findings}
    assert {"code", "id", "order", "placeholder", "impossible"} <= kinds and "regel" not in kinds


def test_an_order_hint_for_another_pair_stays():
    cells = faelle(order=True)
    report = run(cells, rule(spalte="Austritt", art="datum", spalte2="Geburtsdatum"))
    assert any(f.key.startswith("order:") for f in report.findings) and not ruled(report)


def test_many_equal_values_make_one_sentence():
    values = ["Chirurgie", "Medizin", "Medizin", "Chirurgie"] * 5 + ["Medizin "] + [None]
    found = one(run(table("Station", values), rule(spalte="Station", werte="Chirurgie, Orthopädie")))
    assert found.text == (
        "Deine Regel für „Station“ (nur „Chirurgie“ oder „Orthopädie“): 11 Zellen passen nicht: „Medizin“ in B3, B4, B7 "
        "und 7 weiteren Zellen steht nicht in der Liste; „Medizin “ in B22 steht nicht in der Liste.")
    assert found.cell == "B3:B4,B7:B8,B11:B12,B15:B16,B19:B20,B22"


def test_beyond_five_sentences_the_rest_is_counted():
    values = [str(v) for v in (150, 151, 152, 153, 154, 155, 156)] + ["20"] * 5
    found = one(run(table("Alter", values), rule(spalte="Alter", max=120)))
    assert found.text.endswith("7 Zellen passen nicht: 150 in B2 ist über 120; 151 in B3 ist über 120; 152 in B4 ist über "
                               "120; 153 in B5 ist über 120; 154 in B6 ist über 120. Dazu 2 weitere Zellen.")


def test_kinds_say_what_is_there_instead():
    values = ["12", Cell("12", text="12"), "WAHR", "01.10.2024", "#NV", "2,5"]
    found = one(run(table("Dosis", values + ["1"] * 4), rule(spalte="Dosis", pflicht=True, art="ganz")))
    assert found.text.endswith(
        "5 Zellen passen nicht (1 Zahl als Text, 1 keine Zahl, 1 Datum statt Zahl, 1 Fehlerwert, 1 keine ganze Zahl): „12“ in B3 "
        "ist als Text gespeichert, nicht als Zahl; „WAHR“ in B4 ist keine Zahl; 01.10.2024 in B5 ist ein Datum, keine Zahl; "
        "#NV in B6 ist ein Fehlerwert; 2,5 in B7 ist keine ganze Zahl."), found.text


def test_a_date_column_tells_numbers_and_text_from_dates():
    values = ["01.03.2024", "45000", Cell("01.10.2024", text="01.10.2024"), "gestern", "03.02.2019", "05.06.2031"]
    found = one(run(table("Eintritt", values + ["01.04.2024"] * 4),
                    rule(spalte="Eintritt", art="datum", nicht_vor="2020-01-01", keine_zukunft=True), name="F"))
    for part in ["45000 in B3 ist eine Zahl, kein Datum (als Datum formatiert wäre es der 15.03.2023)",
                 "„01.10.2024“ in B4 ist als Text gespeichert, Excel rechnet damit nicht wie mit einem Datum",
                 "„gestern“ in B5 ist kein Datum", "03.02.2019 in B6 liegt vor dem 01.01.2020", "05.06.2031 in B7 liegt in der Zukunft"]:
        assert part in found.text, (part, found.text)


def test_the_future_is_today():
    values = ["01.03.2026", "02.03.2026"] + ["01.01.2026"] * 8
    cells = table("Termin", values)
    sheet = Analysed(cells, layout(Sheet("T", cells)), "T")
    found, _ = spaltenregeln.checks(sheet, [rule(spalte="Termin", keine_zukunft=True)], [], today=date(2026, 3, 1))
    assert found[0].text.endswith("1 Zelle passt nicht: 02.03.2026 in B3 liegt in der Zukunft.")


def test_patterns_say_which_character_is_wrong():
    values = ["PAT-0001", "pat-0002", "PAT-00003", "PAX-0004", "PAT-00A5", "PAT-0006"]
    found = one(run(table("Patient", values + ["PAT-0010"] * 4), rule(spalte="Patient", muster="PAT-####")))
    for part in ["„pat-0002“ in B3 passt nicht zum Muster „PAT-####“: nur anders geschrieben (groß oder klein)",
                 "„PAT-00003“ in B4 passt nicht zum Muster „PAT-####“: zu lang, 9 statt 8 Zeichen",
                 "„PAX-0004“ in B5 passt nicht zum Muster „PAT-####“: Zeichen 3 ist „X“, dort gehört „T“ hin",
                 "„PAT-00A5“ in B6 passt nicht zum Muster „PAT-####“: Zeichen 7 ist „A“, dort gehört eine Ziffer hin"]:
        assert part in found.text, (part, found.text)
    zeros = one(run(table("PLZ", ["8001", "3000", "123"] + ["8002"] * 7), rule(spalte="PLZ", muster="####")))
    assert "123 in B4 passt nicht zum Muster „####“: zu kurz, 3 statt 4 Zeichen. Als Zahl hat Excel die Nullen vorne " \
           "weggenommen; als Text gespeichert bleiben sie" in zeros.text
    star = run(table("Code", ["A-1", "A-22", "B-3"] + ["A-9"] * 7), rule(spalte="Code", muster="A-*"))
    assert "„B-3“ in B4 passt nicht zum Muster „A-*“." in one(star).text


def test_allowed_numbers_and_numbers_as_text():
    values = ["1", "2", "3", Cell("2", text="2"), "1,0"]
    found = one(run(table("Geschlecht", values + ["1"] * 5), rule(spalte="Geschlecht", werte="1; 2")))
    assert found.text.endswith("2 Zellen passen nicht: 3 in B4 steht nicht in der Liste; „2“ in B5 steht nicht in der Liste "
                               "(als Text gespeichert, nicht als Zahl).")
    # Nur Text in der Spalte: „001“ als Code ist genau so gemeint.
    codes = table("Code", [Cell(v, text=v) for v in ["001", "002", "003"] + ["001"] * 7])
    assert one(run(codes, rule(spalte="Code", werte="001, 002"))).text.endswith("„003“ in B4 steht nicht in der Liste.")
    short = table("Code", [Cell(v, text=v) for v in ["001", "1"] + ["002"] * 8])
    assert one(run(short, rule(spalte="Code", werte="001, 002"))).text.endswith(
        "„1“ in B3 steht nicht in der Liste (anders geschrieben als „001“).")


def test_percent_cells_are_compared_in_percent():
    values = ["12 %", "98 %", "120 %"] + ["50 %"] * 7
    found = one(run(table("Anteil", values), rule(spalte="Anteil", min=0, max=100)))
    assert found.text.endswith("1 Zelle passt nicht: 120 % in B4 ist über 100.")


def test_comparing_two_number_columns():
    cells = table("Diastolisch", ["80", "95", "70"] + ["75"] * 7, {"Systolisch": ["120", "90", "110"] + ["130"] * 7})
    found = one(run(cells, rule(spalte="Diastolisch", art="zahl", spalte2="Systolisch", beziehung="le")))
    assert found.text.endswith("1 Zelle passt nicht: Diastolisch 95 in B3 ist größer als Systolisch 90 (C3).")


def test_blank_rows_and_total_rows_are_no_cases():
    values = ["5", "7", None, "9"]
    cells = table("Tage", values)
    cells[(4, 1)] = Cell("")
    del cells[(4, 1)]
    cells[(7, 1)], cells[(7, 2)] = Cell("Summe"), Cell("21", "SUM(B2:B5)")
    assert ruled(run(cells, rule(spalte="Tage", pflicht=True, art="ganz", max=10))) == {}


def test_a_rule_for_one_sheet_wins_and_others_do_not_apply():
    cells = table("Alter", ["30", "130"] + ["40"] * 8)
    general, here = rule(spalte="Alter", max=120), rule(spalte="Alter", max=150, blatt="Studie")
    assert ruled(run(cells, general, here, name="Studie")) == {}
    assert one(run(cells, general, here, name="Fälle")).text.startswith("Deine Regel für „Alter“ (Zahl bis 120)")
    elsewhere = rule(spalte="Alter", max=100, blatt="Andere")
    assert ruled(run(cells, elsewhere, name="Fälle")) == {}
    strict = rule(spalte="Alter", max=100, blatt="Studie")
    assert one(run(cells, strict, name="studie")).text.startswith("Deine Regel für „Alter“ auf „Studie“ (Zahl bis 100)")


# --- Speichern ---


def test_rules_are_saved_replaced_and_removed(tmp_path):
    assert spaltenregeln.load(tmp_path) == []
    spaltenregeln.save(tmp_path, {"spalte": "Alter", "max": 120})
    spaltenregeln.save(tmp_path, {"spalte": "alter", "max": 110})
    spaltenregeln.save(tmp_path, {"spalte": "Station", "werte": "A, B"})
    saved = spaltenregeln.load(tmp_path)
    assert [(r["spalte"], r["max"]) for r in saved] == [("alter", 110), ("Station", None)]
    # Nur auf einem Blatt statt auf allen: die alte Regel für alle Blätter ist damit weg.
    spaltenregeln.save(tmp_path, {"spalte": "Alter", "max": 99, "blatt": "Fälle"}, before={"spalte": "alter", "blatt": ""})
    assert [(r["spalte"], r["blatt"]) for r in spaltenregeln.load(tmp_path)] == [("Alter", "Fälle"), ("Station", "")]
    assert [r["spalte"] for r in spaltenregeln.remove(tmp_path, "station")] == ["Alter"]
    assert json.loads((tmp_path / "spaltenregeln.json").read_text(encoding="utf-8"))["regeln"][0]["max"] == 99


def test_a_broken_file_is_kept_aside(tmp_path):
    (tmp_path / "spaltenregeln.json").write_text("{kaputt", encoding="utf-8")
    assert spaltenregeln.load(tmp_path) == []
    spaltenregeln.save(tmp_path, {"spalte": "Alter", "max": 120})
    assert (tmp_path / "spaltenregeln.json.kaputt").read_text(encoding="utf-8") == "{kaputt"
    assert len(spaltenregeln.load(tmp_path)) == 1


# --- Seitenfenster und Bericht ---


def case_payload():
    """Eine kleine Fallliste, wie pruefer.js sie schickt: Daten als Zahl mit Zahlenart „d“."""
    values = [["Fall-Nr", "Alter", "Station", "Eintritt", "Austritt"]]
    for i in range(12):
        values.append([f"F-{100 + i}", 30 + i, "A" if i % 2 else "B", 45300 + i, 45305 + i])
    values[4][1] = 140
    values[6][4] = 45200
    kinds = [[None] * 5] + [[None, None, None, "d", "d"] for _ in range(12)]
    formulas = [[None] * 5 for _ in values]
    sheet = {"name": "Fälle", "row": 1, "col": 1, "values": values, "formulas": formulas, "kinds": kinds}
    return {"book": "Fälle.xlsx", "active": "Fälle", "sheets": [sheet]}


def test_the_pane_makes_changes_and_removes_rules(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    assert client.post("/api/addin/regeln").status_code == 403
    assert client.post("/api/addin/regeln", headers=HEADERS).json() == {"state": "ok", "regeln": []}
    assert client.post("/api/addin/regeln/spalte", json={"sheet": "Fälle", "cell": "B3"}, headers=HEADERS).json()["state"] == "again"
    first = client.post("/api/addin/pruefen", json=case_payload(), headers=HEADERS).json()
    assert not [h for h in first["hints"] if h["key"].startswith("regel:")]
    form = client.post("/api/addin/regeln/spalte", json={"sheet": "Fälle", "cell": "B3"}, headers=HEADERS).json()
    assert form["spalte"] == "Alter" and form["blatt"] == "Fälle" and form["art"] == "ganz" and form["regel"] is None
    assert form["jetzt"] == "12 Zeilen: 12 Zahlen von 30 bis 140. Jeder Wert kommt nur einmal vor."
    assert form["spalten"] == ["Fall-Nr", "Station", "Eintritt", "Austritt"]
    outside = client.post("/api/addin/regeln/spalte", json={"sheet": "Fälle", "cell": "H30"}, headers=HEADERS).json()
    assert outside["state"] == "none"
    bad = client.post("/api/addin/regeln/speichern", json={"regel": {"spalte": "Alter"}}, headers=HEADERS).json()
    assert bad["state"] == "error" and "mindestens eine Bedingung" in bad["text"]
    saved = client.post("/api/addin/regeln/speichern", json={"regel": {"spalte": "Alter", "art": "ganz", "max": "120"}},
                        headers=HEADERS).json()
    assert saved["regeln"][0]["text"] == "ganze Zahl bis 120"
    client.post("/api/addin/regeln/speichern", json={"regel": {"spalte": "Austritt", "art": "datum", "spalte2": "Eintritt"}},
                headers=HEADERS)
    again = client.post("/api/addin/pruefen", json=case_payload(), headers=HEADERS).json()
    hints = {h["key"]: h for h in again["hints"] if h["key"].startswith("regel:")}
    assert hints["regel:B1"]["text"].endswith("1 Zelle passt nicht: 140 in B5 ist über 120.")
    assert hints["regel:B1"]["cell"] == "B5" and hints["regel:B1"]["action"]["kind"] == "rule"
    assert hints["regel:E1"]["text"].endswith("Austritt 01.10.2023 in E7 liegt vor Eintritt 14.01.2024 (D7).")
    # Ändern aus dem Hinweis: die Spalte nach ihrem Namen, mit ihrer Regel.
    edit = client.post("/api/addin/regeln/spalte", json={"sheet": "Fälle", "spalte": "alter", "blatt": ""}, headers=HEADERS).json()
    assert edit["spalte"] == "Alter" and edit["regel"]["max"] == 120 and edit["regel"]["text"] == "ganze Zahl bis 120"
    missing = client.post("/api/addin/regeln/spalte", json={"sheet": "Fälle", "spalte": "Gewicht", "blatt": ""}, headers=HEADERS).json()
    assert missing["jetzt"] == "In der offenen Mappe gibt es keine Spalte „Gewicht“." and missing["regel"] is None
    full = client.post("/api/addin/bericht", json={"daten": True}, headers=HEADERS).json()
    page = client.get(full["url"]).text
    assert "Eigene Regeln je Spalte" in page and "<b>Alter</b> (Fälle): ganze Zahl bis 120" in page
    plain = client.get(client.post("/api/addin/bericht", json={"daten": False}, headers=HEADERS).json()["url"]).text
    assert "Geprüft auch mit 2 eigenen Regeln je Spalte." in plain and "Alter" not in plain
    assert "Eigene Regel für eine Spalte verletzt" in plain
    left = client.post("/api/addin/regeln/loeschen", json={"spalte": "Alter", "blatt": ""}, headers=HEADERS).json()
    assert [r["spalte"] for r in left["regeln"]] == ["Austritt"]
    last = client.post("/api/addin/pruefen", json=case_payload(), headers=HEADERS).json()
    assert [h["key"] for h in last["hints"] if h["key"].startswith("regel:")] == ["regel:E1"]


def test_a_rule_saved_by_the_other_app_counts_at_the_next_check(tmp_path):
    checker = Checker(None, tmp_path)
    assert checker.rules()[0] == []
    spaltenregeln.save(tmp_path, {"spalte": "Alter", "max": 120})
    assert [r["spalte"] for r in checker.rules()[0]] == ["Alter"]
    found = checker.check_file("x.xlsx", b"kein xlsx")
    assert found["state"] == "error"


def test_a_checker_without_a_data_folder_has_no_rules():
    checker = Checker(None)
    assert checker.rules()[0] == [] and checker.rule_save({"spalte": "Alter", "max": 1})["state"] == "error"


def test_the_report_lists_the_rules_with_their_sheets():
    used = spaltenregeln.used({"Fälle": faelle(), "Leer": {}}, [rule(spalte="Alter", max=120), rule(spalte="Puls", max=250)])
    assert [(r["spalte"], r["blaetter"]) for r in used] == [("Alter", ["Fälle"])]
    page = bericht.render("x.xlsx", [("Fälle", {"hints": [{"key": "regel:E1", "text": "t", "level": "fehler", "tip": "",
                                                          "cell": "E5,E10"}]})], False, rules=used)
    assert "Geprüft auch mit 1 eigenen Regel je Spalte." in page and "E5, E10" in page


def test_a_code_only_partly_against_the_rule_keeps_its_own_hint():
    values = ["20", "999", "-1", "999", "-1"] + [str(v) for v in range(30, 40)]
    report = run(table("Alter", values), rule(spalte="Alter", max=120))
    assert "code:B4" in {f.key for f in report.findings}
    assert one(report).text.endswith("999 in B3 und B5 ist über 120. 999 steht 2-mal da, weit weg von den übrigen Werten: "
                                     "wohl ein Code für „fehlt“, und MITTELWERT und SUMME rechnen damit.")


def test_fixing_one_cell_keeps_the_rule_hint_and_the_fix_together():
    """Behebt „Einsetzen“ einen Verstoß und andere bleiben, ist es derselbe Hinweis: „Alle einsetzen“ nimmt die Korrektur."""
    from test_einsetzen import months

    from voiceagent import pruefer

    cells = months([("Mittel Umsatz", "AVERAGE($B$2:$B$11)"), ("Median Kosten", "MEDIAN(C2:C11)")], text={6: "142"})
    rules = [rule(spalte="Umsatz", art="zahl", max=180)]
    hints = pruefer.sheet_result("A.xlsx", "T", {"T": cells}, rules=rules)["hints"]
    mine = next(h for h in hints if h["key"] == "regel:B1")
    assert mine["text"].endswith("3 Zellen passen nicht (2 über 180, 1 Zahl als Text): „142“ in B6 ist als Text gespeichert, "
                                 "nicht als Zahl; 184 in B12 ist über 180; 191 in B13 ist über 180.")
    joint = pruefer.together("A.xlsx", "T", {"T": cells}, hints, rules=rules)
    assert joint["state"] == "ok" and joint["count"] == 3 and joint["left"] == 0


def test_the_form_says_what_is_in_the_column_now():
    cells = faelle(dupe_id=True, missing=True)
    sheet = Analysed(cells, layout(Sheet("Fälle", cells)), "Fälle")
    ids = spaltenregeln.column(sheet, (2, 1), [])
    assert ids["jetzt"] == "40 Zeilen: 40 Zahlen von 240100 bis 240139. 1 Wert kommt mehrmals vor." and ids["art"] == "ganz"
    station = spaltenregeln.column(sheet, (3, 9), [])
    assert station["jetzt"].startswith("40 Zeilen: 40 Text (3 verschieden: ") and "mehrmals" not in station["jetzt"]
    assert station["art"] == ""
    born = spaltenregeln.column(sheet, (1, 2), [rule(spalte="geburtsdatum", keine_zukunft=True)])
    assert born["art"] == "datum" and born["jetzt"].startswith("40 Zeilen: 40 Daten vom ") and born["regel"]["keine_zukunft"]
    weight = spaltenregeln.column(sheet, (5, 6), [])
    assert weight["jetzt"].startswith("40 Zeilen: 36 Zahlen von ") and ", 1 Text (1 verschieden: „k.A.“), 3 leer." in weight["jetzt"]
    shares = table("Anteil", ["12 %", "50 %"] + ["30 %"] * 8)
    share = spaltenregeln.column(Analysed(shares, layout(Sheet("A", shares)), "A"), (3, 2), [])
    assert "Gib von und bis in Prozent an" in share["jetzt"]
    assert spaltenregeln.column(sheet, (60, 2), [])["state"] == "none"
