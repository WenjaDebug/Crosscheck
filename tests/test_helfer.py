"""Excel-Helfer: Auftrag verstehen, Pivot-Übersicht mit Formeln bauen, in ein neues Blatt schreiben, nachprüfen."""

import json
import re
import subprocess
import sys

import pytest
from test_excel_live import client_with

from voiceagent.analyse import value_of
from voiceagent.auftrag import parse
from voiceagent.excel import Snapshot, Status
from voiceagent.formel import Evaluator, _serial
from voiceagent.formel import parse as parse_formula
from voiceagent.helfer import Helper
from voiceagent.pivot import PivotError, build, quote, sheet_name
from voiceagent.schreiben import ExcelWriter, _as, applescript, powershell, verify
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Sheet, column_number, number_text

HEADERS = {1: "Partei", 2: "Jahr", 3: "Sitze", 4: "Stimmen", 5: "Kanton"}


def table(rows: list[list]) -> dict:
    cells = {}
    for i, row in enumerate(rows, start=1):
        for j, value in enumerate(row, start=1):
            if value is None:
                continue
            if isinstance(value, tuple):
                cells[(i, j)] = Cell(value[0], value[1])
            else:
                cells[(i, j)] = Cell(number_text(float(value)) if isinstance(value, (int, float)) else value)
    return cells


def wahl() -> dict:
    return table([
        ["Partei", "Jahr", "Sitze", "Stimmen", "Kanton"],
        ["SP", 2019, 10, 1200.5, "ZH"],
        ["SVP", 2019, 12, 1500, "BE"],
        ["SP", 2023, 11, 1300, "ZH"],
        ["FDP", 2023, 7, 800, "ZH"],
        ["SVP", 2023, 13, 1650, "BE"],
        ["fdp", 2019, "unbekannt", 790, "BE"],
        ["Mitte*", 2023, 3, 400, "ZH"],
        ["Total", None, ("56", "SUM(C2:C8)"), None, None],
    ])


def draft_for(text: str, cells=None, existing=("Daten",)):
    cells = cells or wahl()
    plan = layout(Sheet("Daten", cells))
    understood = parse(text, plan.headers)
    assert understood.auftrag is not None, understood.text
    return build(cells, plan, "Daten", list(existing), understood.auftrag), cells


# --- Nachrechnen, wie Excel es täte ---

OFFSET = 10_000


def excel_values(draft, cells: dict, sheet: str = "Daten") -> dict:
    """Die Formeln der Übersicht mit dem eigenen Formelrechner ausrechnen (die Daten liegen weit unten)."""
    prefix = re.escape(quote(sheet) + "!")

    def local(formula: str) -> str:
        def move(match):
            c1, r1, c2, r2 = match.groups()
            return f"${c1}${int(r1) + OFFSET}:${c2}${int(r2) + OFFSET}"

        return re.sub(prefix + r"\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)", move, formula)

    grid = {(i + 1, j + 1): v for i, row in enumerate(draft.grid) for j, v in enumerate(row) if v != ""}
    done: dict = {}

    def value(pos):
        row, col = pos
        if row > OFFSET:
            cell = cells.get((row - OFFSET, col))
            date = re.fullmatch(r"(\d\d)\.(\d\d)\.(\d{4})", cell.value) if cell is not None else None
            # In Excel ist ein Datum eine Zahl, die Vorschau kennt es als Text.
            return _serial(int(date[3]), int(date[2]), int(date[1])) if date else value_of(cell)
        if pos in done:
            return done[pos]
        entry = grid.get(pos)
        if entry is None:
            result = None
        elif entry.startswith("="):
            result = Evaluator(value).run(parse_formula(local(entry[1:])))
        elif entry.startswith("'"):
            result = entry[1:]
        elif entry in ("TRUE", "FALSE"):
            result = entry == "TRUE"
        else:
            try:
                result = float(entry)
            except ValueError:
                result = entry
        done[pos] = result
        return result

    return {pos: value(pos) for pos in draft.expected}


def same(got, want) -> bool:
    if want == "":
        return got in ("", None)
    return isinstance(got, float) and abs(got - want) < 1e-9


def check_formulas(draft, cells):
    values = excel_values(draft, cells)
    wrong = {pos: (values[pos], want) for pos, want in draft.expected.items() if not same(values[pos], want)}
    assert not wrong


# --- Auftrag verstehen ---


@pytest.mark.parametrize("text, rows, across, werte, order", [
    ("Summe der Sitze pro Partei", [1], None, [("summe", 3)], ""),
    ("Pivot-Tabelle mit Summe Sitze nach Parteien und Jahren", [1, 2], None, [("summe", 3)], ""),
    ("Anzahl pro Kanton", [5], None, [("anzahl", None)], ""),
    ("wie viele Einträge pro Partei", [1], None, [("anzahl", None)], ""),
    ("Durchschnitt und Maximum der Stimmen je Partei, größte zuerst", [1], None, [("mittelwert", 4), ("max", 4)], "absteigend"),
    ("Sitze summiert pro Partei", [1], None, [("summe", 3)], ""),
    ("Summe der Sitze pro Partei, Jahre in den Spalten", [1], 2, [("summe", 3)], ""),
    ("Summe Sitze pro Partei, in den Spalten das Jahr", [1], 2, [("summe", 3)], ""),
    ("Anzahl Partei gegen Kanton", [1], 5, [("anzahl", None)], ""),
    ("Summe der Sitze und Stimmen pro Patei sortiert nach Sitze", [1], None, [("summe", 3), ("summe", 4)], "absteigend"),
    ("Sitze nach Kanton", [5], None, [(None, 3)], ""),
])
def test_parse_understands_simple_sentences(text, rows, across, werte, order):
    result = parse(text, HEADERS)
    assert result.auftrag is not None, result.text
    assert result.auftrag.zeilen == rows and result.auftrag.spalte == across and result.auftrag.sortierung == order
    assert [(w.funktion, w.spalte) for w in result.auftrag.werte] == werte


def test_parse_says_what_is_missing():
    assert "von welcher Spalte" in parse("Summe pro Partei", HEADERS, numeric={3, 4}).text
    assert "Summe von Sitze pro Partei" in parse("Summe pro Partei", HEADERS, numeric={3, 4}).text
    assert "keine Spalte" in parse("mach mir eine Pivot", HEADERS).text
    assert "Wonach soll ich aufteilen" in parse("Summe der Sitze", HEADERS).text


class StrongWriter:
    strong = True

    def __init__(self, answer: dict):
        self.answer, self.calls = answer, []

    def chat(self, system, user, num_predict, json_mode=False, schema=None):
        self.calls.append((system, user, schema))
        return json.dumps(self.answer)


# --- Pivot-Übersicht ---


def test_summary_with_formulas_matches_own_calculation():
    draft, cells = draft_for("Summe der Sitze pro Partei")
    assert draft.name == "Pivot Sitze" and draft.title == "Summe von Sitze pro Partei"
    assert draft.grid[3] == ["Partei", "Summe von Sitze"]
    # FDP und fdp sind für SUMMEWENNS dasselbe, „Total“ unten fällt weg, Text in Sitze zählt nicht.
    assert [r[0] for r in draft.preview] == ["Partei", "FDP", "Mitte*", "SP", "SVP", "Gesamt"]
    assert [r[1] for r in draft.preview] == ["Summe von Sitze", "7", "3", "21", "25", "56"]
    assert any("Zeile 9" in n and "Gesamtsumme" in n for n in draft.notes)
    assert draft.grid[4][1] == "=SUMIFS('Daten'!$C$2:$C$8,'Daten'!$A$2:$A$8,$A5)"
    # „Mitte*“ wäre sonst ein Platzhalter für alles, was mit Mitte anfängt.
    assert draft.grid[5][1] == "=SUMIFS('Daten'!$C$2:$C$8,'Daten'!$A$2:$A$8,\"=Mitte~*\")"
    assert draft.grid[-1] == ["Gesamt", "=SUM(B5:B8)"]
    assert "SUMMEWENNS" in draft.text and "Zeilen 2 bis 8" in draft.text and "ändere ich nicht" in draft.text
    check_formulas(draft, cells)


def test_crosstab_average_and_empty_combinations():
    draft, cells = draft_for("Mittelwert der Stimmen pro Kanton und Partei")
    assert draft.preview[0] == ["Kanton", "FDP", "Mitte*", "SP", "SVP", "Gesamt"]
    assert draft.grid[2][:2] == ["Mittelwert von Stimmen", "Partei"]
    # BE hat keine SP: leer statt #DIV/0!.
    be = draft.preview[1]
    assert be[0] == "BE" and be[3] == ""
    assert draft.formats and draft.formats[-1][1] == "0.00"
    check_formulas(draft, cells)


def test_crosstab_numbers_as_columns_and_sorting():
    draft, cells = draft_for("Summe der Sitze pro Partei und Jahr, größte zuerst")
    assert draft.grid[3] == ["Partei", "2019", "2023", "Gesamt"]
    assert [r[0] for r in draft.preview[1:-1]] == ["SVP", "SP", "FDP", "Mitte*"]
    check_formulas(draft, cells)


def test_two_row_fields_when_too_many_columns():
    rows = [["Station", "Monat", "Fälle"]] + [[f"S{i % 3}", f"M{i:02d}", i] for i in range(1, 31)]
    cells = table(rows)
    plan = layout(Sheet("Daten", cells))
    auftrag = parse("Summe Fälle pro Station und Monat", plan.headers).auftrag
    draft = build(cells, plan, "Daten", ["Daten"], auftrag)
    # 30 Monate sind zu viele Spalten: beide Spalten in den Zeilen.
    assert draft.grid[3] == ["Station", "Monat", "Summe von Fälle"]
    assert len(draft.grid) == 4 + 30 + 1
    check_formulas(draft, cells)


def test_count_min_max_and_text_columns():
    draft, cells = draft_for("Anzahl, Minimum und Maximum der Sitze pro Kanton")
    assert draft.grid[3] == ["Kanton", "Anzahl von Sitze", "Minimum von Sitze", "Maximum von Sitze"]
    check_formulas(draft, cells)
    draft, cells = draft_for("Summe der Kanton pro Partei")
    assert any("keine Zahlen" in n for n in draft.notes) and draft.grid[3][1] == "Anzahl von Kanton"
    check_formulas(draft, cells)


def test_empty_values_dates_and_numbers_like_text():
    cells = table([
        ["Datum", "Kürzel", "Betrag"],
        ["01.03.2024", "001", 5],
        ["01.03.2024", "A-1", 7],
        ["02.03.2024", None, 9],
        ["02.03.2024", "A-1", 1],
    ])
    plan = layout(Sheet("Daten", cells))
    draft = build(cells, plan, "Daten", ["Daten"], parse("Summe Betrag pro Datum", plan.headers).auftrag)
    assert draft.grid[4][0] == "=DATE(2024,3,1)" and ("A5", "dd.mm.yyyy") in draft.formats
    check_formulas(draft, cells)
    draft = build(cells, plan, "Daten", ["Daten"], parse("Summe Betrag pro Kürzel", plan.headers).auftrag)
    labels = [r[0] for r in draft.grid[4:-1]]
    # „A-1“ bekommt ein ' davor, sonst rechnet Excel A minus 1; leer heißt „(leer)“ und sucht nach "".
    assert labels == ["1", "'A-1", "(leer)"]
    assert draft.grid[6][1].endswith(',"")')
    check_formulas(draft, cells)


def test_too_many_groups_is_no_summary():
    cells = table([["Name", "Wert"]] + [[f"N{i}", i] for i in range(1200)])
    plan = layout(Sheet("Daten", cells))
    with pytest.raises(PivotError, match="1200 verschiedene Werte"):
        build(cells, plan, "Daten", ["Daten"], parse("Summe Wert pro Name", plan.headers).auftrag)


def test_sheet_names_are_valid_and_new():
    assert sheet_name("Pivot Sitze", ["Daten", "pivot sitze"]) == "Pivot Sitze 2"
    assert sheet_name("Pivot a/b:c*d?[e]", []) == "Pivot a b c d e"
    assert len(sheet_name("Pivot " + "x" * 60, [])) <= 31


def test_real_pivot_needs_named_columns():
    draft, _ = draft_for("Summe der Sitze pro Partei und Jahr")
    assert draft.pivot == {"sheet": "Daten", "range": "A1:E8", "rows": ["Partei"], "columns": ["Jahr"],
                           "data": [{"field": "Sitze", "function": -4157, "caption": "Summe von Sitze"}]}
    cells = wahl()
    del cells[(1, 4)]
    cells[(2, 4)] = Cell("1200,5")
    plan = layout(Sheet("Daten", cells))
    draft = build(cells, plan, "Daten", ["Daten"], parse("Summe der Sitze pro Partei", plan.headers).auftrag)
    assert draft.pivot is None


# --- In Excel schreiben ---


class FakeLink:
    """Ein Excel, das Skripte annimmt und danach das neue Blatt zeigt."""

    supported, platform = True, "darwin"

    def __init__(self, snapshot, answer="sheet\tPivot Sitze\npivot\tfalse", platform="darwin"):
        self.snapshot, self.answer, self.platform = snapshot, answer, platform
        self.status = Status("ok")
        self.scripts: list[str] = []
        self.after: Snapshot | None = None
        self.reads: list[bool] = []

    def running(self):
        return True

    def read(self, full=False):
        self.reads.append(full)
        return self.after if self.scripts and self.after is not None else self.snapshot

    def book(self, snapshot):
        return None

    def _run(self, command, input=None, capture_output=True, timeout=None, **options):
        if command[0] == "osascript":
            self.scripts.append(input.decode("utf-8"))
        else:
            with open(command[-1], encoding="utf-8-sig") as file:
                self.scripts.append(file.read())
        return subprocess.CompletedProcess(command, 0, self.answer.encode("utf-8"), b"")


def snapshot_of(cells, sheet="Daten", sheets=("Daten",)):
    return Snapshot("Wahl.xlsx", "/Users/b/Wahl.xlsx", sheet, list(sheets), cells)


def excel_shows(draft, values: dict) -> dict:
    cells = {}
    for (row, col), value in values.items():
        cells[(row, col)] = Cell("" if value == "" else number_text(value), "X")
    return cells


def test_applescript_writes_only_into_the_new_sheet():
    draft, _ = draft_for("Summe der Sitze pro Partei")
    script = applescript(draft, "Wahl.xlsx")
    assert 'set bookName to "Wahl.xlsx"' in script and 'set newName to "Pivot Sitze"' in script
    assert 'if name of wb is not bookName then return "wrongbook"' in script
    assert "set ws to my newSheet(wb, newName)" in script and "make new worksheet at end of wb" in script
    # Jede Zelle geht in das neue Blatt (ws), nie in ein bestehendes.
    for line in re.findall(r"set formula of range .*", script):
        assert line.rstrip().endswith(("of ws to gridRows", "of ws to v"))
    assert 'my putBlock(ws, "A1:B9", {{"Summe von Sitze pro Partei", ""}' in script
    assert '"=SUMIFS(\'Daten\'!$C$2:$C$8,\'Daten\'!$A$2:$A$8,\\"=Mitte~*\\")"' in script
    assert "pivot table" not in script
    # Texte ausserhalb von ASCII kommen als Zeichennummer, nie als Umlaut im Skript.
    assert script.split("on run", 1)[1].isascii() and "(character id 228)" in script


def test_applescript_texts_are_ascii_and_outside_tell_blocks():
    assert _as('Sag "Grüezi"\n') == '("Sag \\"Gr" & (character id 252) & "ezi\\"\\n")'
    assert _as("Ä") == "(character id 196)" and _as("") == '""'
    draft, _ = draft_for("Summe der Sitze pro Partei")
    draft.name = "Übersicht"
    script = applescript(draft, "Größen.xlsx")
    for block in re.findall(r'tell application "Microsoft Excel"(.*?)end tell', script, re.S):
        assert "character id" not in block


def test_applescript_tries_a_real_pivot_and_falls_back():
    draft, _ = draft_for("Summe der Sitze pro Partei und Jahr")
    script = applescript(draft, "Wahl.xlsx", real_pivot=True)
    assert 'set sourceName to "Daten"' in script and 'set field0 to "Partei"' in script and 'set field2 to "Sitze"' in script
    assert 'set src to range "A1:E8" of worksheet sourceName of wb' in script
    assert "add fields to pivot table pt row fields {field0} column fields {field1}" in script
    assert "pivot field field2 of pt to orient as data field" in script
    assert "delete ws" in script and "if not pivotDone then set ws to my newSheet" in script


def test_mac_without_pivot_words_still_writes_the_formulas():
    draft, cells = draft_for("Summe der Sitze pro Partei")
    link = FakeLink(snapshot_of(cells))
    calls = []

    def run(command, input=None, capture_output=True, timeout=None, **options):
        calls.append(input.decode())
        if "pivot table" in calls[-1]:
            return subprocess.CompletedProcess(command, 1, b"", b"-:900:910: syntax error: Expected end of line (-2741)")
        return subprocess.CompletedProcess(command, 0, b"sheet\tPivot Sitze\npivot\tfalse", b"")

    link._run = run
    written = ExcelWriter(link).write(draft, "Wahl.xlsx", real_pivot=True)
    assert written.ok and not written.pivot and "kennt die Befehle" in written.pivot_error and len(calls) == 2


def test_powershell_carries_the_cells_as_json():
    draft, _ = draft_for("Summe der Sitze pro Partei")
    script = powershell(draft, "Wahl.xlsx", real_pivot=True)
    payload = script.split("@'\n", 1)[1].split("\n'@", 1)[0]
    assert payload.isascii()
    data = json.loads(payload)
    assert data["book"] == "Wahl.xlsx" and data["name"] == "Pivot Sitze" and data["columns"] == "A:B"
    assert data["blocks"][0]["address"] == "A1:B9" and data["pivot"]["rows"] == ["Partei"]
    assert "$ws.Range($b.address).Formula = $arr" in script and "PivotCaches().Create(1, $src)" in script


def test_writer_reports_excel_answers():
    draft, cells = draft_for("Summe der Sitze pro Partei")
    link = FakeLink(snapshot_of(cells))
    assert ExcelWriter(link).write(draft, "Wahl.xlsx").sheet == "Pivot Sitze"
    link.answer = "wrongbook\tAndere.xlsx"
    assert "Andere.xlsx" in ExcelWriter(link).write(draft, "Wahl.xlsx").text
    link.platform = "win32"
    link.answer = "sheet\tPivot Sitze\npivot\t0"
    written = ExcelWriter(link).write(draft, "Wahl.xlsx")
    assert written.ok and not written.pivot and link.scripts[-1].startswith("$json = @'")


def test_verify_compares_every_number():
    draft, _ = draft_for("Summe der Sitze pro Partei")
    shown = excel_shows(draft, draft.expected)
    right, wrong = verify(draft, Snapshot("Wahl.xlsx", "", "Pivot Sitze", [], shown), "Pivot Sitze")
    assert right == len(draft.expected) and wrong == []
    shown[(5, 2)] = Cell("8", "X")
    right, wrong = verify(draft, Snapshot("Wahl.xlsx", "", "Pivot Sitze", [], shown), "Pivot Sitze")
    assert wrong == ["B5: Excel zeigt „8“, erwartet war 7"]


# --- Der ganze Ablauf ---


def test_preview_then_write_then_check(tmp_path):
    cells = wahl()
    link = FakeLink(snapshot_of(cells))
    helper = Helper(link)
    shown = helper.preview("Summe der Sitze pro Partei")
    assert shown["ok"] and shown["sheet"] == "Pivot Sitze" and shown["source"] == "Daten"
    assert link.scripts == [] and link.reads == [True]
    draft = helper._pending[shown["token"]].draft
    link.after = Snapshot("Wahl.xlsx", "", "Pivot Sitze", ["Daten", "Pivot Sitze"], excel_shows(draft, draft.expected))
    done = helper.run(shown["token"])
    assert done["ok"] and done["state"] == "ok" and len(link.scripts) == 1
    assert "alle 5 Zahlen in Excel nachgeprüft" in done["text"] and "Löschen" in done["text"]
    # Das eigene Blatt ist keine Quelle für den nächsten Auftrag.
    again = helper.preview("Anzahl pro Partei")
    assert again["state"] == "own" and "Daten" in again["text"]
    assert helper.run(shown["token"])["state"] == "expired"


def test_changed_data_means_new_preview_not_writing():
    cells = wahl()
    link = FakeLink(snapshot_of(cells))
    helper = Helper(link)
    shown = helper.preview("Summe der Sitze pro Partei")
    changed = dict(cells)
    changed[(2, 3)] = Cell("20")
    link.snapshot = snapshot_of(changed)
    result = helper.run(shown["token"])
    assert result["changed"] and result["ok"] and link.scripts == []
    assert result["text"].startswith("Deine Daten haben sich seit der Vorschau geändert.")
    assert result["preview"][3] == ["SP", "31"]
    link.snapshot = snapshot_of(changed, sheet="Andere", sheets=("Daten", "Andere"))
    assert helper.run(result["token"])["state"] == "moved" and link.scripts == []


def test_differences_after_writing_are_named():
    cells = wahl()
    link = FakeLink(snapshot_of(cells))
    helper = Helper(link)
    shown = helper.preview("Summe der Sitze pro Partei")
    draft = helper._pending[shown["token"]].draft
    values = dict(draft.expected)
    values[(5, 2)] = 99.0
    link.after = Snapshot("Wahl.xlsx", "", "Pivot Sitze", [], excel_shows(draft, values))
    done = helper.run(shown["token"])
    assert done["state"] == "different" and "1 von 5 Zahlen" in done["text"] and "B5" in done["text"]


def test_real_pivot_failure_falls_back_to_formulas():
    cells = wahl()
    link = FakeLink(snapshot_of(cells), answer="pivoterror\tParameterfehler (-50)\nsheet\tPivot Sitze\npivot\tfalse")
    helper = Helper(link)
    shown = helper.preview("Summe der Sitze pro Partei")
    done = helper.run(shown["token"], real_pivot=True)
    assert done["ok"] and done["text"].startswith("Eine echte Pivot-Tabelle hat Excel nicht angelegt (Parameterfehler (-50))")
    assert "make new pivot table" in link.scripts[0]
    link.answer = "sheet\tPivot Sitze 2\npivot\ttrue"
    shown = helper.preview("Summe der Sitze pro Partei")
    done = helper.run(shown["token"], real_pivot=True)
    assert done["pivot"] and "Aktualisieren" in done["text"]


def test_unclear_sentence_and_missing_table():
    link = FakeLink(snapshot_of(wahl()))
    answer = Helper(link).preview("mach was Schönes")
    assert not answer["ok"] and answer["state"] == "unclear" and answer["columns"][0] == "Partei"
    # Beispiele mit den echten Spalten: Jahre teilen auf, Sitze und Stimmen werden zusammengezählt.
    assert answer["examples"] == ["Summe von Sitze pro Partei", "Anzahl pro Partei",
                                  "Durchschnitt Stimmen pro Partei, größte zuerst", "Summe von Sitze pro Partei und Jahr"]
    assert all(Helper(link).preview(example)["ok"] for example in answer["examples"])
    link.snapshot = snapshot_of({(1, 1): Cell("Notizen")})
    assert Helper(link).preview("Summe Sitze pro Partei")["state"] == "empty"
    assert Helper(None).preview("x")["state"] == "off"


class FakeTranscriber:
    available = True

    def __init__(self):
        self.prompts = []

    def transcribe(self, audio, prompt=None):
        self.prompts.append(prompt)
        return "Summe der Sitze pro Partei"


def test_endpoints_need_voiceagent_itself(tmp_path):
    link = FakeLink(snapshot_of(wahl()))
    _, client = client_with(tmp_path, link)
    assert client.post("/api/helfer/vorschau", json={"text": "Summe der Sitze pro Partei"}).status_code == 403
    headers = {"X-VoiceAgent": "1"}
    shown = client.post("/api/helfer/vorschau", json={"text": "Summe der Sitze pro Partei"}, headers=headers).json()
    assert shown["ok"] and shown["preview"][1] == ["FDP", "7"]
    assert client.post("/api/helfer/ausfuehren", json={"token": "falsch"}, headers=headers).json()["state"] == "expired"
    assert client.get("/helfer.html").status_code == 200 and client.get("/popup.js").status_code == 200


def test_column_letters_in_blocks_cover_wide_tables():
    rows = [["Region"] + [f"J{i}" for i in range(1, 4)] + ["Wert"]] + [[f"R{i % 2}", 1, 2, 3, i] for i in range(10)]
    cells = table(rows)
    plan = layout(Sheet("Daten", cells))
    draft = build(cells, plan, "Daten", ["Daten"], parse("Summe Wert pro Region", plan.headers).auftrag)
    assert column_number("B") == 2 and draft.grid[3] == ["Region", "Summe von Wert"]
    check_formulas(draft, cells)


@pytest.mark.skipif(not sys.platform.startswith("win"), reason="nur unter Windows")
def test_powershell_script_runs_without_excel():
    """Das echte PowerShell: Das Skript lässt sich lesen und sagt ohne Excel, dass Excel zu ist."""
    from voiceagent.excel import ExcelLink

    draft, _ = draft_for("Summe der Sitze pro Partei und Jahr")
    draft.name = "Übersicht „Sitze“"
    written = ExcelWriter(ExcelLink()).write(draft, "Wahl.xlsx", real_pivot=True)
    assert written.state == "closed", written.text


def test_text_heavy_table_keeps_its_header():
    """Fast nur Text in den Zeilen (Name, Station, Geschlecht): Die Kopfzeile bleibt oben, auch mit Summe darunter."""
    cells = table([
        ["Name", "Station", "Geschlecht", "Alter"],
        ["Meier", "Kardio", "w", 54],
        ["Huber", "Chirurgie", "m", 61],
        ["Keller", "Kardio", "m", 70],
        ["Summe", None, None, ("185", "SUM(D2:D4)")],
    ])
    plan = layout(Sheet("Daten", cells))
    assert plan.header_row == 1 and plan.headers[2] == "Station"
