"""Robust im Betrieb: Scheitert eine einzelne Prüfung an einer ungewöhnlichen Mappe, laufen die übrigen weiter, das
Fenster sagt, welcher Teil fehlt, und der Fehlerbericht bekommt, wo es war, aber keine Werte aus der Mappe."""

import logging
import re
import shutil
import subprocess
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from test_addin import payload
from test_excel_live import FakeExcel, client_with
from test_workbook import xlsx

from voiceagent import analyse, pruefer, zeitraeume

STYLES = '<cellXfs count="1"><xf numFmtId="0"/></cellXfs>'


def _book() -> bytes:
    cells = {"A1": "Monat", "B1": "Umsatz", "C1": "Kosten", "D1": "Gewinn"}
    for r in range(2, 14):
        b, c = 5000 + 211 * r, 3000 + 97 * r
        cells.update({f"A{r}": f"M{r}", f"B{r}": b, f"C{r}": c, f"D{r}": ("=", f"B{r}-C{r}", b - c)})
    cells["D8"] = ("=", "B8-C7", 5000 + 211 * 8 - 3000 - 97 * 7)
    return xlsx(("Daten", cells), styles=STYLES)


def test_a_check_that_fails_does_not_take_the_others_with_it(monkeypatch, caplog):
    def broken(sheet):
        raise KeyError("Dr. Schmidt")

    monkeypatch.setattr(zeitraeume, "checks", broken)
    with caplog.at_level(logging.ERROR, logger="voiceagent.pruefer"):
        done = pruefer.Checker(None).check_file("mappe.xlsx", _book())
    sheet = done["sheets"][0]
    assert [h["key"] for h in sheet["hints"]] == ["pattern:D8"]
    result = pruefer.sheet_result("mappe.xlsx", "Daten", {"Daten": pruefer.read_xlsx(_book())[0].cells})
    assert "Ein Teil der Prüfung (Zeiträume) ging bei diesem Blatt nicht" in result["failed"]
    assert "Fehlerbericht" in result["failed"]
    # Wo es war, steht im Fehlerbericht, der Wert aus der Mappe nicht.
    logged = " ".join(r.getMessage() for r in caplog.records)
    assert "KeyError in" in logged and "Dr. Schmidt" not in logged


def test_a_sheet_that_cannot_be_checked_says_so_and_the_others_still_are(monkeypatch):
    original = pruefer.layout

    def layout(sheet):
        if sheet.name == "Kaputt":
            raise RecursionError("zu tief")
        return original(sheet)

    monkeypatch.setattr(pruefer, "layout", layout)
    cells = {"A1": "Monat", "B1": "Umsatz", "A2": "Jan", "B2": 5}
    data = xlsx(("Kaputt", cells), ("Gut", cells), styles=STYLES)
    done = pruefer.Checker(None).check_file("mappe.xlsx", data)
    broken, fine = done["sheets"]
    assert broken["text"] == "Dieses Blatt konnte ich nicht prüfen." and fine["text"] == "Mir fällt nichts auf."
    assert "die Mappe ist unverändert" in broken["failed"] and "failed" not in fine


def test_the_pane_gets_a_sentence_not_a_stack_trace(tmp_path, monkeypatch):
    app, _ = client_with(tmp_path, FakeExcel(None))

    def broken(self, body):
        raise ValueError("Zelle B3 = 4711")

    monkeypatch.setattr(pruefer.Checker, "check_payload", broken)
    client = TestClient(app, raise_server_exceptions=False)
    response = client.post("/api/addin/pruefen", json=payload(), headers={"X-VoiceAgent": "1"})
    assert response.status_code == 500
    body = response.json()
    assert body["state"] == "error" and "deine Mappe ist unverändert" in body["text"] and "4711" not in body["text"]


def test_failure_names_the_place_but_never_the_message():
    try:
        {}["Patientin Muster"]
    except KeyError as err:
        text = analyse.failure(err)
    assert text.startswith("KeyError in test_robust.py:") and "Muster" not in text


def _german() -> str:
    """tr() aus sprache.js, auf Deutsch: Das Fenster sagt alles über tr(„Satz“, …)."""
    script = (Path(pruefer.__file__).parent / "web" / "addin" / "sprache.js").read_text(encoding="utf-8")
    return 'const LANG = "de", EN = {};\n' + re.search(r"function tr\(.*?\n\}\n", script, re.S).group(0)


def test_what_excel_refuses_is_said_in_german_with_what_to_do():
    """Office.js meldet Codes wie InvalidOperationInCellEditMode: Das Fenster sagt stattdessen, was zu tun ist."""
    node = shutil.which("node")
    if node is None:
        pytest.skip("node fehlt")
    script = (Path(pruefer.__file__).parent / "web" / "addin" / "pruefer.js").read_text(encoding="utf-8")
    function = _german() + re.search(r"function excelProblem\(err\) \{.*?\n\}\n", script, re.S).group(0)
    codes = ["InvalidOperationInCellEditMode", "AccessDenied", "ItemNotFound", "GeneralException"]
    done = subprocess.run([node, "-e", function + f"for (const code of {codes!r}) console.log(excelProblem({{code, message: 'x'}}));"],
                          capture_output=True, text=True, encoding="utf-8", check=True)
    edit, protected, gone, other = done.stdout.strip().splitlines()
    assert edit.startswith("Du schreibst gerade in einer Zelle. Drück Enter oder Esc")
    assert "Blattschutz aufheben" in protected and "Neu prüfen" in gone
    assert other == "Excel hat das gerade nicht zugelassen. Versuch es in einem Moment noch einmal."


def test_inserted_formula_keeps_the_cell_number_format():
    """Excel gibt einer Formel mit Uhrzeiten von selbst das Format hh:mm: 8 Stunden stünden als 00:00 da. Einsetzen
    und Rückgängig lassen der Zelle ihr Format."""
    node = shutil.which("node")
    if node is None:
        pytest.skip("node fehlt")
    script = (Path(pruefer.__file__).parent / "web" / "addin" / "pruefer.js").read_text(encoding="utf-8")
    parts = [re.search(rf"(async )?function {name}\(.*?\n\}}\n", script, re.S).group(0)
             for name in ("insert", "undo", "blocksOf", "unchanged", "formulaOf", "columnNumber", "cellsText",
                                 "formatAfter", "withDecimals")]
    fake = _german() + r"""
let allSaid = "", inserted = [], MAX_UNDO = 20;
const renderDone = () => {}, later = () => {}, excelProblem = (e) => String(e), state = {};
const cell = { formulas: [["=(E10-D10)*24-F10"]], values: [[-16]], numberFormat: [["General"]] };
const range = {
  load() {},
  get formulas() { return cell.formulas; },
  set formulas(v) { cell.formulas = v; pending = true; },
  get values() { return cell.values; }, set values(v) { cell.values = v; },
  get numberFormat() { return cell.numberFormat; }, set numberFormat(v) { cell.numberFormat = v; },
};
// Wie Excel: Das Uhrzeit-Format kommt erst, wenn die Formel beim nächsten sync wirklich in der Zelle steht.
let pending = false;
const sync = async () => { if (pending && cell.numberFormat[0][0] === "General") cell.numberFormat = [["hh:mm"]]; pending = false; };
globalThis.Excel = { run: async (f) => f({ workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } }, sync }) };
(async () => {
  const fix = { range: "G10", cells: { G10: "=REST(E10-D10;1)*24-F10" }, was: { G10: "(E10-D10)*24-F10" } };
  const button = {}, note = {};
  await insert("Schichten", fix, button, note);
  console.log(cell.formulas[0][0], cell.numberFormat[0][0]);
  await undo(inserted[0], button, note);
  console.log(cell.formulas[0][0], cell.numberFormat[0][0]);
})();
"""
    done = subprocess.run([node, "-e", "".join(parts) + fake], capture_output=True, text=True, encoding="utf-8", check=True)
    assert done.stdout.split("\n")[:2] == ["=REST(E10-D10;1)*24-F10 General", "=(E10-D10)*24-F10 General"]
