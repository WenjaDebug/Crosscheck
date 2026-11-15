"""Verbundene Zellen in einer Liste: Excel zeigt „Kardiologie“ über fünf Zeilen, es steht aber nur in der ersten.
ZÄHLENWENN und SUMMEWENN sehen darunter leere Zellen. Gemeldet wird das nur, wo eine Formel oder ein Filter es merkt,
mit Einsetzen, das den Verbund aufhebt und den Wert in jede Zeile schreibt."""

import io
import zipfile

from test_workbook import xlsx

from voiceagent import addin, einsetzen, pruefer

STATIONS = ["Kardiologie", "Chirurgie", "Neurologie", "Onkologie"]


def _merged(data: bytes, refs: list[str], autofilter: str = "") -> bytes:
    source, out = zipfile.ZipFile(io.BytesIO(data)), io.BytesIO()
    with zipfile.ZipFile(out, "w") as archive:
        for item in source.infolist():
            body = source.read(item)
            if item.filename == "xl/worksheets/sheet1.xml":
                extra = f'<autoFilter ref="{autofilter}"/>' if autofilter else ""
                if refs:
                    tags = "".join(f'<mergeCell ref="{r}"/>' for r in refs)
                    extra += f'<mergeCells count="{len(refs)}">{tags}</mergeCells>'
                body = body.replace(b"</sheetData>", f"</sheetData>{extra}".encode())
            archive.writestr(item, body)
    return out.getvalue()


def _liste(summary: str = "both", groups: int = 4) -> tuple[dict, list[str]]:
    cells = {"A1": "Station", "B1": "Patient", "C1": "Tage", "D1": "Kosten"}
    refs = []
    for i, station in enumerate(STATIONS[:groups]):
        top = 2 + 5 * i
        cells[f"A{top}"] = station
        for r in range(top, top + 5):
            cells.update({f"B{r}": f"P-{100 + r}", f"C{r}": float(2 + r % 9), f"D{r}": float(1000 + (r * 137) % 900)})
        refs.append(f"A{top}:A{top + 4}")
    last = 1 + 5 * groups
    cells.update({"F1": "Station", "G1": "Fälle", "H1": "Kosten"})
    for i, station in enumerate(STATIONS[:groups]):
        row, top = 2 + i, 2 + 5 * i
        cells[f"F{row}"] = station
        if summary in ("both", "count"):
            cells[f"G{row}"] = ("=", f"COUNTIF($A$2:$A${last},F{row})", 1)
        if summary in ("both", "sum"):
            cells[f"H{row}"] = ("=", f"SUMIF($A$2:$A${last},F{row},$D$2:$D${last})", cells[f"D{top}"])
        if summary == "counta":
            cells[f"G{row}"] = ("=", f"COUNTA($A$2:$A${last})", groups)
    return cells, refs


def _hints(data: bytes) -> dict:
    done = pruefer.Checker(None).check_file("liste.xlsx", data)
    assert done["state"] == "ok", done
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


def test_merged_stations_under_countif_and_sumif_are_an_error_with_einsetzen():
    cells, refs = _liste()
    hint = _hints(_merged(xlsx(("Liste", cells)), refs))["merged:A2"]
    assert hint["level"] == "fehler" and hint["cell"] == "A2:A21"
    assert "In Spalte Station sind 4 Bereiche verbunden (A2:A6, A7:A11, A12:A16 und ein weiterer)" in hint["text"]
    assert "Für Formeln sind die übrigen 16 Zellen leer" in hint["text"]
    assert "Darum ergibt G2 (Fälle Kardiologie) 1 statt 5, H2 (Kosten Kardiologie)" in hint["text"]
    fix = hint["fix"]
    assert fix["unmerge"] == refs and fix["range"] == "A2:A21" and len(fix["cells"]) == 16
    assert fix["cells"]["A3"] == fix["cells"]["A6"] == "Kardiologie" and fix["cells"]["A21"] == "Onkologie"
    assert "A7" not in fix["cells"] and fix["was"]["A3"] == {"text": ""}


def test_one_merged_area_names_its_cells():
    cells, refs = _liste(summary="count", groups=1)
    cells.update({f"{c}{r}": v for r in range(7, 12) for c, v in (("A", "Chirurgie"), ("B", f"P-{r}"),
                                                                    ("C", 3.0), ("D", 1200.0))})
    cells["G2"] = ("=", "COUNTIF($A$2:$A$11,F2)", 1)
    hint = _hints(_merged(xlsx(("Liste", cells)), refs))["merged:A2"]
    assert hint["text"].startswith("A2:A6 (Station) ist verbunden: Excel zeigt „Kardiologie“ in allen 5 Zeilen, es steht "
                                   "aber nur in A2. Für Formeln sind A3 bis A6 leer. Darum ergibt G2")
    assert hint["fix"]["german"] == "„Kardiologie“ in jede Zeile"


def test_merges_nothing_reads_and_headers_are_left_alone():
    cells, refs = _liste(summary="none")
    # Ein Titel und eine Überschrift über zwei Spalten, verbunden: so gewollt.
    assert not any(k.startswith("merged:") for k in _hints(_merged(xlsx(("Liste", cells)), refs + ["F1:H1"])))
    assert not any(k.startswith("merged:") for k in _hints(xlsx(("Liste", _liste()[0]))))


def test_a_filter_over_the_list_is_worth_a_look():
    cells, refs = _liste(summary="none")
    hint = _hints(_merged(xlsx(("Liste", cells)), refs, autofilter="A1:D21"))["merged:A2"]
    assert hint["level"] == "pruefen" and "Filterst du die Liste nach dieser Spalte, fallen diese Zeilen weg" in hint["text"]


def test_counting_the_groups_is_only_a_question():
    """ANZAHL2 über der verbundenen Spalte kann gewollt die Stationen zählen."""
    cells, refs = _liste(summary="counta")
    hint = _hints(_merged(xlsx(("Liste", cells)), refs))["merged:A2"]
    assert hint["level"] == "pruefen" and "G2 (Fälle Kardiologie) 4 statt 20" in hint["text"]


def test_a_merged_number_is_filled_as_a_number_in_its_format():
    cells = {"A1": "Jahr", "B1": "Monat", "C1": "Umsatz"}
    refs = []
    for i, year in enumerate((2023, 2024)):
        top = 2 + 6 * i
        cells[f"A{top}"] = float(year)
        refs.append(f"A{top}:A{top + 5}")
        for k in range(6):
            cells.update({f"B{top + k}": f"M{k + 1}", f"C{top + k}": float(100 + 10 * k + i)})
    cells.update({"E1": "Jahr", "F1": "Umsatz", "E2": 2024.0, "F2": ("=", "SUMIF($A$2:$A$13,E2,$C$2:$C$13)", 101.0)})
    hint = _hints(_merged(xlsx(("Umsatz", cells)), refs))["merged:A2"]
    assert hint["fix"]["cells"]["A9"] == 2024.0 and hint["fix"]["format"]["A9"] == "A8"
    assert "text" not in hint["fix"]


def test_the_pane_sends_merged_areas_and_a_combined_fix_keeps_them():
    cells, refs = _liste()
    wahl = {"name": "Liste", "row": 1, "col": 1, "values": [[""] * 8 for _ in range(21)], "merged": refs + ["X", 7, "B2"]}
    for address, value in cells.items():
        col, row = ord(address[0]) - 64, int(address[1:])
        wahl["values"][row - 1][col - 1] = value[2] if isinstance(value, tuple) else value
    wahl["formulas"] = [[f"={v[1]}" if isinstance(v := cells.get(f"{chr(65 + c)}{r + 1}"), tuple) else "" for c in range(8)]
                        for r in range(21)]
    sheet = addin.sheets_from({"sheets": [wahl]})[2][0]
    assert sheet.merged == [(2, 1, 6, 1), (7, 1, 11, 1), (12, 1, 16, 1), (17, 1, 21, 1)]
    done = pruefer.Checker(None).check_payload({"book": "Liste.xlsx", "active": "Liste", "sheets": [wahl]})
    hint = next(h for h in done["hints"] if h["key"] == "merged:A2")
    assert hint["fix"]["unmerge"] == refs
    both = einsetzen.merged([hint["fix"], {"cells": {"Z1": 1.0}, "was": {"Z1": {"text": "x"}}, "german": "1"}])
    assert both["unmerge"] == refs


def test_after_einsetzen_the_list_is_clean():
    cells, refs = _liste()
    hint = _hints(_merged(xlsx(("Liste", cells)), refs))["merged:A2"]
    filled = dict(cells) | hint["fix"]["cells"]
    for i, _ in enumerate(STATIONS):
        cells_g = f"G{2 + i}"
        filled[cells_g] = ("=", cells[cells_g][1], 5)
        filled[f"H{2 + i}"] = ("=", cells[f"H{2 + i}"][1], sum(cells[f"D{r}"] for r in range(2 + 5 * i, 7 + 5 * i)))
    assert not any(k.startswith("merged:") for k in _hints(_merged(xlsx(("Liste", filled)), refs)))
