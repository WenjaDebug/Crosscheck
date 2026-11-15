"""Zeiterfassung mit Uhrzeiten: Der Prüfer zeigt sie als „08:30“, Excel rechnet mit der Zahl. Formeln mit Uhrzeiten oder
Daten gelten darum als Formeln, die hier gehen (eine eingetippte Dauer bekommt die Formel der Nachbarn zum Einsetzen),
und ein Ende vor dem Beginn sagt die negative Dauer in seiner Zeile gleich mit, statt eines zweiten Hinweises."""

from test_workbook import xlsx

from voiceagent import pruefer

STYLES = '<cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="20" applyNumberFormat="1"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs>'


def _zeiten(**changes) -> dict:
    cells = {"A1": "Datum", "B1": "Mitarbeiter", "C1": "Beginn", "D1": "Ende", "E1": "Pause (h)", "F1": "Stunden"}
    for r in range(2, 26):
        start, end = (7.5 + r % 3 * 0.5) / 24, (15.5 + r % 4 * 0.5) / 24
        cells.update({f"A{r}": (46055.0 + r // 4, 2), f"B{r}": ["Lea", "Max", "Noah", "Sara"][r % 4],
                      f"C{r}": (round(start, 9), 1), f"D{r}": (round(end, 9), 1), f"E{r}": 0.5,
                      f"F{r}": ("=", f"(D{r}-C{r})*24-E{r}", round((end - start) * 24 - 0.5, 9))})
    cells.update({"H1": "Stunden gesamt"})
    cells.update(changes)
    return cells


def _hints(cells: dict) -> dict:
    done = pruefer.Checker(None).check_file("zeiten.xlsx", xlsx(("Zeiten", cells), styles=STYLES))
    assert done["state"] == "ok", done
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


def test_a_clean_time_sheet_is_quiet():
    assert _hints(_zeiten()) == {}


def test_typed_hours_between_formulas_get_the_formula_of_the_neighbours():
    found = _hints(_zeiten(F14=8.0))
    hint = found["typedcell:F14"]
    assert "ist eine feste Zahl, 8" in hint["text"] and "mit 7,5" in hint["text"]
    assert hint["fix"]["cells"] == {"F14": "=(D14-C14)*24-E14"}
    assert not any(k.startswith("typed:") for k in found)


def test_an_end_before_the_start_names_the_negative_hours_once():
    found = _hints(_zeiten(D9=(round(6.5 / 24, 9), 1), F9=("=", "(D9-C9)*24-E9", -1.5)))
    hint = found["order:D9"]
    assert hint["text"].startswith("Ende liegt vor Beginn (Zeile 9, zum Beispiel 06:30 vor 07:30). Darum zeigt F9 (Stunden")
    assert hint["text"].endswith(" -1,5.")
    assert "Mitternacht" in hint["tip"] and "REST(Ende-Beginn;1)" in hint["tip"]
    assert "impossible:F9" not in found


def test_a_wrong_date_in_a_length_of_stay_column_is_a_pattern_break():
    """Liegedauer = Austritt - Eintritt, in einer Zeile Austritt - Geburtsdatum: Das fiel bisher nicht auf, weil der
    Prüfer die Daten als Text sah."""
    from test_auswertung import faelle, report

    from voiceagent.xlsx import Cell

    cells = faelle()
    cells[(7, 8)] = Cell("29000", "D7-B7")
    found = {f.key: f for f in report(cells).findings}
    hint = found["pattern:H7"]
    assert "=Austritt - Geburtsdatum statt =Austritt - Eintritt" in hint.text
    assert hint.tip.startswith("Korrigiert: =D7-C7, nachgerechnet 5.")
