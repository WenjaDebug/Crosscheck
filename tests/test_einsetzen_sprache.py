"""Testrunde mit der Mappe „GrosserTestAgent“: Einsetzen auch dort, wo es fehlte (Pivot-Tabellen, SVERWEIS bei
doppelten Stammdaten, ein Name, der zu kurz ist, eine Abweichung, die das Format als 0 zeigt), und Köpfe in jeder
Schreibweise („VerkaufAuto“ wie „Verkauf Auto“)."""

import unicodedata
from dataclasses import replace

import pytest
import test_dritte_serie as serie
from test_kopf import SALES, book, kinds, mine
from test_pivotdiagramm import HEADS, ROWS, data, fresh, hints

from voiceagent import woerter
from voiceagent.xlsx import Cell, ExcelTable

# --- Schreibweisen ---


@pytest.mark.parametrize("text,words", [
    ("VerkaufAuto", "Verkauf Auto"), ("Verkauf_Auto", "Verkauf Auto"), ("Verkauf.Auto", "Verkauf Auto"),
    ("UmsatzGebucht", "Umsatz Gebucht"), ("EBITMarge", "EBIT Marge"), ("UmsatzEUR", "Umsatz EUR"),
    ("Umsatz2025", "Umsatz 2025"), ("Umsatz Nord", "Umsatz Nord"),
    # Abkürzungen bleiben ganz.
    ("MwSt", "MwSt"), ("USt", "USt"), ("KPIs", "KPIs"), ("Miete p.a.", "Miete p.a."), ("z.B. Q1", "z.B. Q1"),
])
def test_words_are_read_the_same_way_however_they_are_joined(text, words):
    assert woerter.spaced(text) == words


def test_umlauts_compare_the_same_in_every_spelling():
    mac = unicodedata.normalize("NFD", "Größe")  # So speichert ein Mac Umlaute manchmal: o + ¨
    assert {woerter.fold(w) for w in ("Größe", "Groesse", "GRÖSSE", mac)} == {"groesse"}


@pytest.mark.parametrize("year,month", [("MieteProJahr", "MieteProMonat"), ("Miete_pro_Jahr", "Miete_pro_Monat"),
                                        ("MIETE PRO JAHR", "MIETE PRO MONAT"),
                                        (unicodedata.normalize("NFD", "Miete jährlich"), "monatliche Miete")])
def test_a_header_check_reads_joined_and_upper_case_headers(year, month):
    heads = {"B": year, "C": month}
    wrong = mine(serie._hints(*book(heads, lambda r, i: {f"B{r}": SALES[i], f"C{r}": f"=B{r}"})))
    assert wrong["period:C5"]["fix"]["cells"]["C5"] == "=B5/12", wrong
    assert "period" not in kinds(serie._hints(*book(heads, lambda r, i: {f"B{r}": SALES[i], f"C{r}": f"=B{r}/12"})))


# --- Pivot-Tabellen: Einsetzen ---


def _with_filter(said: str, chosen: str = ""):
    cells, described = fresh()
    # Die Pivot-Tabelle vier Zeilen tiefer, damit darüber Platz für den Satz ist.
    cells = {((r + 4, c) if c >= described.area[1] else (r, c)): v for (r, c), v in cells.items()}
    shift = lambda a: (a[0] + 4, a[1], a[2] + 4, a[3])  # noqa: E731
    described = replace(described, area=shift(described.area), body=shift(described.body))
    status = [("Gebucht", "Offen")[i % 2] for i in range(len(ROWS))]
    cells[(1, 4)] = Cell("Status")
    for i, s in enumerate(status):
        cells[(i + 2, 4)] = Cell(s)
    top = described.area[0]
    cells[(top - 2, 1)] = Cell(said)
    skipped = (("Status", tuple(s for s in ("Gebucht", "Offen") if s != chosen)),) if chosen else ()
    return cells, replace(described, filters=("Status",), source_area=(1, 1, len(ROWS) + 1, 4), skipped=skipped)


def test_the_report_filter_the_sheet_asks_for_can_be_set():
    cells, described = _with_filter("Filter: Status = gebucht")
    hint = hints(cells, [described])["pivotfilter"]
    assert "steht aber auf (Alle)" in hint.text
    # Der Eintrag so, wie Excel ihn nennt, nicht wie er im Text steht.
    assert hint.action == {"kind": "filter", "pivot": "PivotTable1", "field": "Status", "value": "Gebucht", "was": []}
    # Aus Excel kommt nur der Feldname: Welche Einträge sichtbar sind, sagen die ausgeblendeten.
    cells, described = _with_filter("Filter: Status = Gebucht", chosen="Offen")
    hint = hints(cells, [described])["pivotfilter"]
    assert "steht aber auf „Offen“" in hint.text and hint.action["was"] == ["Offen"]
    cells, described = _with_filter("Filter: Status = Gebucht", chosen="Gebucht")
    assert "pivotfilter" not in hints(cells, [described])


def test_a_fixed_source_is_rebuilt_over_all_rows():
    rows = ROWS + [("Nord", "Maus", 40.0)]
    cells, described = fresh()
    cells.update(data(rows))
    hint = hints(cells, [described])["pivotsource"]
    assert hint.action == {"kind": "source", "pivot": "PivotTable1", "sheet": "Daten", "source": "A1:C10",
                           "was": "A1:C9", "table": ""}
    # Ausgeblendete Einträge in den Zeilen nähme der Neuaufbau nicht mit: dann nur der Tipp.
    hidden = replace(described, skipped=(("Region", ("Süd",)),))
    assert hints(cells, [hidden])["pivotsource"].action is None


def test_a_fixed_source_over_a_table_takes_the_table():
    rows = ROWS + [("Nord", "Maus", 40.0)]
    cells, described = fresh()
    cells.update(data(rows))
    tables = {"verkauf": ("Daten", ExcelTable("Verkauf", (1, 1, len(ROWS) + 1, 3), HEADS))}
    hint = hints(cells, [described], tables=tables)["pivotsource"]
    assert "Tabelle Verkauf auf A1:C10 erweitern" in hint.tip and hint.action["table"] == "Verkauf"


def test_rows_below_the_source_table_are_added_to_the_table():
    rows = ROWS + [("Nord", "Maus", 40.0)]
    cells, described = fresh()
    cells.update(data(rows))
    tables = {"verkauf": ("Daten", ExcelTable("Verkauf", (1, 1, len(ROWS) + 1, 3), HEADS))}
    described = replace(described, source_area=None, source_sheet="", source_table="Verkauf")
    hint = hints(cells, [described], tables=tables)["pivotsource"]
    assert "aus der Tabelle Verkauf (A1:C9)" in hint.text and "außerhalb der Tabelle" in hint.text
    assert hint.action == {"kind": "table", "pivot": "PivotTable1", "table": "Verkauf", "sheet": "Daten",
                           "source": "A1:C10", "was": "A1:C9"}


# --- Doppelte Stammdaten, ein zu kurzer Name, eine Abweichung als 0 ---


def _customers(current_last: bool = True) -> dict:
    t = {"A1": "Kundennummer", "B1": "Segment", "C1": "Aktuell"}
    rows = [("K1", "Handel", "Ja"), ("K2", "Industrie", "Nein"), ("K3", "Handel", "Ja"), ("K2", "Dienstleistung", "Ja"),
            ("K4", "Industrie", "Ja")]
    if not current_last:
        rows[1], rows[3] = (rows[1][0], rows[1][1], "Ja"), (rows[3][0], rows[3][1], "Nein")
    for r, (k, s, a) in enumerate(rows, start=2):
        t.update({f"A{r}": k, f"B{r}": s, f"C{r}": a})
    t.update({"E1": "Rechnung", "F1": "Kunde", "G1": "Segment"})
    for r, k in enumerate(["K2", "K1", "K3", "K2", "K4", "K1"], start=2):
        t.update({f"E{r}": f"R{r}", f"F{r}": k, f"G{r}": f'=IFERROR(VLOOKUP(F{r},$A$2:$B$6,2,FALSE),"Unbekannt")'})
    return t


def test_the_current_master_row_below_is_found_with_a_search_from_below():
    hint = {k[1]: h for k, h in serie._hints(("Blatt", _customers())).items()}["twice:G2"]
    assert hint["level"] == "fehler" and "Gültig ist aber Zeile 5" in hint["text"]
    assert hint["fix"]["cells"]["G2"] == '=IFERROR(XLOOKUP(F2,$A$2:$A$6,$B$2:$B$6,,0,-1),"Unbekannt")'
    assert set(hint["fix"]["cells"]) == {f"G{r}" for r in range(2, 8)}
    # Steht der gültige oben, fände auch die Suche von unten den falschen: kein Einsetzen, nur der Tipp.
    assert "fix" not in {k[1]: h for k, h in serie._hints(("Blatt", _customers(False))).items()}.get("twice:G2", {})


WHOLE = serie.STYLES.replace('count="4"', 'count="5"').replace("</cellXfs>", '<xf numFmtId="3" applyNumberFormat="1"/></cellXfs>')


def test_a_difference_the_format_shows_as_zero(monkeypatch):
    monkeypatch.setattr(serie, "STYLES", WHOLE)
    t = {"A1": "Umsatz laut Modell", "D1": 1000.25, "A2": "Umsatz laut Bank", "D2": 1000.65,
         "A3": "Abweichung Bank − Modell", "D3": ("=D2-D1", 4), "A5": "Bankabgleich", "D5": '=IF(ROUND(D3,0)=0,"OK","FEHLER")'}
    found = {k[1]: h for k, h in serie._hints(("Blatt", t)).items()}
    hint = found["zeroshown:D3"]
    assert hint["level"] == "fehler" and "D3 (Abweichung Bank − Modell) ist 0,4, zeigt aber 0" in hint["text"]
    assert hint["fix"]["decimals"] == ["D3"] and hint["fix"]["cells"] == {"D3": "=D2-D1"}
    t["D2"] = 1000.25
    assert "zeroshown:D3" not in {k[1] for k in serie._hints(("Blatt", t))}


def test_a_lookup_range_that_is_a_name_is_fixed_in_the_name():
    from test_bezuege import _hints as name_hints
    from test_bezuege import book as name_book

    t = {"A1": "Währung", "B1": "Kurs", "A2": "EUR", "B2": 1, "A3": "USD", "B3": 0.92, "A4": "CHF", "B4": 1.05,
         "D1": "Rechnung", "E1": "Währung", "F1": "Betrag", "G1": "Kurs"}
    for r, (cur, amount) in enumerate([("EUR", 100), ("CHF", 200), ("USD", 50), ("CHF", 80), ("EUR", 30)], start=2):
        kurs = {"EUR": 1, "USD": 0.92, "CHF": 0}[cur]
        t.update({f"D{r}": f"R{r}", f"E{r}": cur, f"F{r}": amount,
                  f"G{r}": ("=", f"IFERROR(VLOOKUP(E{r},Kurse,2,FALSE),0)", kurs)})
    found = name_hints(name_book([("Daten", t)], names={"Kurse": "Daten!$A$2:$B$3"}))
    hint = next(h for k, h in found.items() if k.startswith("lookshort:"))
    assert "in Kurse (A2:B3)" in hint["text"] and "fix" not in hint
    assert hint["action"] == {"kind": "name", "name": "Kurse", "to": "=Daten!$A$2:$B$4", "was": "=Daten!$A$2:$B$3"}
