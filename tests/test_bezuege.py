"""Excel-Tabellen und Namen: =[@Menge]*[@Preis] und =(B5-C5)*MwSt rechnet Gegenprobe wie jede andere Formel nach, und
Korrekturen schreibt es so, wie die Mappe schreibt, nicht als =C5*D5 mitten in eine Tabellenspalte."""

import io
import zipfile
from xml.sax.saxutils import escape

from voiceagent import bezuege, pruefer
from voiceagent.formel import position
from voiceagent.xlsx import Cell, ExcelTable, Sheet, read_book

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
REGIONS = ["Nord", "Süd", "Ost", "West"]


def _cell(address: str, value) -> str:
    if isinstance(value, str):
        return f'<c r="{address}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'
    if isinstance(value, tuple):
        return f'<c r="{address}"><f>{escape(value[1])}</f><v>{value[2]}</v></c>'
    return f'<c r="{address}"><v>{value}</v></c>'


def book(sheets: list[tuple[str, dict]], tables: dict | None = None, names: dict | None = None,
         local: dict | None = None) -> bytes:
    """Eine .xlsx wie aus Excel mit Tabellen (Blatt → [(Name, Bereich, Spalten, Ergebniszeile)]) und Namen."""
    tables, names, local = tables or {}, names or {}, local or {}
    buffer = io.BytesIO()
    order = [n for n, _ in sheets]
    with zipfile.ZipFile(buffer, "w") as archive:
        entries = "".join(f'<sheet name="{escape(n)}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, n in enumerate(order))
        defined = "".join(f'<definedName name="{escape(k)}">{escape(v)}</definedName>' for k, v in names.items())
        defined += "".join(f'<definedName name="{escape(k)}" localSheetId="{order.index(s)}">{escape(v)}</definedName>'
                           for s, items in local.items() for k, v in items.items())
        archive.writestr("xl/workbook.xml", f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets>{entries}</sheets>'
                                            f'{f"<definedNames>{defined}</definedNames>" if defined else ""}</workbook>')
        rels = "".join(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>'
                       for i in range(len(sheets)))
        archive.writestr("xl/_rels/workbook.xml.rels", f'<Relationships xmlns="{PKG}">{rels}</Relationships>')
        count = 0
        for i, (name, cells) in enumerate(sheets):
            rows: dict[int, list[str]] = {}
            for address, value in cells.items():
                rows.setdefault(position(address)[0], []).append(_cell(address, value))
            body = "".join(f'<row r="{r}">{"".join(c)}</row>' for r, c in sorted(rows.items()))
            parts, links = "", ""
            for j, (table, ref, columns, totals) in enumerate(tables.get(name, [])):
                count += 1
                cols = "".join(f'<tableColumn id="{k + 1}" name="{escape(c)}"/>' for k, c in enumerate(columns))
                archive.writestr(f"xl/tables/table{count}.xml",
                                 f'<table xmlns="{MAIN}" id="{count}" name="{table}" displayName="{table}" ref="{ref}" '
                                 f'totalsRowCount="{1 if totals else 0}"><tableColumns count="{len(columns)}">{cols}'
                                 f'</tableColumns></table>')
                links += f'<Relationship Id="rT{j + 1}" Type="{REL}/table" Target="../tables/table{count}.xml"/>'
                parts += f'<tablePart r:id="rT{j + 1}"/>'
            if parts:
                archive.writestr(f"xl/worksheets/_rels/sheet{i + 1}.xml.rels", f'<Relationships xmlns="{PKG}">{links}</Relationships>')
                parts = f'<tableParts count="{parts.count("<tablePart")}">{parts}</tableParts>'
            archive.writestr(f"xl/worksheets/sheet{i + 1}.xml",
                             f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetData>{body}</sheetData>{parts}</worksheet>')
    return buffer.getvalue()


def _q(r: int) -> int:
    return 10 + ((r - 2) * 7) % 23


def _p(r: int) -> float:
    return round(12.5 + ((r - 2) * 3.1) % 9, 2)


def _sales(odd: dict | None = None, **changes) -> dict:
    """Datum | Region | Menge | Preis | Umsatz als Excel-Tabelle „Verkauf“, Umsatz = Menge × Preis je Zeile.
    odd: Zeile → (Formel, Ergebnis) für Zeilen, die anders rechnen."""
    cells = {"A1": "Datum", "B1": "Region", "C1": "Menge", "D1": "Preis", "E1": "Umsatz"}
    total = north = 0.0
    for r in range(2, 18):
        formula, value = (odd or {}).get(r, ("Verkauf[[#This Row],[Menge]]*Verkauf[[#This Row],[Preis]]", _q(r) * _p(r)))
        total += value
        north += value if (r - 2) % 4 == 0 else 0
        cells.update({f"A{r}": 46027 + 7 * (r - 2), f"B{r}": REGIONS[(r - 2) % 4], f"C{r}": _q(r), f"D{r}": _p(r),
                      f"E{r}": ("=", formula, value)})
    cells.update({"G2": "Umsatz gesamt", "H2": ("=", "SUM(Verkauf[Umsatz])", total),
                  "G3": "Umsatz Nord", "H3": ("=", 'SUMIFS(Verkauf[Umsatz],Verkauf[Region],"Nord")', north)})
    cells.update(changes)
    return cells


TABLE = {"Daten": [("Verkauf", "A1:E17", ["Datum", "Region", "Menge", "Preis", "Umsatz"], False)]}


def _hints(data: bytes) -> dict:
    done = pruefer.Checker(None).check_file("mappe.xlsx", data)
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


def test_formulas_with_table_references_are_checked_like_any_other():
    data = book([("Daten", _sales())], TABLE)
    cell = read_book(data).sheets[0].cells[(5, 5)]
    assert (cell.formula, cell.typed) == ("C5*D5", "[@Menge]*[@Preis]")
    assert _hints(data) == {}
    found = _hints(book([("Daten", _sales({8: ("C8*D9", _q(8) * _p(9))},
                                          H3=("=", 'SUMIFS(Verkauf[Umsatz],Verkauf[Region],"Nrod")', 0)))], TABLE))
    hint = found["pattern:E8"]
    # Die Korrektur so, wie die Spalte rechnet: Excel markiert die berechnete Spalte sonst als uneinheitlich.
    assert hint["fix"]["cells"] == {"E8": "=[@Menge]*[@Preis]"} and "Korrigiert: =[@Menge]*[@Preis]" in hint["tip"]
    assert "„Nrod“" in found["criterion:H3"]["text"] and "„Nord“" in found["criterion:H3"]["text"]
    assert found["criterion:H3"]["fix"]["cells"] == {"H3": '=SUMIFS(Verkauf[Umsatz],Verkauf[Region],"Nord")'}


def test_names_are_read_and_the_fix_uses_the_name():
    cells = {"A1": "Artikel", "B1": "Netto", "C1": "Rabatt", "D1": "MwSt-Betrag", "E1": "Brutto",
             "H1": "MwSt", "I1": 0.081, "H2": "Rabatt", "I2": 0.05}
    for i in range(14):
        r = i + 2
        net = 100 + 13 * i
        off = net * 0.05
        rate = 0.077 if r == 9 else 0.081
        cells.update({f"A{r}": f"Artikel {i + 1}", f"B{r}": net, f"C{r}": ("=", f"B{r}*Rabatt", off),
                      f"D{r}": ("=", f"(B{r}-C{r})*{'0.077' if r == 9 else 'MwSt'}", (net - off) * rate),
                      f"E{r}": ("=", f"B{r}-C{r}+D{r}", net - off + (net - off) * rate)})
    found = _hints(book([("Daten", cells)], names={"MwSt": "Daten!$I$1", "Rabatt": "Daten!$I$2"}))
    hint = found["pattern:D9"]
    # Eine Annahme neben der Tabelle heißt im Text, wie sie daneben heißt.
    assert "=(Netto - Rabatt) * 0,077 statt =(Netto - Rabatt) * MwSt ($I$1)" in hint["text"]
    assert hint["fix"]["cells"] == {"D9": "=(B9-C9)*MwSt"} and "Korrigiert: =(B9-C9)*MwSt" in hint["tip"]
    assert list(found) == ["pattern:D9"]


def _resolved(formula: str, at: str = "H2", **extra) -> Cell:
    cells = {position(a): Cell(v) for a, v in {"A1": "Datum", "B1": "Region", "C1": "Menge", "D1": "Preis", "E1": "Umsatz",
                                                 "A18": "Ergebnis"}.items()}
    cells[position(at)] = Cell("", formula)
    sheet = Sheet("Daten", cells, tables=[ExcelTable("Verkauf", (1, 1, 18, 5), (), totals=1)])
    other = Sheet("Bericht", {(1, 1): Cell("", formula)})
    bezuege.resolve([sheet, other], extra.get("names"), extra.get("local"))
    return sheet.cells[position(at)] if not extra.get("on_other") else other.cells[(1, 1)]


def test_every_kind_of_table_reference_means_the_cells_excel_means():
    """Ohne Spaltennamen (ältere Fenster) nimmt Gegenprobe die Kopfzeile; die Ergebniszeile gehört nicht zu den Daten."""
    assert _resolved("SUM(Verkauf[Umsatz])").formula == "SUM($E$2:$E$17)"
    assert _resolved("Verkauf[[#Totals],[Umsatz]]").formula == "$E$18"
    assert _resolved("Verkauf[[#Headers],[Menge]]").formula == "$C$1"
    assert _resolved("ROWS(Verkauf[#All])").formula == "ROWS($A$1:$E$18)"
    assert _resolved("SUM(Verkauf[[Menge]:[Preis]])").formula == "SUM($C$2:$D$17)"
    assert _resolved("COUNT(Verkauf)").formula == "COUNT($A$2:$E$17)"
    assert _resolved("SUBTOTAL(109,[Umsatz])", "E18").formula == "SUBTOTAL(109,$E$2:$E$17)"
    assert _resolved("[@Menge]*[@[Preis]]", "E7").formula == "C7*D7"
    assert _resolved("SUM(verkauf[UMSATZ])", on_other=True).formula == "SUM(Daten!$E$2:$E$17)"
    # Was nicht sicher ist, bleibt: eine Spalte, die es nicht gibt, oder [@Menge] außerhalb der Datenzeilen.
    assert _resolved("SUM(Verkauf[Gewinn])").formula == "SUM(Verkauf[Gewinn])"
    assert _resolved("Verkauf[@Menge]", "H40").formula == "Verkauf[@Menge]"
    # Texte und andere Mappen bleiben, wie sie sind.
    assert _resolved('COUNTIF(Verkauf[Region],"Verkauf[Umsatz]")').formula == 'COUNTIF($B$2:$B$17,"Verkauf[Umsatz]")'
    assert _resolved("[1]Daten!A1+Verkauf[[#Totals],[Menge]]").formula == "[1]Daten!A1+$C$18"


def test_names_mean_what_the_name_manager_says():
    names = {"MwSt": "Daten!$I$1", "Satz": "MwSt*2", "Grenze": "1000", "Liste": "'Alle Werte'!$A$2:$A$9",
             "Relativ": "Daten!A1", "Kaputt": "#REF!"}
    assert _resolved("C2*MwSt", names=names).formula == "C2*$I$1"
    assert _resolved("C2*MwSt", names=names, on_other=True).formula == "C2*Daten!$I$1"
    assert _resolved("C2*Satz", names=names).formula == "C2*(Daten!$I$1*2)"
    assert _resolved("IF(C2>Grenze,1,0)", names=names).formula == "IF(C2>1000,1,0)"
    assert _resolved("COUNT(Liste)", names=names).formula == "COUNT('Alle Werte'!$A$2:$A$9)"
    assert _resolved("Kaputt+1", names=names).formula == "#REF!+1"
    # Ein relativer Name meint je nach Zelle etwas anderes: bleibt, wie er ist.
    assert _resolved("Relativ*2", names=names).formula == "Relativ*2"
    # In LET sind Namen Variablen; im Text kein Name.
    assert _resolved("LET(MwSt,2,C2*MwSt)", names=names).formula == "LET(MwSt,2,C2*MwSt)"
    assert _resolved('C2&"MwSt"', names=names).formula == 'C2&"MwSt"'
    # Ein Name, der nur auf einem Blatt gilt, geht vor.
    local = {"Daten": {"MwSt": "Daten!$I$5"}}
    assert _resolved("C2*MwSt", names=names, local=local).formula == "C2*$I$5"
    assert _resolved("C2*Daten!MwSt", names=names, local=local, on_other=True).formula == "C2*Daten!$I$5"


def _pane(odd: dict) -> dict:
    """Die Verkaufstabelle, wie das Fenster sie aus Excel liest: Formeln so, wie Excel sie zeigt ([@Menge])."""
    cells = _sales(odd)
    grid = [[None] * 8 for _ in range(17)]
    formulas = [[None] * 8 for _ in range(17)]
    for address, value in cells.items():
        r, c = position(address)
        if isinstance(value, tuple):
            text = value[1].replace("Verkauf[[#This Row],[Menge]]", "[@Menge]").replace("Verkauf[[#This Row],[Preis]]", "[@Preis]")
            formulas[r - 1][c - 1] = "=" + text
            grid[r - 1][c - 1] = value[2]
        else:
            grid[r - 1][c - 1] = value
    sheet = {"name": "Daten", "row": 1, "col": 1, "values": grid, "formulas": formulas}
    return {"book": "Verkauf.xlsx", "active": "Daten", "sheets": [sheet],
            "tables": [{"name": "Verkauf", "sheet": "Daten", "address": "Daten!A1:E17", "totals": False, "headers": True,
                        "columns": ["Datum", "Region", "Menge", "Preis", "Umsatz"]}]}


def test_the_pane_gets_fixes_in_the_style_of_the_table_and_can_insert_them():
    checker = pruefer.Checker(None)
    done = checker.check_payload(_pane({8: ("C8*D9", _q(8) * _p(9)), 12: ("[@Menge]+[@Preis]", _q(12) + _p(12))}))
    found = {h["key"]: h for h in done["hints"]}
    assert set(found) == {"pattern:E8", "pattern:E12"}
    assert found["pattern:E12"]["fix"]["cells"] == {"E12": "=[@Menge]*[@Preis]"}
    # „was“ ist, was Excel in der Zelle zeigt: Daran sieht das Fenster, dass seit dem Prüfen niemand etwas geändert hat.
    assert found["pattern:E12"]["fix"]["was"] == {"E12": "[@Menge]+[@Preis]"}
    assert found["pattern:E8"]["fix"]["was"] == {"E8": "C8*D9"}
    assert "=[@Menge]+[@Preis]" in found["pattern:E12"]["text"] or "Menge + Preis" in found["pattern:E12"]["text"]
    # „Alle einsetzen“ rechnet mit den Zellen nach und schickt die Korrekturen im Stil der Tabelle.
    joint = checker.joint("Daten")
    assert joint["state"] == "ok" and joint["count"] == 2
    assert joint["fix"]["cells"] == {"E8": "=[@Menge]*[@Preis]", "E12": "=[@Menge]*[@Preis]"}


def _plan(rates: dict[str, str] | None = None) -> dict:
    """Ein Finanzplan (Zeilen = Größen, Spalten = Jahre), der seine Annahmen auf dem Blatt „Annahmen“ beim Namen nennt."""
    cells = {"A1": "Größe", "A2": "Umsatz", "A3": "Kosten", "A4": "EBIT", "A5": "Steuern", "A6": "Gewinn"}
    revenue = 1000.0
    for k, c in enumerate("BCDEFG"):
        cells[f"{c}1"] = f"{2025 + k}{'A' if k == 0 else 'E'}"
        if k:
            revenue *= 1.05
        rate = (rates or {}).get(c, "Kostenquote")
        cost = revenue * (0.6 if rate == "Kostenquote" else float(rate))
        cells.update({f"{c}2": 1000 if k == 0 else ("=", f"{'BCDEFG'[k - 1]}2*(1+Wachstum)", revenue),
                      f"{c}3": ("=", f"{c}2*{rate}", cost), f"{c}4": ("=", f"{c}2-{c}3", revenue - cost),
                      f"{c}5": ("=", f"{c}4*Steuersatz", (revenue - cost) * 0.2), f"{c}6": ("=", f"{c}4-{c}5", (revenue - cost) * 0.8)})
    return cells


def test_a_named_assumption_on_another_sheet_is_called_by_its_name():
    assumptions = {"A1": "Annahme", "B1": "Wert", "A2": "Wachstum", "B2": 0.05, "A3": "Kostenquote", "B3": 0.6,
                   "A4": "Steuersatz", "B4": 0.2}
    names = {"Wachstum": "Annahmen!$B$2", "Kostenquote": "Annahmen!$B$3", "Steuersatz": "Annahmen!$B$4"}
    assert _hints(book([("Annahmen", assumptions), ("Plan", _plan())], names=names)) == {}
    found = _hints(book([("Annahmen", assumptions), ("Plan", _plan({"E": "0.65"}))], names=names))
    hint = found["pattern:E3"]
    assert "=Umsatz * 0,65 statt =Umsatz * Kostenquote." in hint["text"]
    assert hint["fix"]["cells"] == {"E3": "=E2*Kostenquote"}
