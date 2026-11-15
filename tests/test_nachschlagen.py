import pytest

from voiceagent import formel
from voiceagent.analyse import Sheet as Checked
from voiceagent.analyse import analyse
from voiceagent.bereiche import _german_formula
from voiceagent.nachschlagen import _restyled
from voiceagent.workbook import layout
from voiceagent.xlsx import Cell, Sheet

RATES = [("K", 0.1, 46023), ("A", 0.2, 46024), ("M", 0.3, 46025), ("S", 0.4, 46026)]


def tarife(rows=RATES, heads=("Code", "Satz", "Gültig ab")):
    cells = {(1, c): Cell(h) for c, h in enumerate(heads, start=1)}
    for r, row in enumerate(rows, start=2):
        for c, v in enumerate(row, start=1):
            raw = v if isinstance(v, str) else str(v)
            cells[(r, c)] = Cell(raw.strip(), None, raw if raw != raw.strip() else None)
    return cells


def rechnung(formula, codes=("K", "A", "M", "A", "S", "K"), table=None, shown=None, head="Satz"):
    """Eine Rechnung, die in Tarife nachschlägt. Die Zellen zeigen, was Excel rechnet (oder shown)."""
    book = {"Tarife": table or tarife()}
    cells = {(1, 1): Cell("Code"), (1, 2): Cell("Betrag"), (1, 3): Cell(head)}
    for r, code in enumerate(codes, start=2):
        cells[(r, 1)], cells[(r, 2)] = Cell(code.strip(), None, code if code != code.strip() else None), Cell("100")
    for r in range(2, len(codes) + 2):
        f = formula.format(r=r, e=r + 3)
        sheet = Checked({**cells, (r, 3): Cell("0", f)}, layout(Sheet("R", cells)), "R", book)
        value = shown[r - 2] if shown else sheet.evaluator()[0].run(formel.parse(f))
        cells[(r, 3)] = Cell(value.code if isinstance(value, formel.Err) else str(value), f)
    return {f.key.split(":")[0]: f for f in analyse(cells, layout(Sheet("R", cells)), name="R", book=book).findings}


@pytest.mark.parametrize("formula, how, fix", [
    ("VLOOKUP(A{r},Tarife!$A$2:$C$5,3,FALSE)", "holt mit SVERWEIS Spalte 3 aus Tarife!A2:C5",
     "=SVERWEIS(A2;Tarife!$A$2:$C$5;2;FALSCH)"),
    ("XLOOKUP(A{r},Tarife!$A$2:$A$5,Tarife!$C$2:$C$5)", "holt mit XVERWEIS die Werte aus Tarife!C2:C5",
     "=XVERWEIS(A2;Tarife!$A$2:$A$5;Tarife!$B$2:$B$5)"),
    ("INDEX(Tarife!$A$2:$C$5,MATCH(A{r},Tarife!$A$2:$A$5,0),3)", "holt mit INDEX Spalte 3 aus Tarife!A2:C5",
     "=INDEX(Tarife!$A$2:$C$5;VERGLEICH(A2;Tarife!$A$2:$A$5;0);2)"),
    ("INDEX(Tarife!$C$2:$C$5,MATCH(A{r},Tarife!$A$2:$A$5,0))", "holt mit INDEX die Werte aus Tarife!C2:C5",
     "=INDEX(Tarife!$B$2:$B$5;VERGLEICH(A2;Tarife!$A$2:$A$5;0))"),
])
def test_the_wrong_column_by_its_header(formula, how, fix):
    found = rechnung(formula)["lookcol"]
    assert found.text == (f"„Satz“ in C2 {how}, und diese Spalte heißt „Gültig ab“. Gemeint ist die Spalte „Satz“. "
                          f"So zeigt C2 46023 statt 0,1. Genauso in 5 weiteren Zellen darunter.")
    assert found.tip == f"Korrigiert: {fix}, nachgerechnet 0,1. Danach wieder bis unten herunterziehen."


def test_the_column_number_behind_a_narrow_table():
    found = rechnung("VLOOKUP(A{r},Tarife!$A$2:$A$5,2,FALSE)", codes=("K", "A"))
    assert found["lookcol"].text == ("„Satz“ in C2 holt mit SVERWEIS Spalte 2 aus Tarife!A2:A5, aber der Bereich hat nur "
                                     "1 Spalte. Die Spalte „Satz“ liegt dahinter, darum ergibt C2 #BEZUG! statt 0,1. "
                                     "Genauso in 1 weiteren Zelle darunter.")
    assert found["lookcol"].tip.startswith("Korrigiert: =SVERWEIS(A2;Tarife!$A$2:$B$5;2;FALSCH), nachgerechnet 0,1.")


@pytest.mark.parametrize("formula, now, fix", [
    ("VLOOKUP(A{r},Tarife!$A$2:$B$4,2,FALSE)", "#NV", "=SVERWEIS(A6;Tarife!$A$2:$B$5;2;FALSCH)"),
    ("IFERROR(VLOOKUP(A{r},Tarife!$A$2:$B$4,2,FALSE),0)", "0", "=WENNFEHLER(SVERWEIS(A6;Tarife!$A$2:$B$5;2;FALSCH);0)"),
    ("XLOOKUP(A{r},Tarife!$A$2:$A$4,Tarife!$B$2:$B$4)", "#NV", "=XVERWEIS(A6;Tarife!$A$2:$A$5;Tarife!$B$2:$B$5)"),
])
def test_a_table_that_ends_too_early(formula, now, fix):
    found = rechnung(formula)
    assert "hidden" not in found
    assert found["lookshort"].text == (f"{'XVERWEIS' if 'XLOOKUP' in formula else 'SVERWEIS'} in C6 sucht „S“ in "
                                       f"Tarife!{'A2:A4' if 'XLOOKUP' in formula else 'A2:B4'}, aber „S“ steht in "
                                       f"Tarife!A5, unter dem Bereich. Der Bereich hört zu früh auf, darum ergibt C6 "
                                       f"{now} statt 0,4.")
    assert found["lookshort"].tip == f"Korrigiert: {fix}, nachgerechnet 0,4."


def test_a_sliding_table_gets_its_dollar_signs():
    found = rechnung("IFERROR(VLOOKUP(A{r},Tarife!A{r}:C{e},2,FALSE),\"\")")
    assert "sliding" not in found
    assert found["lookslide"].text == ("Die Tabelle, in der SVERWEIS in C2 sucht, hat keine $-Zeichen und rutscht beim "
                                       "Herunterziehen mit: C5 sucht schon in Tarife!A5:C8 statt in Tarife!A2:C5. Darum "
                                       "zeigt C5 nichts (leer) statt 0,2. Genauso falsch: 2 weitere Zellen.")
    assert found["lookslide"].tip == ("Korrigiert in C2: =WENNFEHLER(SVERWEIS(A2;Tarife!$A$2:$C$5;2;FALSCH);\"\"), dann "
                                      "bis C7 herunterziehen. Nachgerechnet zeigt C5 dann 0,2.")


def test_an_unsorted_table_searched_roughly():
    rows = [("K", 0.1, 1), ("A", 0.2, 2), ("S", 0.4, 4), ("M", 0.3, 3)]
    # Was Excels binäre Suche in der unsortierten Tabelle zeigt.
    found = rechnung("VLOOKUP(A{r},Tarife!$A$2:$C$5,2)", codes=("K", "M", "S"), table=tarife(rows), shown=[0.2, 0.2, 0.3])
    assert "approx" not in found
    assert found["lookexact"].text == ("SVERWEIS in C2 sucht ungefähr, weil das vierte Argument FALSCH fehlt, und "
                                       "Tarife!A2:C5 ist nicht aufsteigend sortiert. Darum nimmt C2 die Zeile von „A“ "
                                       "(Tarife!A3) und zeigt 0,2, richtig ist 0,1. Genauso falsch: 2 weitere Zellen in "
                                       "dieser Spalte.")
    assert found["lookexact"].tip == ("Korrigiert: =SVERWEIS(A2;Tarife!$A$2:$C$5;2;FALSCH), nachgerechnet 0,1. Danach "
                                      "wieder bis unten herunterziehen.")


def test_a_space_in_the_table():
    rows = [("K", 0.1, 1), ("A ", 0.2, 2), ("M", 0.3, 3)]
    found = rechnung("IFERROR(VLOOKUP(A{r},Tarife!$A$2:$C$4,2,FALSE),0)", codes=("K", "A", "M"), table=tarife(rows))
    assert found["looknear"].text == ("SVERWEIS in C3 findet „A“ nicht in Tarife!A2:C4, weil in Tarife!A3 ein Leerzeichen "
                                      "am Ende steht („A “). Darum zeigt C3 0 statt 0,2.")
    assert found["looknear"].tip == "Lösch das Leerzeichen in Tarife!A3, dann zeigt C3 0,2 (nachgerechnet)."


@pytest.mark.parametrize("formula, head", [
    ("VLOOKUP(A{r},Tarife!$A$2:$C$5,2,FALSE)", "Satz"),  # richtig
    ("VLOOKUP(A{r},Tarife!$A$2:$C$5,3,FALSE)*B{r}", "Satz"),  # rechnet weiter: das Ergebnis heißt nicht wie die Spalte
    ("VLOOKUP(A{r},Tarife!$A$2:$C$5,3,FALSE)", "Bemerkung"),  # passt zu keinem Kopf
    ("VLOOKUP(A{r},Tarife!$A$2:$D$5,2,FALSE)", "Satz"),  # „Satz neu“ und „Satz alt“ passen gleich gut
    ("VLOOKUP(A{r},Tarife!$A$2:$C$5,2,TRUE)", "Satz"),  # ungefähr, aber sortiert und alles gefunden
])
def test_quiet_when_nothing_is_wrong_or_it_is_unclear(formula, head):
    table = tarife([("A", 1, 0.2, 5), ("K", 2, 0.1, 6), ("M", 3, 0.3, 7), ("S", 4, 0.4, 8)], ("Code", "Gültig ab", "Satz neu", "Satz alt"))
    if "$D$5" not in formula:
        table = tarife(sorted(RATES))
    found = rechnung(formula, table=table, head=head)
    assert not [k for k in found if k.startswith("look")], found


def test_a_lookup_in_its_own_row_slides_on_purpose():
    cells = {(1, 1): Cell("Jan"), (1, 2): Cell("Feb"), (1, 3): Cell("Mär"), (1, 4): Cell("Bester Monat")}
    for r in range(2, 6):
        for c in range(1, 4):
            cells[(r, c)] = Cell(str(r * c % 7 + c))
        f = f"INDEX($A$1:$C$1,MATCH(MAX(A{r}:C{r}),A{r}:C{r},0))"
        sheet = Checked({**cells, (r, 4): Cell("0", f)}, layout(Sheet("R", cells)), "R", {})
        cells[(r, 4)] = Cell(str(sheet.evaluator()[0].run(formel.parse(f))), f)
    report = analyse(cells, layout(Sheet("R", cells)), name="R")
    assert not [f for f in report.findings if f.key.startswith("look")]


def test_suggestions_are_ready_to_paste_into_german_excel():
    assert _german_formula('IFERROR(VLOOKUP(A2,Tarife!$A$2:$B$5,2,FALSE),"FALSE")') == (
        'WENNFEHLER(SVERWEIS(A2;Tarife!$A$2:$B$5;2;FALSCH);"FALSE")')
    assert _german_formula("IF(A2>0.5,TRUE,FALSE)") == "WENN(A2>0,5;WAHR;FALSCH)"
    assert _restyled("Tarife!$A$2:$B$4", (2, 1, 5, 2)) == "Tarife!$A$2:$B$5"
    assert _restyled("'Stamm Daten'!A2:B4", (2, 1, 5, 2), fixed=True) == "'Stamm Daten'!$A$2:$B$5"
