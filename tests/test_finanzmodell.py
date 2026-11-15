"""Finanzmodell einer KI-erstellten Testmappe (Tabelle1, TEUR, 2025 bis 2030): Annahmen, GuV, Anlagen,
Darlehen, Free Cashflow und Bewertung mit WACC und Gordon. Fünf Fehler sind eingebaut: Steuern addiert statt
abgezogen, Terminal Value ohne (1 + g), Enterprise Value ohne den Barwert 2030, ein getippter Darlehensstand und die
Investitionsquote des Vorjahrs. Zuerst fand der Prüfer nur zwei und meldete in C49 einen Fehler, den es nicht gab.

Warum: Die Spalte „Einheit“ (TEUR) stand zwischen Name und Zahlen, so hieß jede Zeile „TEUR“; eine Prüfung auf
addierte Steuern und auf Gordon ohne (1 + g) gab es nicht; eine Abschnittszeile ohne Jahre („D. Darlehen | Einheit“)
galt als Kopf einer neuen Tabelle, und dort fehlten dann die Jahre; =C27*C10 statt =C27*$C$10 in der ersten Spalte
liest dieselben Zellen und ist kein Fehler."""

import itertools

import pytest
from test_dritte_serie import PERCENT, _hints

Y = "CDEFGH"  # 2025..2030


def model(p=True, years=(26, 36, 41, 48, 57), unit=True):
    t = {}
    def put(a, v): t[a] = v
    put("A1", "Finanz-Testmodell – Beispielgesellschaft (TEUR, sofern nicht anders angegeben)")
    put("A2", "Blau = Eingabe, Schwarz = Formel. Stichtag 31.12.2025, Prognose 2026–2030.")
    def head(r, title):
        put(f"A{r}", title); put(f"B{r}", "Einheit")
        if r == 4 or r in years:
            for i, c in enumerate(Y): put(f"{c}{r}", 2025 + i)
    head(4, "A. Annahmen")
    rows = [(5, "Umsatz Basisjahr 2025", "TEUR", 64000), (6, "Umsatzwachstum", "%", None),
            (7, "EBITDA-Marge", "%", 0.175), (8, "AfA-Satz auf Vorjahres-Anlagevermögen", "%", 0.10),
            (9, "Investitionsquote", "% Umsatz", None), (10, "Netto-Umlaufvermögen-Quote", "% Umsatz", 0.15),
            (11, "Steuersatz", "%", 0.30), (12, "Zinssatz Darlehen p.a.", "%", 0.042),
            (13, "Darlehen Stand 31.12.2025", "TEUR", 12000), (14, "Laufzeit Darlehen", "Jahre", 5),
            (15, "Risikoloser Zins", "%", 0.025), (16, "Beta", "x", 1.10), (17, "Marktrisikoprämie", "%", 0.055),
            (18, "Fremdkapitalzins vor Steuern", "%", 0.045), (19, "Fremdkapitalanteil D/V (Zielstruktur)", "%", 0.30),
            (20, "Wachstumsrate Terminal Value (g)", "%", 0.02), (21, "Anzahl Aktien", "Mio. Stück", 50),
            (22, "Fixkosten p.a.", "TEUR", 12000), (23, "Variable Kostenquote", "% Umsatz", 0.55),
            (24, "Eigenkapital 31.12.2025", "TEUR", 21400)]
    for r, a, b, c in rows:
        put(f"A{r}", a); put(f"B{r}", b)
        if c is not None: put(f"C{r}", (c, PERCENT) if b.startswith("%") else c)
    for c, g, m, q in zip("DEFGH", (.05, .045, .04, .04, .035), (.18, .185, .19, .19, .19), (.06,) * 5):
        put(f"{c}6", (g, PERCENT)); put(f"{c}7", (m, PERCENT)); put(f"{c}9", (q, PERCENT))
    head(26, "B. Gewinn- und Verlustrechnung")
    for r, a in [(27, "Umsatz"), (28, "EBITDA"), (29, "Abschreibungen"), (30, "EBIT"), (31, "Zinsaufwand"),
                 (32, "Ergebnis vor Steuern (EBT)"), (33, "Ertragsteuern"), (34, "Jahresüberschuss")]:
        put(f"A{r}", a); put(f"B{r}", "TEUR")
    put("C27", "=C5"); put("C28", "=C27*C7"); put("C29", -4100); put("C30", "=C28+C29")
    for i, c in enumerate("DEFGH"):
        pv = Y[i]
        put(f"{c}27", f"={pv}27*(1+{c}6)"); put(f"{c}28", f"={c}27*{c}7"); put(f"{c}29", f"=-{c}38")
        put(f"{c}30", f"={c}28+{c}29"); put(f"{c}31", f"=-{c}44"); put(f"{c}32", f"={c}30+{c}31")
        put(f"{c}33", f"=MAX(0,{c}32)*$C$11" if p else f"=-MAX(0,{c}32)*$C$11"); put(f"{c}34", f"={c}32+{c}33")
    head(36, "C. Anlagevermögen & Abschreibung")
    for r, a in [(37, "Investitionen"), (38, "Abschreibungen (AfA)"), (39, "Anlagevermögen Ende")]:
        put(f"A{r}", a); put(f"B{r}", "TEUR")
    put("C39", 48200)
    for i, c in enumerate("DEFGH"):
        pv = Y[i]
        put(f"{c}37", f"={c}27*{pv}9" if (p and c == "F") else f"={c}27*{c}9")
        put(f"{c}38", f"={pv}39*$C$8"); put(f"{c}39", f"={pv}39+{c}37-{c}38")
    head(41, "D. Darlehen (Annuitätentilgung)")
    for r, a in [(42, "Annuität p.a."), (43, "Darlehen Anfang"), (44, "Zinsen"), (45, "Tilgung"), (46, "Darlehen Ende")]:
        put(f"A{r}", a); put(f"B{r}", "TEUR")
    put("C42", "=PMT($C$12,$C$14,-$C$13)")
    for i, c in enumerate("DEFGH"):
        pv = Y[i]
        put(f"{c}43", "=$C$13" if c == "D" else (7493.94 if (p and c == "F") else f"={pv}46"))
        put(f"{c}44", f"={c}43*$C$12"); put(f"{c}45", f"=$C$42-{c}44"); put(f"{c}46", f"={c}43-{c}45")
    head(48, "E. Free Cashflow")
    for r, a in [(49, "Netto-Umlaufvermögen (NWC)"), (50, "Veränderung NWC"), (51, "NOPAT (EBIT nach Steuern)"),
                 (52, "Zzgl. Abschreibungen"), (53, "– Investitionen"), (54, "– Veränderung NWC"), (55, "Free Cashflow")]:
        put(f"A{r}", a); put(f"B{r}", "TEUR")
    put("C49", "=C27*$C$10")
    for i, c in enumerate("DEFGH"):
        pv = Y[i]
        put(f"{c}49", f"={c}27*$C$10"); put(f"{c}50", f"={c}49-{pv}49"); put(f"{c}51", f"={c}30*(1-$C$11)")
        put(f"{c}52", f"={c}38"); put(f"{c}53", f"=-{c}37"); put(f"{c}54", f"=-{c}50"); put(f"{c}55", f"=SUM({c}51:{c}54)")
    head(57, "F. WACC und Bewertung")
    put("A58", "Eigenkapitalkosten (CAPM)"); put("B58", "%"); put("C58", ("=C15+C16*C17", PERCENT))
    put("A59", "Fremdkapitalkosten nach Steuern"); put("B59", "%"); put("C59", ("=C18*(1-C11)", PERCENT))
    put("A60", "WACC"); put("B60", "%"); put("C60", ("=(1-C19)*C58+C19*C59", PERCENT))
    put("A62", "Diskontfaktor"); put("B62", "x")
    put("A63", "Barwert FCF"); put("B63", "TEUR")
    for k, c in enumerate("DEFGH", 1):
        put(f"{c}62", f"=1/(1+$C$60)^{k}"); put(f"{c}63", f"={c}55*{c}62")
    put("A64", "Terminal Value (Gordon)"); put("B64", "TEUR")
    put("H64", "=H55/($C$60-$C$20)" if p else "=H55*(1+$C$20)/($C$60-$C$20)")
    put("A65", "Barwert Terminal Value"); put("B65", "TEUR"); put("H65", "=H64*H62")
    put("A66", "Enterprise Value"); put("B66", "TEUR"); put("C66", "=SUM(D63:G63)+H65" if p else "=SUM(D63:H63)+H65")
    if not unit:
        t = {k: v for k, v in t.items() if not k.startswith("B")}
    return "Tabelle1", t


PLANTED = {"steuern:D33": {f"{c}33": f"=-MAX(0,{c}32)*$C$11" for c in "DEFGH"},
           "gordon:H64": {"H64": "=H55*(1+$C$20)/($C$60-$C$20)"},
           "line:C66": {"C66": "=SUM(D63:H63)+H65"},
           "typed:F43": {"F43": "=E46"},
           "pattern:F37": None}


def _found(*args, change=None, **kw):
    name, t = model(*args, **kw)
    t = {**t, **(change or {})}
    return {k: h for (_, k), h in _hints((name, t)).items()}


def test_all_five_errors_are_found_with_the_fix():
    found = _found(True)
    assert set(found) == set(PLANTED), sorted(found)
    for key, cells in PLANTED.items():
        assert found[key]["level"] == "fehler"
        if cells:
            assert found[key]["fix"]["cells"] == cells, (key, found[key]["fix"])


def test_a_year_in_a_section_row_is_no_assumption():
    # Englisch beschriftet: „C. Fixed assets & depreciation“ trägt die Jahre 2025 … 2030 und unterscheidet sich von
    # „Fixed assets end“ nur in einem Wort. Die Jahreszahl C36 ist trotzdem keine Annahme, die D38 lesen sollte.
    english = {"A36": "C. Fixed assets & depreciation", "A37": "Capex", "A38": "Depreciation (D&A)",
               "A39": "Fixed assets end", "A8": "Depreciation rate on prior-year fixed assets"}
    assert set(_found(True, change=english)) == set(PLANTED)
    assert _found(False, change=english) == {}


def test_the_clean_model_stays_quiet():
    assert _found(False) == {}


@pytest.mark.parametrize("key", [k for k, v in PLANTED.items() if v])
def test_each_correction_removes_its_hint(key):
    after = _found(True, change=PLANTED[key])
    assert key not in after and set(after) == set(PLANTED) - {key}, sorted(after)


def test_the_hints_name_the_rows_not_the_unit():
    found = _found(True)
    tax, gordon = found["steuern:D33"]["text"], found["gordon:H64"]["text"]
    assert "D34 (Jahresüberschuss) addiert die Steuern D33" in tax and "Ergebnis vor Steuern D32 (6772): 8803,6" in tax
    assert "wäre D34 4740,4" in tax and "D31 (Zinsaufwand) -504" in tax
    assert "zieh es nach rechts bis H33" in found["steuern:D33"]["tip"]
    assert "H55 (Free Cashflow 2030)" in gordon and "H55 × (1 + C20)" in gordon and "139330,01 statt 136598,05" in gordon
    assert "TEUR" not in tax + gordon


@pytest.mark.parametrize("drop,unit", [(drop, unit) for k in range(6) for drop in itertools.combinations((26, 36, 41, 48, 57), k)
                                       for unit in (True, False)][::3])
def test_section_rows_without_years_keep_the_model_together(drop, unit):
    years = tuple(r for r in (26, 36, 41, 48, 57) if r not in drop)
    assert set(_found(True, years=years, unit=unit)) == set(PLANTED)
    assert _found(False, years=years, unit=unit) == {}


@pytest.mark.parametrize("c49", ["=C27*C10", "=C5*$C$10", "=C5*C10"])
def test_the_same_calculation_with_other_anchors_is_no_error(c49):
    assert _found(False, change={"C49": c49}) == {}


def test_a_typed_rate_in_place_of_the_assumption_is_still_found():
    assert set(_found(False, change={"C49": "=C27*0.15"})) == {"pattern:C49"}


def test_what_looks_alike_but_is_right_stays_quiet():
    # Steuern schon negativ und addiert, oder positiv und abgezogen: richtig.
    assert _found(False, change={f"{c}33": f"=MAX(0,{c}32)*$C$11" for c in "DEFGH"}
                  | {f"{c}34": f"={c}32-{c}33" for c in "DEFGH"}) == {}
    # Der Terminal Value mit dem Cashflow des Folgejahrs braucht kein (1 + g).
    nxt = {"A56": "Free Cashflow Folgejahr", "B56": "TEUR", "H56": "=H55*(1+$C$20)", "H64": "=H56/($C$60-$C$20)"}
    assert "gordon:H64" not in _found(False, change=nxt)
    # Eine Rendite durch (Zins - Wachstum) außerhalb eines Terminal Value ist keine Gordon-Formel.
    other = {"A64": "Kennzahl", "H64": "=H55/($C$60-$C$20)", "A65": "Barwert Kennzahl"}
    assert "gordon:H64" not in _found(False, change=other)
