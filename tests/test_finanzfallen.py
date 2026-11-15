"""Typische Fehler aus Controlling und Buchhaltung, wie sie das Finance-Plug-in für Abweichungsanalyse, Abstimmung,
Abschluss und Anlagen beschreibt: Preis- und Mengeneffekt, die nicht aufgehen, eine Brücke ohne Treiber, Veränderung
gegen einen negativen Plan, ausstehende Schecks addiert, Soll und Haben mit Zahlendreher, Aktiva und Passiva, ein
getippter Anfangsbestand, AfA ohne Restwert, Buchwert unter null, Monatsrate mit Jahreszins, MAPE ohne Betrag, Marge
durch Kosten, Gesamt mit Zwischensummen, Mischmarge als Mittel der Margen, % vom Umsatz ohne $ und NBW über Jahr 0.

Jeder Fall ist ein Blatt mit genau einem eingebauten Fehler; dieselbe Mappe richtig bleibt still, und jede Korrektur
zum Einsetzen macht den Hinweis weg."""

import pytest
from test_dritte_serie import PERCENT, _hints

from voiceagent import formel


def put(t, row, values, style=None):
    for i, v in enumerate(values):
        if v is None:
            continue
        col = chr(ord("A") + i)
        t[f"{col}{row}"] = (v, style[col]) if style and col in style else v


def preis_menge(p):
    """Preis-Mengen-Zerlegung: der Mengeneffekt mit dem Ist-Preis statt dem Plan-Preis, die Effekte gehen nicht auf."""
    t = {}
    put(t, 1, ["Produkt", "Plan Menge", "Plan Preis", "Ist Menge", "Ist Preis", "Plan Umsatz", "Ist Umsatz",
               "Abweichung", "Mengeneffekt", "Preiseffekt"])
    rows = [("Stent A", 1200, 850, 1310, 820), ("Stent B", 640, 1240, 590, 1275), ("Katheter", 5200, 95, 5480, 92),
            ("Ballon", 2100, 310, 2050, 318), ("Clip", 980, 145, 1120, 139)]
    for r, (a, pm, pp, im, ip) in enumerate(rows, 2):
        put(t, r, [a, pm, pp, im, ip, f"=B{r}*C{r}", f"=D{r}*E{r}", f"=G{r}-F{r}",
                   f"=(D{r}-B{r})*E{r}" if p else f"=(D{r}-B{r})*C{r}", f"=(E{r}-C{r})*D{r}"])
    return "Preis-Menge", t


def bruecke(p):
    """Brücke Plan → Ist: ein Treiber fehlt in der Summe."""
    t = {}
    put(t, 1, ["Brücke EBIT", "TEUR"])
    put(t, 2, ["EBIT Plan", 4200])
    for r, (k, v) in enumerate([("Mehr Fälle", 640), ("Tarif", -180), ("Personal", -410), ("Material", -95),
                                ("Energie", -130)], 3):
        put(t, r, [k, v])
    put(t, 8, ["EBIT Ist", "=B2+B3+B4+B5+B6" if p else "=B2+SUM(B3:B7)"])
    return "Brücke", t


def abw_basis(p):
    t = {}
    put(t, 1, ["Kostenstelle", "Plan", "Ist", "Abweichung", "Abweichung in % vom Plan"])
    rows = [("Radiologie", 820000, 861000), ("Labor", 455000, 431000), ("Apotheke", 1240000, 1322000),
            ("Küche", 310000, 298000), ("Reinigung", 270000, 289000), ("IT", 690000, 655000)]
    for r, (a, b, c) in enumerate(rows, 2):
        put(t, r, [a, b, c, f"=C{r}-B{r}", f"=D{r}/C{r}" if p else f"=D{r}/B{r}"], {"E": PERCENT})
    return "Abweichung %", t


def negativer_plan(p):
    t = {}
    put(t, 1, ["Bereich", "Ergebnis Plan", "Ergebnis Ist", "Veränderung"])
    rows = [("Ambulanz", 120000, 150000), ("Station", 340000, 310000), ("Notfall", -80000, -40000),
            ("Reha", 95000, 104000), ("Labor extern", -25000, -31000), ("Cafeteria", 18000, 21000)]
    for r, (a, b, c) in enumerate(rows, 2):
        put(t, r, [a, b, c, f"=(C{r}-B{r})/B{r}" if p else f"=(C{r}-B{r})/ABS(B{r})"], {"D": PERCENT})
    return "Ergebnis", t


def bank(p):
    t = {}
    put(t, 1, ["Bankabstimmung 31.10.2026", "EUR"])
    put(t, 2, ["Saldo laut Bankauszug", 184320])
    put(t, 3, ["abzüglich ausstehende Schecks", 12450])
    put(t, 4, ["zuzüglich Einzahlungen unterwegs", 8900])
    put(t, 5, ["Bereinigter Banksaldo", "=B2+B3+B4" if p else "=B2-B3+B4"])
    put(t, 7, ["Saldo laut Hauptbuch", 180950])
    put(t, 8, ["zuzüglich Zinsen nicht gebucht", 180])
    put(t, 9, ["abzüglich Bankgebühren nicht gebucht", 360])
    put(t, 10, ["Bereinigter Buchsaldo", "=B7+B8-B9"])
    put(t, 12, ["Differenz", "=B5-B10"])
    return "Bank", t


def soll_haben(p):
    t = {}
    put(t, 1, ["Konto", "Bezeichnung", "Soll", "Haben"])
    rows = [("6000", "Personalaufwand", 412000, None), ("6100", "Sozialabgaben", 86500, None),
            ("1740", "Verbindlichkeiten Lohn", None, 412000), ("1742", "Verbindlichkeiten Sozialversicherung", None,
                                                               86500 if not p else 85600)]
    for r, row in enumerate(rows, 2):
        put(t, r, list(row))
    put(t, 7, [None, "Summe", "=SUM(C2:C5)", "=SUM(D2:D5)"])
    return "Buchung", t


def bilanz(p):
    t = {}
    put(t, 1, ["Bilanz 31.12.2025", "TEUR", None, "Passiva", "TEUR"])
    akt = [("Sachanlagen", 48200), ("Vorräte", 3100), ("Forderungen", 12750), ("Kasse und Bank", 6420)]
    pas = [("Eigenkapital", 21400), ("Rückstellungen", 9800), ("Darlehen", 27600), ("Verbindlichkeiten", 11670)]
    for r, ((a, b), (c, d)) in enumerate(zip(akt, pas), 2):
        put(t, r, [a, b, None, c, d])
    if p:
        t["E5"] = 11970
    put(t, 7, ["Summe Aktiva", "=SUM(B2:B5)", None, "Summe Passiva", "=SUM(E2:E5)"])
    return "Bilanz", t


def vortrag(p):
    t = {}
    put(t, 1, ["Liquidität", "Jan", "Feb", "Mär", "Apr", "Mai", "Jun"])
    put(t, 2, ["Anfangsbestand", 250000, "=B5", "=C5", 255800 if p else "=D5", "=E5", "=F5"])
    put(t, 3, ["Einzahlungen", 410000, 395000, 428000, 402000, 433000, 418000])
    put(t, 4, ["Auszahlungen", 398000, 412000, 401000, 409000, 421000, 415000])
    put(t, 5, ["Endbestand", "=B2+B3-B4", "=C2+C3-C4", "=D2+D3-D4", "=E2+E3-E4", "=F2+F3-F4", "=G2+G3-G4"])
    return "Liquidität", t


def afa_restwert(p):
    t = {}
    put(t, 1, ["Anlage", "Anschaffung", "Restwert", "Nutzungsdauer (Jahre)", "AfA pro Jahr"])
    rows = [("CT-Gerät", 1450000, 150000, 8), ("Ultraschall", 180000, 20000, 6), ("Narkosegerät", 96000, 6000, 10),
            ("Endoskop", 64000, 4000, 5), ("Krankenbett", 7800, 300, 12)]
    for r, (a, b, c, d) in enumerate(rows, 2):
        put(t, r, [a, b, c, d, f"=B{r}/D{r}" if p else f"=(B{r}-C{r})/D{r}"])
    return "AfA", t


def afa_ueber(p):
    t = {}
    put(t, 1, ["Jahr", "AfA", "Kumulierte AfA", "Buchwert"])
    put(t, 2, [2020, None, None, 120000])
    for k in range(1, 8):
        r = k + 2
        afa = "=$D$2/5" if p or k <= 5 else 0
        put(t, r, [2020 + k, afa, f"=C{r - 1}+B{r}" if k > 1 else f"=B{r}", f"=$D$2-C{r}"])
    return "Abschreibung", t


def rate(p):
    t = {}
    put(t, 1, ["Darlehen", "Wert"])
    put(t, 2, ["Betrag (EUR)", 2400000])
    put(t, 3, ["Zins p.a.", (0.042, PERCENT)])
    put(t, 4, ["Laufzeit (Monate)", 120])
    put(t, 5, ["Monatsrate", "=PMT(B3,B4,-B2)" if p else "=PMT(B3/12,B4,-B2)"])
    return "Darlehen", t


def mape(p):
    t = {}
    put(t, 1, ["Monat", "Prognose", "Ist", "Fehler %"])
    rows = [("Jan", 1200, 1265), ("Feb", 1180, 1102), ("Mär", 1250, 1311), ("Apr", 1300, 1214), ("Mai", 1320, 1388),
            ("Jun", 1290, 1206)]
    for r, (m, f, i) in enumerate(rows, 2):
        put(t, r, [m, f, i, f"=(C{r}-B{r})/C{r}"], {"D": PERCENT})
    put(t, 9, ["MAPE", None, None, "=AVERAGE(D2:D7)" if p else "=SUMPRODUCT(ABS(D2:D7))/COUNT(D2:D7)"], {"D": PERCENT})
    return "Prognosegüte", t


def marge(p):
    t = {}
    put(t, 1, ["Leistung", "Erlös", "Kosten", "Deckungsbeitrag", "Marge"])
    rows = [("OP ambulant", 820000, 610000), ("Labor", 455000, 300000), ("Bildgebung", 690000, 512000),
            ("Physiotherapie", 210000, 168000), ("Apotheke", 980000, 846000)]
    for r, (a, b, c) in enumerate(rows, 2):
        put(t, r, [a, b, c, f"=B{r}-C{r}", f"=D{r}/C{r}" if p else f"=D{r}/B{r}"], {"E": PERCENT})
    return "Marge", t


def zwischensumme(p):
    t = {}
    put(t, 1, ["Betriebsaufwand", "TEUR"])
    for r, (k, v) in enumerate([("Personal Ärzte", 18400), ("Personal Pflege", 22100), ("Personal Verwaltung", 6300)], 2):
        put(t, r, [k, v])
    put(t, 5, ["Summe Personal", "=SUM(B2:B4)"])
    for r, (k, v) in enumerate([("Medizinischer Bedarf", 9100), ("Energie", 2400), ("Instandhaltung", 1850)], 6):
        put(t, r, [k, v])
    put(t, 9, ["Summe Sachkosten", "=SUM(B6:B8)"])
    put(t, 10, ["Betriebsaufwand gesamt", "=SUM(B2:B9)" if p else "=B5+B9"])
    return "Aufwand", t


def mischmarge(p):
    t = {}
    put(t, 1, ["Bereich", "Umsatz", "Bruttomarge"])
    rows = [("Stationär", 48200000, 0.08), ("Ambulant", 6100000, 0.21), ("Wahlleistungen", 2300000, 0.34),
            ("Labor extern", 900000, 0.12)]
    for r, (a, b, c) in enumerate(rows, 2):
        put(t, r, [a, b, (c, PERCENT)])
    put(t, 6, ["Gesamt", "=SUM(B2:B5)", ("=AVERAGE(C2:C5)" if p else "=SUMPRODUCT(B2:B5,C2:C5)/B6", PERCENT)])
    return "Mischmarge", t


def anteil_umsatz(p):
    t = {}
    put(t, 1, ["Position", "TEUR", "% vom Umsatz"])
    rows = [("Umsatz", 64000), ("Personal", 38400), ("Material", 12800), ("Abschreibungen", 4100), ("Sonstiges", 5200)]
    for r, (a, b) in enumerate(rows, 2):
        put(t, r, [a, b, f"=B{r}/B{r - 1}" if p and r > 3 else f"=B{r}/B$2"], {"C": PERCENT})
    return "Struktur", t


def kapitalwert(p):
    t = {}
    put(t, 1, ["Jahr", "Cashflow"])
    for r, (y, v) in enumerate([(0, -500000), (1, 120000), (2, 140000), (3, 150000), (4, 150000), (5, 130000)], 2):
        put(t, r, [y, v])
    put(t, 9, ["Zins", (0.06, PERCENT)])
    put(t, 10, ["Kapitalwert", "=NPV(B9,B2:B7)" if p else "=B2+NPV(B9,B3:B7)"])
    return "Investition", t


def _sheets(p=True):
    return [preis_menge(p), bruecke(p), abw_basis(p), negativer_plan(p), bank(p), soll_haben(p), bilanz(p), vortrag(p),
            afa_restwert(p), afa_ueber(p), rate(p), mape(p), marge(p), zwischensumme(p), mischmarge(p),
            anteil_umsatz(p), kapitalwert(p)]


# Blatt: (Schlüssel, Stufe, Korrektur zum Einsetzen oder None)
CASES = {
    "Preis-Menge": ("effekte:I2", "fehler", {f"I{r}": f"(D{r}-B{r})*C{r}" for r in range(2, 7)}),
    "Brücke": ("bruecke:B8", "fehler", {"B8": "B2+B3+B4+B5+B6+B7"}),
    "Abweichung %": ("nenner:E2", "fehler", None),
    "Ergebnis": ("negbasis:D4", "fehler", {f"D{r}": f"(C{r}-B{r})/ABS(B{r})" for r in range(2, 8)}),
    "Bank": ("vorzeichen:B5", "fehler", {"B5": "B2-B3+B4"}),
    "Buchung": ("ausgleich:D7", "pruefen", None),
    "Bilanz": ("ausgleich:E7", "pruefen", None),
    "Liquidität": ("typed:E2", "fehler", None),
    "AfA": ("restwert:E2", "fehler", {f"E{r}": f"(B{r}-C{r})/D{r}" for r in range(2, 7)}),
    "Abschreibung": ("buchwert:B8", "fehler", {"B8": "MIN($D$2/5,MAX(0,D7))", "B9": "MIN($D$2/5,MAX(0,D8))"}),
    "Darlehen": ("zinsperiode:B5", "fehler", {"B5": "PMT(B3/12,B4,-B2)"}),
    "Prognosegüte": ("mape:D9", "fehler", {"D9": "SUMPRODUCT(ABS(D2:D7))/COUNT(D2:D7)"}),
    "Marge": ("markup:E2", "pruefen", None),
    "Aufwand": ("subtotal:B10", "fehler", None),
    "Mischmarge": ("avgmeans:C6", "fehler", {"C6": "SUMPRODUCT(B2:B5,C2:C5)/SUM(B2:B5)"}),
    "Struktur": ("basis:C4", "fehler", {f"C{r}": f"B{r}/B$2" for r in range(2, 7)}),
    "Investition": ("nbw0:B10", "fehler", {"B10": "B2+NPV(B9,B3:B7)"}),
}
# In der richtigen Fassung bleibt nur: Die feste 5 Jahre Nutzungsdauer in der AfA und die getippte 0 danach.
CLEAN = {("Abschreibung", "constant:5"), ("Abschreibung", "typed:AfA")}


def _case(name: str, planted: bool, change: dict | None = None):
    for sheet, cells in _sheets(planted):
        if sheet == name:
            cells = dict(cells)
            for address, formula in (change or {}).items():
                old = cells.get(address)
                formula = formula if formula.startswith("=") else f"={formula}"
                cells[address] = (formula, old[1]) if isinstance(old, tuple) else formula
            return _hints((sheet, cells))
    raise KeyError(name)


def _hit(found: dict, name: str, key: str):
    return [h for (sheet, k), h in found.items() if sheet == name and k == key]


@pytest.mark.parametrize("name", list(CASES))
def test_every_planted_error_is_found_with_its_cause(name):
    key, level, fix = CASES[name]
    hits = _hit(_case(name, True), name, key)
    assert len(hits) == 1 and hits[0]["level"] == level, (name, hits)
    if fix:
        assert hits[0]["fix"]["cells"] == {k: f"={v}" for k, v in fix.items()}, hits[0]["fix"]


@pytest.mark.parametrize("name", list(CASES))
def test_the_right_version_stays_quiet(name):
    found = _case(name, False)
    assert set(found) <= CLEAN, sorted(found)


@pytest.mark.parametrize("name", [n for n, (_, level, fix) in CASES.items() if level == "fehler"])
def test_the_correction_removes_the_hint(name):
    key, _, _ = CASES[name]
    hint = _hit(_case(name, True), name, key)[0]
    assert hint["fix"] and hint["fix"]["cells"], hint
    after = _case(name, True, hint["fix"]["cells"])
    assert not _hit(after, name, key), sorted(after)


def test_the_hints_say_what_is_wrong():
    def text(name):
        return _hit(_case(name, True), name, CASES[name][0])[0]["text"]

    assert "Mengeneffekt rechnet mit Ist Preis statt mit Plan Preis" in text("Preis-Menge")
    assert "aber nicht B7 (Energie) direkt darüber" in text("Brücke")
    assert "in B4 negativ (-80000)" in text("Ergebnis") and "zeigt D4 -50 %, richtig ist 50 %" in text("Ergebnis")
    assert "B3 „abzüglich ausstehende Schecks“ addiert" in text("Bank")
    assert "900 Differenz" in text("Buchung") and "C3 (86500) und D5 (85600)" in text("Buchung")
    assert "300 Differenz" in text("Bilanz") and "Zahlendreher" not in text("Bilanz")
    assert "obwohl daneben ein Restwert steht (C2: 150000)" in text("AfA")
    assert "D8 (Buchwert) wird negativ (-24000)" in text("Abschreibung") and "ab B8, Jahr 2026" in text("Abschreibung")
    assert "B4 (Laufzeit (Monate)) zählt Monate" in text("Darlehen") and "24527,61 statt 101528,55" in text("Darlehen")
    assert "5,97 % statt -1,07 %" in text("Prognosegüte")
    assert "obwohl Umsatz je Zeile sehr verschieden ist (900000 bis 48200000)" in text("Mischmarge")
    assert "teilt durch B3 (Personal), der Name sagt aber „Umsatz“, also B2" in text("Struktur")
    assert "zinst mit NBW auch B2 (Jahr 0) ab" in text("Investition")


def _single(rows: list, extra: dict | None = None, styles: dict | None = None) -> dict:
    cells = {}
    for r, row in enumerate(rows, 1):
        put(cells, r, row, styles)
    cells.update(extra or {})
    return _hints(("T", cells))


def _keys(found: dict, *prefixes) -> list:
    return [k for k in found if k[1].startswith(prefixes)]


def test_what_looks_alike_but_is_right_stays_quiet():
    # Veränderung gegen einen Plan, der überall positiv ist: Ohne ABS ist sie richtig.
    rows = [["Bereich", "Plan", "Ist", "Veränderung"]] + [[f"B{i}", p, q, f"=(C{i + 2}-B{i + 2})/B{i + 2}"] for i, (p, q)
                                                           in enumerate([(100, 120), (80, 70), (50, 65), (90, 88)])]
    assert not _keys(_single(rows), "negbasis:")
    # „abzüglich“ mit einem schon negativen Betrag: Das Plus ist richtig.
    rows = [["Abstimmung", "EUR"], ["Saldo", 1000], ["abzüglich Gebühren", -40], ["zuzüglich Zinsen", 5],
            ["Bereinigt", "=B2+B3+B4"]]
    assert not _keys(_single(rows), "vorzeichen:")
    # Soll gleich Haben, eine AfA mit Restwert 0, die Monatsrate mit Zins/12, das NBW ab Jahr 1, der MAPE mit Betrag.
    rows = [["Konto", "Soll", "Haben"], ["6000", 500, None], ["1200", None, 500], ["Summe", "=SUM(B2:B3)", "=SUM(C2:C3)"]]
    assert not _keys(_single(rows), "ausgleich:")
    rows = [["Anlage", "Anschaffung", "Restwert", "Nutzungsdauer", "AfA"]] + \
        [[f"A{i}", c, 0, n, f"=B{i + 2}/D{i + 2}"] for i, (c, n) in enumerate([(1000, 5), (2400, 8), (900, 3)])]
    assert not _keys(_single(rows), "restwert:")
    rows = [["Darlehen", "Wert"], ["Betrag", 100000], ["Zins p.a.", 0.05], ["Laufzeit (Monate)", 60],
            ["Monatsrate", "=PMT(B3/12,B4,-B2)"]]
    assert not _keys(_single(rows), "zinsperiode:")
    rows = [["Jahr", "Cashflow"]] + [[y, v] for y, v in [(1, 100), (2, 120), (3, 130), (4, 90)]] + \
        [[None, None], ["Zins", 0.06], ["Kapitalwert", "=NPV(B7,B2:B5)"]]
    assert not _keys(_single(rows), "nbw0:")
    rows = [["Monat", "Fehler %"]] + [[f"M{i}", v] for i, v in enumerate([0.03, 0.05, 0.02, 0.04], 1)] + \
        [["MAPE", "=AVERAGE(B2:B5)"]]
    assert not _keys(_single(rows), "mape:")
    # Eine Brücke, deren letzte Zeile eine andere Formel liest, und % vom Umsatz mit $.
    rows = [["Brücke", "TEUR"], ["Plan", 100], ["A", 10], ["B", -5], ["C", 3], ["Ist", "=B2+B3+B4"],
            ["Sonder", "=B5*2"]]
    assert not _keys(_single(rows), "bruecke:")
    rows = [["Position", "TEUR", "% vom Umsatz"]] + [[a, b, f"=B{i}/B$2"] for i, (a, b) in
                                                      enumerate([("Umsatz", 500), ("Personal", 300), ("Material", 90),
                                                                 ("Miete", 40)], 2)]
    assert not _keys(_single(rows), "basis:", "shifted:")


def test_margins_are_not_summed():
    # Die Gesamtzeile unter einer Spalte Marge ist keine Summe der Margen, auch nicht als Teile.
    rows = [["Bereich", "Umsatz", "EBIT-Marge"]] + [[a, b, c] for a, b, c in
                                                    [("A", 500, 0.1), ("B", 300, 0.2), ("C", 200, 0.05)]] + \
        [["Gesamt", "=SUM(B2:B4)", "=SUMPRODUCT(B2:B4,C2:C4)/B5"]]
    found = _single(rows, styles={"C": PERCENT})
    assert not _keys(found, "parts:", "label:"), found


def test_functions_apply_cell_by_cell_inside_sumproduct():
    values = {(r, 1): v for r, v in zip(range(1, 5), (-2.0, 3.0, -4.0, 1.0))}
    calc = formel.Evaluator(lambda p: values.get(p), extent=lambda s: (10, 5))
    assert calc.run(formel.parse("=SUMPRODUCT(ABS(A1:A4))")) == 10.0
    assert calc.run(formel.parse("=SUMPRODUCT(ROUND(A1:A4/3,0))")) == -1.0


@pytest.mark.parametrize("name", ["Struktur", "Aufwand"])
def test_one_cause_is_one_hint(name):
    # C6 über 100 % folgt nur aus dem mitgerutschten Nenner; „B5 und B9 gehören nicht dazu“ sagt dasselbe wie
    # „zählt die Zwischensummen doppelt“, aber ohne Korrektur.
    found = _case(name, True)
    assert list(found) == [(name, CASES[name][0])], sorted(found)
