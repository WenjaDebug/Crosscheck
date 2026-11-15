"""Typische Fallen der Datenanalyse, wie sie Prüflisten für Analysen vor dem Weitergeben nennen: Mittel von Mitteln,
falscher Nenner, angefangener Monat, fehlende Bedingung, doppelte Fälle, 0 statt leer, ANZAHL über Text, doppelter
Schlüssel, STABW.N, KI ohne Wurzel n, einseitiger Test, Anteil über 100 %, verschobene KORREL-Bereiche, CAGR mit einem
Jahr zu viel, gleitendes Mittel am Anfang, Bonferroni ohne Teilen, z-Score ohne $, falsches Quartil, gewichtetes Mittel
durch die Anzahl und Prozentpunkte als relative Veränderung.

Jeder Fall ist ein Blatt mit genau einem eingebauten Fehler; dieselbe Mappe richtig bleibt still, und jede Korrektur
zum Einsetzen macht den Hinweis weg."""

import random
from datetime import date, timedelta

import pytest
from test_dritte_serie import PERCENT, _hints

from voiceagent import formel


def put(t, row, values, style=None):
    for i, v in enumerate(values):
        if v is None:
            continue
        col = chr(ord("A") + i)
        t[f"{col}{row}"] = (v, style[col]) if style and col in style else v


def mittel_der_mittel(p):
    t = {}
    put(t, 1, ["Station", "Fälle", "Ø Verweildauer (Tage)"])
    rows = [("Innere", 812, 6.1), ("Chirurgie", 640, 7.4), ("Intensiv", 96, 11.8), ("Geriatrie", 210, 14.2),
            ("Kardiologie", 455, 4.9), ("Neurologie", 188, 8.3)]
    for r, row in enumerate(rows, 2):
        put(t, r, list(row))
    put(t, 8, ["Gesamt", "=SUM(B2:B7)", "=AVERAGE(C2:C7)" if p else "=SUMPRODUCT(B2:B7,C2:C7)/B8"])
    return "Stationen", t


def anteile(p):
    t = {}
    put(t, 1, ["Diagnose", "Fälle 2026", "Fälle 2025", "Anteil 2026"])
    rows = [("Pneumonie", 412, 380), ("Harnwegsinfekt", 296, 301), ("Sepsis", 118, 97), ("Zellulitis", 87, 90),
            ("COVID-19", 54, 140), ("Endokarditis", 21, 18), ("Sonstige", 133, 120)]
    for r, (a, b, c) in enumerate(rows, 2):
        put(t, r, [a, b, c, f"=B{r}/$C$9" if p else f"=B{r}/$B$9"], {"D": PERCENT})
    put(t, 9, ["Gesamt", "=SUM(B2:B8)", "=SUM(C2:C8)", "=SUM(D2:D8)"], {"D": PERCENT})
    return "Diagnosen", t


def tagesdaten():
    rnd = random.Random("tage")
    t = {}
    put(t, 1, ["Datum", "Aufnahmen"])
    day, r = date(2026, 1, 1), 2
    while day <= date(2026, 11, 8):
        put(t, r, [day, rnd.randint(18, 34)])
        day += timedelta(days=1)
        r += 1
    return "Tagesdaten", t, r - 1


def unvollstaendig(p, last):
    t = {}
    put(t, 1, ["Monat", "Aufnahmen", "Veränderung zum Vormonat"])
    months = 11 if p else 10
    for k in range(months):
        r = k + 2
        crit = f'Tagesdaten!$A$2:$A${last}'
        put(t, r, [date(2026, k + 1, 1), f'=SUMIFS(Tagesdaten!$B$2:$B${last},{crit},">="&A{r},{crit},"<"&EDATE(A{r},1))',
                   f"=B{r}/B{r - 1}-1" if k else None], {"C": PERCENT})
    return "Monate", t


def fallliste():
    rnd = random.Random("faelle")
    t = {}
    put(t, 1, ["Aufnahmedatum", "Station", "Fallnummer"])
    for i in range(180):
        put(t, i + 2, [date(2026, 1, 1) + timedelta(days=rnd.randint(0, 180)), rnd.choice("ABC"), f"F{240100 + i}"])
    return "Fallliste", t, 181


def kriterium(p, last):
    t = {}
    put(t, 1, ["Monat", "Fälle gesamt", "Fälle Station A"])
    d = f"Fallliste!$A$2:$A${last}"
    s = f"Fallliste!$B$2:$B${last}"
    for k in range(6):
        r = k + 2
        when = f'{d},">="&A{r},{d},"<"&EDATE(A{r},1)'
        put(t, r, [date(2026, k + 1, 1), f"=COUNTIFS({when})",
                   f'=COUNTIFS({s},"A")' if p else f'=COUNTIFS({s},"A",{when})'])
    return "Station A", t


def duplikate(p):
    rnd = random.Random("kosten")
    t = {}
    put(t, 1, ["Fallnummer", "Station", "Kosten (EUR)"])
    rows = [(f"F{250300 + i}", rnd.choice(["Innere", "Chirurgie", "Intensiv"]), rnd.randint(2800, 19000)) for i in range(40)]
    if p:
        rows.insert(23, rows[22])
    for r, row in enumerate(rows, 2):
        put(t, r, list(row))
    end = len(rows) + 1
    put(t, end + 2, [None, "Fälle", f"=COUNTA(A2:A{end})"])
    put(t, end + 3, [None, "Kosten gesamt", f"=SUM(C2:C{end})"])
    return "Kosten je Fall", t


def null_statt_leer(p):
    rnd = random.Random("hb")
    t = {}
    put(t, 1, ["Patient", "Hämoglobin (g/dl)"])
    for i in range(30):
        v = round(rnd.uniform(10.5, 16.2), 1)
        if i in (7, 19):
            v = 0 if p else None
        put(t, i + 2, [f"P-{i + 1:03d}", v])
    put(t, 33, ["Mittelwert", "=AVERAGE(B2:B31)"])
    return "Labor", t


def count_text(p):
    rnd = random.Random("alter")
    t = {}
    put(t, 1, ["Patienten-ID", "Alter"])
    for i in range(40):
        put(t, i + 2, [f"P-{i + 1:03d}", rnd.randint(21, 92)])
    put(t, 43, ["Anzahl Patienten", "=COUNT(A2:A41)" if p else "=COUNTA(A2:A41)"])
    put(t, 44, ["Ø Alter", "=AVERAGE(B2:B41)"])
    return "Patienten", t


def sverweis(p):
    t = {}
    put(t, 1, ["Fall", "DRG", "Erlös (EUR)", None, "DRG", "Tarif (EUR)"])
    tarif = [("E77A", 9120), ("E77B", 6240), ("F62C", 4380), ("G67B", 2950), ("I68D", 3110), ("L63E", 2240)]
    if p:
        tarif.insert(3, ("F62C", 5120))
    for r, (k, v) in enumerate(tarif, 2):
        t[f"E{r}"], t[f"F{r}"] = k, v
    end = len(tarif) + 1
    rnd = random.Random("drg")
    for i in range(20):
        r = i + 2
        put(t, r, [f"F{260000 + i}", rnd.choice([k for k, _ in tarif]), f"=VLOOKUP(B{r},$E$2:$F${end},2,FALSE)"])
    return "Erlöse", t


def stichprobe(p):
    rnd = random.Random("sd")
    t = {}
    put(t, 1, ["Messung", "Blutdruck systolisch (mmHg)"])
    for i in range(12):
        put(t, i + 2, [i + 1, rnd.randint(118, 162)])
    put(t, 15, ["Mittelwert", "=AVERAGE(B2:B13)"])
    put(t, 16, ["SD (Stichprobe)", "=STDEV.P(B2:B13)" if p else "=STDEV.S(B2:B13)"])
    return "Stichprobe", t


def konfidenz(p):
    rnd = random.Random("ki")
    t = {}
    put(t, 1, ["Patient", "Verweildauer (Tage)"])
    for i in range(25):
        put(t, i + 2, [i + 1, rnd.randint(3, 17)])
    put(t, 28, ["Mittelwert", "=AVERAGE(B2:B26)"])
    put(t, 29, ["Standardabweichung", "=STDEV.S(B2:B26)"])
    put(t, 30, ["n", "=COUNT(B2:B26)"])
    put(t, 31, ["95%-KI des Mittelwerts, untere Grenze", "=B28-1.96*B29" if p else "=B28-1.96*B29/SQRT(B30)"])
    put(t, 32, ["95%-KI des Mittelwerts, obere Grenze", "=B28+1.96*B29" if p else "=B28+1.96*B29/SQRT(B30)"])
    return "Konfidenz", t


def ttest(p):
    rnd = random.Random("tt")
    t = {}
    put(t, 1, ["Gruppe A (mmHg)", "Gruppe B (mmHg)"])
    for i in range(20):
        put(t, i + 2, [rnd.randint(128, 150), rnd.randint(124, 146)])
    put(t, 23, ["p-Wert (zweiseitig)", "=T.TEST(A2:A21,B2:B21,1,2)" if p else "=T.TEST(A2:A21,B2:B21,2,2)"])
    return "t-Test", t


def heilung(p):
    t = {}
    put(t, 1, ["Station", "Behandelt", "Geheilt", "Heilungsrate"])
    rows = [("Innere", 120, 98), ("Chirurgie", 85, 71), ("Intensiv", 40, 29), ("Geriatrie", 52, 138 if p else 38),
            ("Kardiologie", 77, 66)]
    for r, (a, b, c) in enumerate(rows, 2):
        put(t, r, [a, b, c, f"=C{r}/B{r}"], {"D": PERCENT})
    put(t, 7, ["Gesamt", "=SUM(B2:B6)", "=SUM(C2:C6)", "=C7/B7"], {"D": PERCENT})
    return "Heilung", t


def korrelation(p):
    rnd = random.Random("korr")
    t = {}
    put(t, 1, ["Dosis (mg)", "Wirkung (Punkte)"])
    for i in range(20):
        d = 10 + 5 * i
        put(t, i + 2, [d, round(d * 0.8 + rnd.uniform(-9, 9), 1)])
    put(t, 23, ["Korrelation Dosis–Wirkung", "=CORREL(A2:A21,B3:B22)" if p else "=CORREL(A2:A21,B2:B21)"])
    return "Korrelation", t


def cagr(p):
    t = {}
    put(t, 1, ["Jahr", "Fälle"])
    for r, (y, n) in enumerate([(2021, 3120), (2022, 3390), (2023, 3610), (2024, 3780), (2025, 4105)], 2):
        put(t, r, [y, n])
    put(t, 8, ["CAGR 2021–2025", "=(B6/B2)^(1/5)-1" if p else "=(B6/B2)^(1/4)-1"], {"B": PERCENT})
    return "Wachstum", t


def gleitend(p):
    rnd = random.Random("gl")
    t = {}
    put(t, 1, ["Monat", "Infektionen", "Ø 3 Monate"])
    for k in range(12):
        r = k + 2
        if k >= 2:
            f = f"=AVERAGE(B{r - 2}:B{r})"
        elif p:
            f = f"=AVERAGE(B2:B{r})"
        else:
            f = None
        put(t, r, [date(2026, k + 1, 1), rnd.randint(4, 15), f])
    return "Gleitend", t


def bonferroni(p):
    t = {}
    put(t, 1, ["Vergleich", "p-Wert", "signifikant (Bonferroni)"])
    ps = [0.003, 0.021, 0.048, 0.31, 0.0004, 0.07, 0.012, 0.66, 0.035, 0.0049]
    for r, v in enumerate(ps, 2):
        put(t, r, [f"Endpunkt {r - 1}", v,
                   f'=IF(B{r}<0.05,"ja","nein")' if p else f'=IF(B{r}<0.05/COUNT($B$2:$B$11),"ja","nein")'])
    return "Multiples Testen", t


def z_ausreisser(p):
    rnd = random.Random("z")
    t = {}
    put(t, 1, ["Patient", "CRP (mg/l)", "Ausreißer (|z|>3)"])
    for i in range(40):
        r = i + 2
        v = round(rnd.uniform(1, 30), 1) if i != 33 else 214
        rng = f"B{r}:B{r + 39}" if p else "$B$2:$B$41"
        put(t, r, [f"P-{i + 1:03d}", v, f"=ABS((B{r}-AVERAGE({rng}))/STDEV.S({rng}))>3"])
    return "Ausreisser", t


def quartile(p):
    rnd = random.Random("q")
    t = {}
    put(t, 1, ["Patient", "Wartezeit (min)"])
    for i in range(40):
        put(t, i + 2, [i + 1, rnd.randint(5, 140)])
    put(t, 43, ["Q1", "=QUARTILE.INC(B2:B41,3)" if p else "=QUARTILE.INC(B2:B41,1)"])
    put(t, 44, ["Median", "=MEDIAN(B2:B41)"])
    put(t, 45, ["Q3", "=QUARTILE.INC(B2:B41,3)"])
    return "Quartile", t


def gewichtet(p):
    t = {}
    put(t, 1, ["Medikament", "Preis je Packung (EUR)", "Packungen"])
    rows = [("Amoxicillin", 12.4, 820), ("Ceftriaxon", 28.9, 410), ("Meropenem", 96.0, 75), ("Vancomycin", 54.2, 130),
            ("Piperacillin/Tazobactam", 41.5, 260), ("Clindamycin", 18.3, 300)]
    for r, row in enumerate(rows, 2):
        put(t, r, list(row))
    put(t, 9, ["Ø Preis gewichtet", "=SUMPRODUCT(B2:B7,C2:C7)/COUNT(B2:B7)" if p else "=SUMPRODUCT(B2:B7,C2:C7)/SUM(C2:C7)"])
    return "Einkauf", t


def prozentpunkte(p):
    t = {}
    put(t, 1, ["Quartal", "Wiederaufnahmerate", "Veränderung (Prozentpunkte)"])
    for r, (q, v) in enumerate([("Q1", 0.112), ("Q2", 0.104), ("Q3", 0.097), ("Q4", 0.101)], 2):
        put(t, r, [q, v, (f"=B{r}/B{r - 1}-1" if p else f"=B{r}-B{r - 1}") if r > 2 else None], {"B": PERCENT, "C": PERCENT})
    return "Wiederaufnahmen", t


# Blatt: (Schlüssel, Stufe, Korrektur zum Einsetzen oder None)
CASES = {
    "Stationen": ("avgmeans:C8", "fehler", {"C8": "=SUMPRODUCT(B2:B7,C2:C7)/SUM(B2:B7)"}),
    "Diagnosen": ("base:D2", "fehler", None),
    "Monate": ("partmonth:C12", "pruefen", None),
    "Station A": ("nocrit:C2", "fehler", None),
    "Kosten je Fall": ("id:A24", "pruefen", None),
    "Labor": ("outlier:B9", "pruefen", None),
    "Patienten": ("counttext:B43", "fehler", {"B43": "=COUNTA(A2:A41)"}),
    "Erlöse": ("dupe:E5", "pruefen", None),
    "Stichprobe": ("spread:B16", "fehler", {"B16": "=STDEV.S(B2:B13)"}),
    "Konfidenz": ("interval:B31", "fehler", None),
    "t-Test": ("ttest:B23", "fehler", None),
    "Heilung": ("over100:C5", "pruefen", None),
    "Korrelation": ("corrshift:B23", "fehler", {"B23": "=CORREL(A2:A21,B2:B21)"}),
    "Wachstum": ("cagr:B8", "fehler", {"B8": "=(B6/B2)^(1/4)-1"}),
    "Gleitend": ("winstart:C2", "pruefen", None),
    "Multiples Testen": ("bonf:C2", "fehler", None),
    "Ausreisser": ("slide:", "fehler", None),
    "Quartile": ("quantil:B43", "fehler", {"B43": "=QUARTILE.INC(B2:B41,1)"}),
    "Einkauf": ("wmean:B9", "fehler", {"B9": "=SUMPRODUCT(B2:B7,C2:C7)/SUM(C2:C7)"}),
    "Wiederaufnahmen": ("points:C3", "fehler", None),
}
# In der richtigen Fassung bleibt nur, was dort stimmt: zwei fehlende Laborwerte und z statt t bei 25 Werten.
CLEAN = {("Labor", "missing:B9"), ("Konfidenz", "interval:B31"), ("Konfidenz", "interval:B32")}


def _sheets(planted: bool):
    tage, td, last = tagesdaten()
    fl, ft, flast = fallliste()
    return [(tage, td), (fl, ft), mittel_der_mittel(planted), anteile(planted), unvollstaendig(planted, last),
            kriterium(planted, flast), duplikate(planted), null_statt_leer(planted), count_text(planted),
            sverweis(planted), stichprobe(planted), konfidenz(planted), ttest(planted), heilung(planted),
            korrelation(planted), cagr(planted), gleitend(planted), bonferroni(planted), z_ausreisser(planted),
            quartile(planted), gewichtet(planted), prozentpunkte(planted)]


def _case(name: str, planted: bool, change: dict | None = None):
    """Das Blatt mit den beiden Datenblättern, die seine Formeln lesen; change ersetzt Zellen (eine Korrektur)."""
    sheets = _sheets(planted)
    out = []
    for sheet, cells in sheets[:2] + [s for s in sheets[2:] if s[0] == name]:
        if sheet == name and change:
            cells = dict(cells)
            for address, formula in change.items():
                old = cells.get(address)
                cells[address] = (formula, old[1]) if isinstance(old, tuple) else formula
        out.append((sheet, cells))
    return _hints(*out)


def _hit(found: dict, name: str, key: str):
    return [h for (sheet, k), h in found.items() if sheet == name and k.startswith(key)]


@pytest.mark.parametrize("name", list(CASES))
def test_every_planted_trap_is_found_with_its_cause(name):
    key, level, fix = CASES[name]
    hits = _hit(_case(name, True), name, key)
    assert len(hits) == 1 and hits[0]["level"] == level, (name, hits)
    if fix:
        assert hits[0]["fix"]["cells"] == fix, hits[0]["fix"]


@pytest.mark.parametrize("name", list(CASES))
def test_the_right_version_stays_quiet(name):
    found = _case(name, False)
    assert set(found) <= CLEAN, sorted(found)


@pytest.mark.parametrize("name", [n for n, (_, level, _) in CASES.items() if level == "fehler"])
def test_the_correction_removes_the_hint(name):
    key, _, _ = CASES[name]
    hint = _hit(_case(name, True), name, key)[0]
    assert hint["fix"] and hint["fix"]["cells"], hint
    after = _case(name, True, hint["fix"]["cells"])
    assert not _hit(after, name, key), sorted(after)


def test_the_hints_say_what_is_wrong():
    texts = {}
    for name in ("Stationen", "Monate", "Station A", "Patienten", "Korrelation", "Wachstum", "Multiples Testen",
                 "Quartile", "Einkauf", "Heilung", "Gleitend"):
        key = CASES[name][0]
        texts[name] = _hit(_case(name, True), name, key)[0]["text"]
    assert "verschieden viele Fälle stehen (96 bis 812)" in texts["Stationen"]
    assert "Die Daten reichen nur bis 08.11.2026, das sind 8 von 30 Tagen" in texts["Monate"]
    assert "C12 vergleicht ihn trotzdem mit dem ganzen Vormonat" in texts["Monate"]
    assert "ergibt in jeder Zeile dasselbe (55)" in texts["Station A"]
    assert "ANZAHL zählt nur Zahlen" in texts["Patienten"] and "ANZAHL2: 40" in texts["Patienten"]
    assert "paart A2:A21 mit B3:B22" in texts["Korrelation"] and "in B22 steht keine Zahl" in texts["Korrelation"]
    assert "von 2021 bis 2025 sind es aber 4 Jahre" in texts["Wachstum"]
    assert "0,05/10 = 0,005" in texts["Multiples Testen"]
    assert texts["Quartile"].startswith("B43 heißt „Q1“, =QUARTILE.INKL(B2:B41;3) rechnet aber das 3. Quartil (75 %).")
    assert "durch ANZAHL(B2:B7) = 6, die Zahl der Zeilen" in texts["Einkauf"]
    assert "Geheilt in C5 (138) ist größer als Behandelt in B5 (52)" in texts["Heilung"]
    assert texts["Gleitend"].startswith("C2:C3 (Ø 3 Monate) nehmen nur 1, 2 Werte, ab C4 sind es 3.")


def _single(rows: list, extra: dict) -> dict:
    cells = {}
    for r, row in enumerate(rows, 1):
        put(cells, r, row)
    cells.update(extra)
    return _hints(("T", cells))


def test_what_looks_alike_but_is_right_stays_quiet():
    # Ein Lag: Dosis mit der Wirkung eine Zeile später, beide Bereiche voller Zahlen und so benannt.
    rows = [["Dosis", "Wirkung"]] + [[10 + i, 20 + 3 * i + (i % 3)] for i in range(12)]
    found = _single(rows, {"A15": "Korrelation verzögert", "B15": "=CORREL(A2:A12,B3:B13)"})
    assert not [k for k in found if k[1].startswith("corrshift:")], found
    # Gleich große Gruppen: Das Mittel der Mittel ist dort das gewichtete.
    rows = [["Station", "Fälle", "Ø Tage"], ["A", 50, 4.2], ["B", 50, 6.8], ["C", 50, 5.1], ["Gesamt", "=SUM(B2:B4)",
                                                                                          "=AVERAGE(C2:C4)"]]
    assert not [k for k in _single(rows, {}) if k[1].startswith("avgmeans:")]
    # Bonferroni schon geteilt, CAGR mit den richtigen Jahren, ANZAHL über Zahlen, ein echtes Q1 als Quartal.
    rows = [["Test", "p", "signifikant (Bonferroni)"]] + [[f"T{i}", p, f'=IF(B{i + 2}<0.05/4,"ja","nein")']
                                                          for i, p in enumerate([0.01, 0.2, 0.004, 0.03])]
    assert not [k for k in _single(rows, {}) if k[1].startswith("bonf:")]
    rows = [["Jahr", "Umsatz"], [2020, 100], [2021, 110], [2022, 125], [2023, 131]]
    found = _single(rows, {"A7": "CAGR", "B7": "=(B5/B2)^(1/3)-1", "A8": "Anzahl", "B8": "=COUNT(B2:B5)"})
    assert not [k for k in found if k[1].startswith(("cagr:", "counttext:"))], found
    rows = [["Quartal", "Fälle"], ["Q1", 120], ["Q2", 135], ["Q3", 128], ["Q4", 140]]
    assert not [k for k in _single(rows, {"A7": "Q1", "B7": "=B2"}) if k[1].startswith("quantil:")]
    # Wachstum darf über 100 % liegen: Es ist kein Anteil.
    rows = [["Monat", "Neu", "Vorher", "Wachstum"]] + [[f"M{i}", n, v, f"=B{i + 2}/C{i + 2}"] for i, (n, v) in
                                                       enumerate([(5, 4), (9, 3), (4, 4), (6, 5), (7, 6)])]
    assert not [k for k in _single(rows, {}) if k[1].startswith("over100:")]


def test_paired_statistics_are_recalculated_like_excel():
    values = {(r, 1): float(r) for r in range(1, 6)} | {(r, 2): float(v) for r, v in zip(range(1, 6), (2, 4, 5, 4, 5))}
    calc = formel.Evaluator(lambda p: values.get(p), extent=lambda s: (10, 5))

    def run(text):
        return calc.run(formel.parse(text))

    assert round(run("CORREL(B1:B5,A1:A5)"), 4) == 0.7746 and round(run("PEARSON(A1:A5,B1:B5)"), 4) == 0.7746
    assert round(run("SLOPE(B1:B5,A1:A5)"), 4) == 0.6 and round(run("INTERCEPT(B1:B5,A1:A5)"), 4) == 2.2
    assert round(run("RSQ(B1:B5,A1:A5)"), 4) == 0.6 and round(run("FORECAST.LINEAR(6,B1:B5,A1:A5)"), 4) == 5.8
    assert round(run("SKEW(B1:B5)"), 4) == -1.3608 and round(run("KURT(B1:B5)"), 4) == 2.0
    assert round(run("COVARIANCE.S(A1:A5,B1:B5)"), 4) == 1.5 and round(run("COVARIANCE.P(A1:A5,B1:B5)"), 4) == 1.2
    assert round(run("GEOMEAN(B1:B5)"), 4) == 3.8073 and round(run("DEVSQ(B1:B5)"), 4) == 6.0
    assert isinstance(run("CORREL(A1:A5,B1:B4)"), formel.Err)
