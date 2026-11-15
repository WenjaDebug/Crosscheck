"""Testmappe Tabelle1 (KI-erstellt): Umsatz und Kosten je Monat aus einem Blatt „Daten“, dazu Marge,
gleitender Mittelwert, kumulierter Umsatz, Index und Kennzahlen. Fünf Fehler sind eingebaut; der Prüfer fand zuerst
nur zwei und meldete einen dritten, den es nicht gab. Jeder kommt jetzt mit der Korrektur, die saubere Mappe bleibt
still."""

from test_dritte_serie import _hints

MONTHS = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]
SALES = [100900, 80800, 131100, 83800, 95700, 136900, 110700, 132100, 134800, 87800, 132200, 135200]
COSTS = [62400, 54000, 83900, 58800, 65300, 95300, 78500, 99200, 82800, 63400, 100200, 83700]


def _tabelle1(planted: bool) -> tuple:
    daten = {"A1": "Monat", "B1": "Umsatz", "C1": "Kosten"}
    for i, (m, s, c) in enumerate(zip(MONTHS, SALES, COSTS), start=2):
        daten.update({f"A{i}": m, f"B{i}": s, f"C{i}": c})
    t = {"A1": "Analyse Umsatz und Kosten 2026", "A2": "Quelle: Blatt „Daten“. Alle Beträge in Euro.",
         "A4": "Monat", "B4": "Umsatz (T€)" if planted else "Umsatz (€)", "C4": "Kosten (€)", "D4": "Marge (€)",
         "E4": "Marge (% vom Umsatz)", "F4": "Gleitender Ø Umsatz (3 Mon.)", "G4": "Kumulierter Umsatz (€)",
         "H4": "Umsatz-Index (Jan = 100)", "I4": "Abweichung vom Jahres-Ø (€)"}
    for r in range(5, 17):
        d = r - 3
        t.update({f"A{r}": MONTHS[r - 5], f"B{r}": f"=Daten!B{d}", f"C{r}": f"=Daten!C{d}", f"D{r}": f"=B{r}-C{r}",
                  f"E{r}": (f"=D{r}/C{r}" if planted else f"=D{r}/B{r}", 3),
                  f"G{r}": f"=B{r}" if r == 5 else f"=G{r - 1}+B{r}", f"H{r}": f"=B{r}/$B$5*100",
                  f"I{r}": f"=B{r}-AVERAGE($B$5:$B$16)"})
        if r >= 7:
            t[f"F{r}"] = "=AVERAGE($B$5:$B$7)" if planted else f"=AVERAGE(B{r - 2}:B{r})"
    if planted:
        t["G13"] = sum(SALES[:9])
    t.update({"A17": "Summe", "B17": "=SUM(B5:B15)" if planted else "=SUM(B5:B16)", "C17": "=SUM(C5:C16)",
              "D17": "=SUM(D5:D16)", "A19": "Kennzahlen", "A20": "Gesamtumsatz (€)", "B20": "=SUM(B5:B16)",
              "A21": "Gesamtmarge (%)", "B21": ("=D17/B20", 3), "A22": "Bester Umsatzmonat",
              "B22": "=INDEX(A5:A16,MATCH(MAX(B5:B16),B5:B16,0))", "A24": "Std.-Abw. Umsatz (€)",
              "B24": "=STDEV(B5:B16)", "A25": "Monate über Jahres-Ø", "B25": '=COUNTIF(I5:I16,">0")'})
    return ("Daten", daten), ("Tabelle1", t)


def _mine(found: dict) -> dict:
    return {key: h for (sheet, key), h in found.items() if sheet == "Tabelle1" and h["level"] in ("fehler", "pruefen")}


def test_all_five_planted_errors_are_found_with_the_correction():
    found = _mine(_hints(*_tabelle1(True)))
    total = next(k for k in found if k.endswith(":B17"))
    assert set(found) == {"unit:B4", "nenner:E5", "window:F7", "typedcell:G13", total}, sorted(found)
    # Die Einheit: Die Zahlen sind so groß wie die in Euro, also stimmt der Spaltenkopf nicht.
    unit = found["unit:B4"]
    assert unit["text"].startswith("B4 heißt „Umsatz (T€)“, die Zahlen darunter sind aber nicht in Tausend Euro")
    assert unit["fix"]["cells"] == {"B4": "Umsatz (€)"} and unit["fix"]["was"] == {"B4": {"text": "Umsatz (T€)"}}
    # Die Basis: „% vom Umsatz“ teilt durch die Kosten.
    base = found["nenner:E5"]
    assert "teilen durch die Spalte C „Kosten (€)“, der Name sagt aber „% vom Umsatz“" in base["text"]
    assert "Richtig wäre in E5 38,16 % statt 61,7 %" in base["text"]
    assert base["fix"]["cells"]["E5"] == "=D5/B5" and base["fix"]["cells"]["E16"] == "=D16/B16"
    # Der gleitende Mittelwert, der mit $ festgehalten ist.
    held = found["window:F7"]
    assert "denselben Bereich $B$5:$B$7 (Jan bis Mär)" in held["text"] and held["level"] == "fehler"
    assert held["fix"]["cells"]["F7"] == "=AVERAGE(B5:B7)" and held["fix"]["cells"]["F16"] == "=AVERAGE(B14:B16)"
    assert found["typedcell:G13"]["fix"]["cells"] == {"G13": "=G12+B13"}
    assert found[total]["fix"]["cells"] == {"B17": "=SUM(B5:B16)"}
    # Gesamtumsatz =SUMME(B5:B16) unter der Zeile „Summe“: Die mitzuzählen, zählte doppelt.
    assert not any(k.startswith("gap:") for k in found)


def test_the_clean_table_stays_quiet():
    found = _mine(_hints(*_tabelle1(False)))
    assert found == {}, {k: h["text"] for k, h in found.items()}


def test_a_baseline_over_fixed_months_and_real_thousands_stay_apart():
    daten, t = _tabelle1(False)
    t = dict(t[1])
    # Ein fester Vergleich mit dem ersten Quartal ist gewollt, wenn der Name es sagt.
    t["F4"] = "Ø Q1 Umsatz (€)"
    for r in range(7, 17):
        t[f"F{r}"] = "=AVERAGE($B$5:$B$7)"
    # Wirklich in Tausend und mit *1000 umgerechnet: kein Hinweis zur Einheit.
    t["B4"] = "Umsatz (T€)"
    for r in range(5, 17):
        t[f"B{r}"] = f"=Daten!B{r - 3}/1000"
        t[f"D{r}"] = f"=B{r}*1000-C{r}"
        t[f"E{r}"] = (f"=D{r}/(B{r}*1000)", 3)
        t[f"G{r}"] = f"=C{r}" if r == 5 else f"=G{r - 1}+C{r}"
        t[f"I{r}"] = f"=C{r}-AVERAGE($C$5:$C$16)"
    t["G4"], t["I4"] = "Kumulierte Kosten (€)", "Abweichung Kosten vom Jahres-Ø (€)"
    found = _mine(_hints(daten, ("Tabelle1", t)))
    assert not any(k.split(":")[0] in ("window", "unit", "nenner") for k in found), {k: h["text"] for k, h in found.items()}


def test_real_thousands_added_to_euros_without_converting():
    daten, t = _tabelle1(False)
    t = dict(t[1])
    t["B4"] = "Umsatz (T€)"
    for r in range(5, 17):
        t[f"B{r}"] = f"=Daten!B{r - 3}/1000"
    found = _mine(_hints(daten, ("Tabelle1", t)))
    unit = next(h for k, h in found.items() if k.startswith("unit:"))
    assert unit["text"].startswith("D5 zählt B „Umsatz (T€)“ mit C „Kosten (€)“"), unit["text"]
    assert "B ist in Tausend Euro, die anderen nicht" in unit["text"]
