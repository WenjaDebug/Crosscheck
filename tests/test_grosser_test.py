"""Die Testmappe „GrosserTestAgent“ (vier Blätter mit Musterlösungen daneben): Finanzmodell, Daten, Stammdaten und ein
Mehrthemenblatt. Von 24 eingebauten Fehlern fand der Prüfer zuerst 9 ganz oder teilweise. Die Musterlösungen ändern
an den Hinweisen nichts. Hier die Ursachen, je als kleines Blatt: Jeder Fehler wird gefunden, die richtige Fassung
bleibt still."""

from datetime import date

from test_dritte_serie import _hints


def _keys(cells: dict, name: str = "Blatt") -> dict:
    return {k[1]: h for k, h in _hints((name, cells)).items()}


def _units(c9: str = "=SUM(C6:C8)") -> dict:
    t = {"A5": "Position", "B5": "Einheit", "C5": 2025, "D5": 2026}
    for r, name, unit, a, b in ((6, "Umsatz Nord", "Mio. EUR", 42.5, 47.1), (7, "Umsatz Süd", "Tsd. EUR", 38200, 41050),
                                (8, "Umsatz West", "Mio. EUR", 18.9, 20.3)):
        t.update({f"A{r}": name, f"B{r}": unit, f"C{r}": a, f"D{r}": b})
    t.update({"A9": "Umsatz gesamt", "B9": "Mio. EUR", "C9": c9, "D9": c9.replace("C", "D")})
    t.update({"A10": "Ertragsteuersatz", "B10": "%", "C10": 0.3, "D10": 0.03,
              "A12": "Ergebnis vor Steuern", "B12": "Mio. EUR", "C12": 25, "D12": 27,
              "A13": "Ertragsteuern", "B13": "Mio. EUR", "C13": "=-C12*C10", "D13": "=-D12*D10"})
    return t


def test_a_sum_over_rows_in_thousands_and_millions_is_found_with_the_conversion():
    found = _keys(_units())
    hint = found["unitmix:C9"]
    assert hint["level"] == "fehler" and "C7 in Tsd. EUR (38200)" in hint["text"] and "statt 99,6" in hint["text"]
    assert "Genauso in D9" in hint["text"]
    assert hint["fix"]["cells"] == {"C9": "=SUM(C6,C7/1000,C8)", "D9": "=SUM(D6,D7/1000,D8)"}
    assert "unitmix:C9" not in _keys(_units("=C6+C7/1000+C8"))


def test_a_rate_that_is_exactly_a_tenth_of_the_other_year_is_asked_about():
    hint = _keys(_units())["rateshift:D10"]
    assert hint["level"] == "pruefen" and "genau ein Zehntel von C10" in hint["text"]
    assert "D13 rechnet damit und zeigt -0,81, mit 0,3 wären es -8,1" in hint["text"]
    t = _units()
    t["D10"] = 0.3
    assert "rateshift:D10" not in _keys(t)


def _months(end="=EOMONTH(A{r},-1)", year="=YEAR(A{r})+(MONTH(A{r})>10)", quarter='="Q"&ROUNDUP(MONTH(A{r})/3,0)'):
    t = {"A1": "Monatsdaten (Geschäftsjahr Oktober bis September)",
         "A2": "Monat", "B2": "Monatsende", "C2": "Geschäftsjahr", "D2": "Quartal GJ", "E2": "Umsatz (EUR)"}
    for i in range(24):
        r = 3 + i
        t[f"A{r}"] = date(2025 + i // 12, i % 12 + 1, 1)
        t[f"B{r}"] = (end.format(r=r), 2)
        t[f"C{r}"], t[f"D{r}"] = year.format(r=r), quarter.format(r=r)
        t[f"E{r}"] = 1000 + 37 * i
    return t


def test_month_end_fiscal_year_and_fiscal_quarter_follow_what_the_sheet_says():
    found = _keys(_months())
    assert "das Monatsende 1 Monat früher" in found["monthend:B3"]["text"]
    assert found["monthend:B3"]["fix"]["cells"]["B3"] == "=EOMONTH(A3,0)"
    assert "So zählt Oktober noch zum alten Geschäftsjahr" in found["fiscal:C3"]["text"]
    assert found["fiscal:C3"]["fix"]["cells"]["C3"] == "=YEAR(A3)+(MONTH(A3)>9)"
    assert "Kalenderquartal" in found["fquarter:D3"]["text"]
    assert found["fquarter:D3"]["fix"]["cells"]["D3"] == '="Q"&ROUNDUP((MOD(MONTH(A3)-10,12)+1)/3,0)'
    right = _keys(_months("=EOMONTH(A{r},0)", "=YEAR(A{r})+(MONTH(A{r})>9)",
                          '="Q"&ROUNDUP((MOD(MONTH(A{r})-10,12)+1)/3,0)'))
    assert not [k for k in right if k.split(":")[0] in ("monthend", "fiscal", "fquarter")]


def test_a_check_without_tolerance_that_rounds_lets_a_difference_pass():
    t = {"A1": "Umsatz laut Modell", "B1": 1000, "A2": "Umsatz laut Bank", "B2": 1000.4,
         "A3": "Abweichung Bank − Modell", "B3": "=B2-B1",
         "A5": "Bankabgleich (ohne Rundungstoleranz)", "B5": '=IF(ROUND(B3,0)=0,"OK","FEHLER")'}
    hint = _keys(t)["lenient:B5"]
    assert hint["level"] == "fehler" and "B3 ist 0,4" in hint["text"]
    assert hint["fix"]["cells"] == {"B5": '=IF(B3=0,"OK","FEHLER")'}
    t["A5"] = "Bankabgleich"
    assert "lenient:B5" not in _keys(t)


def _keywords(end: int = 31) -> dict:
    t = {"A29": "Negative Stichworte", "A30": "defekt", "A31": "beschädigt", "A32": "falsch",
         "C1": "Kommentar", "D1": "Negativ"}
    for r, text in enumerate(("gut", "defekt", "falsch", "beschädigt", "prima"), 2):
        t[f"C{r}"] = text
        t[f"D{r}"] = f'=IF(COUNTIF($A$30:$A${end},C{r})>0,"Ja","Nein")'
    return t


def test_a_keyword_list_read_with_a_fixed_range_that_stops_early():
    hint = _keys(_keywords())["listshort:D2"]
    assert "„falsch“ (A32)" in hint["text"] and "Genauso in 4 weiteren Formeln" in hint["text"]
    assert "$A$30:$A$32" in hint["fix"]["cells"]["D2"]
    assert "listshort:D2" not in _keys(_keywords(32))


def _invoices(column: str = "E", dupe: bool = True, marked: str = "") -> dict:
    t = {"A1": "Gesamtumsatz gebucht (EUR)", "B1": f"=SUM({column}15:{column}34)",
         "A3": "Kunden", "A4": "Kundennummer", "B4": "Kundenname", "C4": "Umsatz (EUR)"}
    names = [f"KD-{100 + i}" for i in range(6)] + (["KD-101"] if dupe else [])
    for r, key in enumerate(names, 5):
        t[f"A{r}"], t[f"B{r}"] = key, f"Kunde {r}"
        t[f"C{r}"] = f"=SUMIFS($F$15:$F$34,$A$15:$A$34,A{r})"
    t["E3"] = "Summe Kundenübersicht"
    t["F3"] = f"=SUM(C5:C{4 + len(names)})"
    if marked:
        t["D4"] = "Aktuell"
        for r, key in enumerate(names, 5):
            t[f"D{r}"] = "Ja" if key != "KD-101" or r != 6 else "Nein"
    t.update({"A14": "Kundennummer", "B14": "Status", "C14": "Betrag", "D14": "Kurs", "E14": "UmsatzEUR",
              "F14": "UmsatzGebucht"})
    for i, r in enumerate(range(15, 35)):
        status = "Gebucht" if i % 4 else "Storniert"
        t[f"A{r}"], t[f"B{r}"], t[f"C{r}"], t[f"D{r}"] = f"KD-{100 + i % 6}", status, 100 + 13 * i, 1
        t[f"E{r}"] = f"=C{r}*D{r}"
        t[f"F{r}"] = f'=IF(B{r}="Gebucht",E{r},0)'
    return t


def test_a_kpi_that_sums_the_neighbour_column_of_its_name():
    # Liest die Kennzahl die Haupttabelle des Blatts, sagt es statcol; liest sie eine Tabelle weiter unten (die
    # Rechnungen unter der Kundenübersicht), sagte es niemand: Das ist colname.
    found = _keys(_invoices())
    hint = found.get("colname:B1") or found["statcol:B1"]
    assert "UmsatzEUR" in hint["text"] and "UmsatzGebucht" in hint["text"]
    assert not [k for k in _keys(_invoices("F")) if k.endswith(":B1")]


def test_a_duplicate_key_that_gets_its_whole_sum_twice():
    found = _keys(_invoices("F"))
    hint = found["dupkey:C5"]
    assert "„KD-101“ steht aber 2-mal in der Tabelle (Zeilen 6, 11)" in hint["text"] and "F3 summiert" in hint["text"]
    assert hint["fix"]["cells"]["C11"] == "=IF(COUNTIF($A$5:$A11,$A11)=1,SUMIFS($F$15:$F$34,$A$15:$A$34,A11),0)"
    marked = _keys(_invoices("F", marked="Aktuell"))["dupkey:C5"]
    assert marked["fix"]["cells"]["C6"] == '=IF($D6="Ja",SUMIFS($F$15:$F$34,$A$15:$A$34,A6),0)'
    assert "dupkey:C5" not in _keys(_invoices("F", dupe=False))


def _cash(interest: str = "=AVERAGE({c}3,{c}9)*$B$12", rate: str = "$B$13") -> dict:
    t = {"A2": "Position", "B2": 2026, "C2": 2027, "D2": 2028}
    rows = {3: "Cash Anfang", 4: "Operativer Cashflow", 5: "Zinsertrag auf Guthaben (Anfangsbestand)", 6: "Kredit Anfang",
            7: "Kredit Ende", 8: "Tilgung", 9: "Cash Ende", 10: "Zinsaufwand Kredit (Anfangsbestand)"}
    for r, name in rows.items():
        t[f"A{r}"] = name
    t.update({"A12": "Zinssatz Guthaben", "B12": 0.02, "A13": "Zinssatz Kredit", "B13": 0.05})
    for i, c in enumerate("BCD"):
        prev = "ABC"[i]
        t[f"{c}3"] = 100 if i == 0 else f"={prev}9"
        t[f"{c}4"], t[f"{c}8"] = 60, 20
        t[f"{c}5"] = interest.format(c=c)
        t[f"{c}6"] = 200 if i == 0 else f"={prev}7"
        t[f"{c}7"] = f"={c}6-{c}8"
        t[f"{c}9"] = f"={c}3+{c}4-{c}10-{c}8"  # ohne Zinsertrag, sonst ein Kreisbezug
        t[f"{c}10"] = f"={c}6*{rate}"
    return t


def test_interest_on_the_closing_balance_and_on_the_wrong_rate():
    found = _keys(_cash(rate="$B$12"))
    assert "liest B9 (Cash Ende)" in found["anfang:B5"]["text"]
    assert found["anfang:B5"]["fix"]["cells"]["B5"] == "=B3*$B$12"
    assert "rechnet mit B12 (Zinssatz Guthaben)" in found["annahme:B10"]["text"]
    assert found["annahme:B10"]["fix"]["cells"]["B10"] == "=B6*$B$13"
    right = _keys(_cash("={c}3*$B$12"))
    assert not [k for k in right if k.split(":")[0] in ("anfang", "annahme")]


def test_rows_below_a_total_are_not_its_parts():
    t = {"A1": "Geschäftsjahr", "B1": "Q1", "C1": "Q2", "D1": "Summe"}
    for r, year in ((2, 2025), (3, 2026), (4, 2027)):
        t[f"A{r}"], t[f"B{r}"], t[f"C{r}"], t[f"D{r}"] = year, 100 * r, 50 * r, f"=SUM(B{r}:C{r})"
    t.update({"A5": "Summe", "B5": "=SUM(B2:B4)", "C5": "=SUM(C2:C4)", "D5": "=SUM(D2:D4)",
              "A6": "Umsatz laut Bank (EUR, Eingabe)", "D6": 1350.4, "A7": "Abweichung Bank − Modell", "D7": "=D6-D5"})
    assert not [k for k in _keys(t) if k.startswith("label:D5") or k.startswith("range:D5")]
