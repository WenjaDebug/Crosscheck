"""Testblatt „Infekt“ (KI-erstellt): Belegung, Infektionen, Erreger mit Wilson-Intervall,
Antibiotika je 1.000 Pflegetage und ein z-Score gegen den Referenzwert. Fünf Fehler sind eingebaut. Der Prüfer übersah
den z-Score in B62 und meldete die Summen der Erreger als falsch, weil ein Erregername länger als 40 Zeichen ist. Jetzt
kommt jeder Fehler mit der Korrektur, und die saubere Mappe bleibt still."""

from test_dritte_serie import PERCENT, _hints


def _infekt(planted: bool = True) -> tuple:
    t = {"A1": "Infektiologie – Belegungs-, Resistenz- und Kostenanalyse",
         "A2": "Stationsbetrieb Infektiologie, Berichtsjahr 2026. Blau = Eingabe, Schwarz = Formel.",
         "A4": "1. Annahmen"}
    ann = [("Bettenkapazität (Planbetten)", 24, "Betten"), ("Ø Verweildauer je Fall", 9.4, "Tage"),
           ("Ø Erlös je Fall (Fallpauschale)", 4850, "EUR"), ("Personalkosten je Pflegetag", 412, "EUR"),
           ("Soll-Compliance Händedesinfektion", (0.9, PERCENT), "%"), ("Referenzwert HAI je 1.000 Pflegetage", 6.5, "Inf./1.000 PT"),
           ("z-Wert Konfidenzniveau 95 %", 1.96, "z"), ("Berichtsjahr", 2026, "Jahr")]
    for i, (a, b, c) in enumerate(ann, start=5):
        t.update({f"A{i}": a, f"B{i}": b, f"C{i}": c})
    t["A14"] = "2. Monatliche Belegung und Infektionen"
    heads = ["Monat", "Tage", "Zugänge (Fälle)", "Pflegetage", "Auslastung %", "HAI (Fälle)", "HAI je 1.000 PT",
             "Erlöse (EUR)", "Personalkosten (EUR)", "Deckungsbeitrag (EUR)"]
    for c, h in zip("ABCDEFGHIJ", heads):
        t[f"{c}15"] = h
    months = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"]
    zug = [61, 57, 70, 66, 59, 54, 47, 51, 63, 72, 75, 67]
    hai = [3, 2, 4, 3, 2, 1, 2, 3, 4, 2, 3, 3]
    for k, r in enumerate(range(16, 28)):
        t.update({f"A{r}": months[k], f"B{r}": f"=DAY(DATE($B$12,{k + 2},0))", f"C{r}": zug[k],
                  f"D{r}": f"=ROUND(C{r}*$B$6,0)", f"E{r}": (f"=D{r}/($B$5*B{r})", PERCENT), f"F{r}": hai[k],
                  f"G{r}": f"=F{r}/D{r}*1000", f"H{r}": f"=C{r}*$B$7", f"I{r}": f"=D{r}*$B$8", f"J{r}": f"=H{r}-I{r}"})
    if planted:
        t["D21"] = 495
    t.update({"A28": "Summe / Gesamt", "B28": "=SUM(B16:B27)", "C28": "=SUM(C16:C26)" if planted else "=SUM(C16:C27)",
              "D28": "=SUM(D16:D27)", "E28": ("=AVERAGE(E16:E27)" if planted else "=D28/($B$5*B28)", PERCENT),
              "F28": "=SUM(F16:F27)", "G28": "=F28/D28*1000", "H28": "=SUM(H16:H27)", "I28": "=SUM(I16:I27)", "J28": "=SUM(J16:J27)"})
    t["A30"] = "3. Erreger und Resistenzen (Wilson-Konfidenzintervall 95 %)"
    for c, h in zip("ABCDEFGH", ["Erreger", "Isolate (n)", "Anteil Isolate", "Resistent (n)", "Resistenzrate", "KI untere Grenze", "KI obere Grenze", "Rate Vorjahr"]):
        t[f"{c}31"] = h
    err = [("E. coli (ESBL-bildend)", 142, 61, .31), ("Klebsiella pneumoniae (Carbapenem-resistent)", 88, 29, .27),
           ("Staphylococcus aureus (MRSA)", 96, 34, .36), ("Enterococcus faecium (VRE)", 41, 27, .58),
           ("Pseudomonas aeruginosa (MDR)", 57, 18, .29), ("Acinetobacter baumannii (MDR)", 19, 12, .66)]
    for r, (a, n, k, v) in enumerate(err, start=32):
        z = "$B$11"
        t.update({f"A{r}": a, f"B{r}": n, f"C{r}": (f"=B{r}/$B$38", PERCENT), f"D{r}": k, f"E{r}": (f"=D{r}/B{r}", PERCENT),
                  f"F{r}": (f"=(E{r}+{z}^2/(2*B{r})-{z}*SQRT(E{r}*(1-E{r})/B{r}+{z}^2/(4*B{r}^2)))/(1+{z}^2/B{r})", PERCENT),
                  f"G{r}": (f"=(E{r}+{z}^2/(2*B{r})+{z}*SQRT(E{r}*(1-E{r})/B{r}+{z}^2/(4*B{r}^2)))/(1+{z}^2/B{r})", PERCENT),
                  f"H{r}": (v, PERCENT)})
    t.update({"A38": "Summe / Gesamt", "B38": "=SUM(B32:B37)", "C38": ("=SUM(C32:C37)", PERCENT), "D38": "=SUM(D32:D37)",
              "E38": ("=D38/B38", PERCENT)})
    t["A40"] = "4. Antibiotika-Verbrauch und Kosten"
    for c, h in zip("ABCDEF", ["Wirkstoffgruppe", "DDD gesamt", "Kosten je DDD (EUR)", "Kosten gesamt (EUR)", "DDD je 1.000 PT", "Kostenanteil"]):
        t[f"{c}41"] = h
    ab = [("Penicilline mit Inhibitor", 2100, 1.2), ("Cephalosporine 3. Generation", 1650, 2.1), ("Glykopeptide (Vancomycin)", 980, 4.8),
          ("Carbapeneme", 720, 14.5), ("Linezolid", 260, 38), ("Daptomycin", 140, 96), ("Colistin", 95, 22)]
    for r, (a, d, k) in enumerate(ab, start=42):
        t.update({f"A{r}": a, f"B{r}": d, f"C{r}": k, f"D{r}": f"=B{r}*C{r}", f"E{r}": f"=B{r}/$D$28*1000",
                  f"F{r}": (f"=D{r}/$D$49", PERCENT)})
    if planted:
        t["E46"] = "=B46/$D$28*100"
    t.update({"A49": "Summe / Gesamt", "B49": "=SUM(B42:B48)", "D49": "=SUM(D42:D48)", "E49": "=SUM(E42:E48)", "F49": ("=SUM(F42:F48)", PERCENT)})
    t.update({"A51": "5. Hygiene-Indikatoren", "A52": "Beobachtete Gelegenheiten Händedesinfektion", "B52": 1200,
              "A53": "davon korrekt durchgeführt", "B53": 1025, "A54": "Compliance Händedesinfektion", "B54": ("=B53/B52", PERCENT),
              "A55": "Ziel erreicht?", "B55": '=IF(B54>=$B$9,"Ja","Nein")',
              "A58": "HAI-Rate gegen Referenz (Poisson-Näherung)", "A59": "Pflegetage gesamt", "B59": "=D28",
              "A60": "Erwartete HAI bei Referenzrate", "B60": "=$B$10*B59/1000", "A61": "Beobachtete HAI", "B61": "=F28",
              "A62": "z-Score", "B62": "=(B61-B60)/B60" if planted else "=(B61-B60)/SQRT(B60)",
              "A63": "Signifikant unter Referenz?", "B63": '=IF(B62<-$B$11,"Ja","Nein")',
              "A66": "6. Dashboard", "A67": "Fälle gesamt", "B67": "=C28", "A68": "Deckungsbeitrag gesamt (EUR)", "B68": "=J28",
              "A69": "Gesamt-Resistenzrate", "B69": ("=E38", PERCENT), "A70": "Auslastung Jahr", "B70": ("=E28", PERCENT)})
    return ("Infekt", t)


def _mine(found: dict) -> dict:
    return {key: h for (sheet, key), h in found.items() if sheet == "Infekt" and h["level"] in ("fehler", "pruefen")}


def test_all_five_planted_errors_come_with_the_correction():
    found = _mine(_hints(_infekt(True)))
    assert set(found) == {"typedcell:D21", "ratio:E28", "zscore:B62", "parts:C28", "pattern:E46"}, sorted(found)
    assert all(h["level"] == "fehler" and h["fix"] for h in found.values())
    z = found["zscore:B62"]
    assert "teilt die Abweichung aber durch B60 („Erwartete HAI bei Referenzrate“) selbst" in z["text"]
    assert "Richtig wäre -1,9701 statt -0,2929" in z["text"]
    assert z["fix"]["cells"] == {"B62": "=(B61-B60)/SQRT(B60)"}
    ratio = found["ratio:E28"]
    assert ratio["text"].startswith("E28 („Summe / Gesamt“ unter Auslastung %) mittelt die Quoten E16:E27")
    assert ratio["fix"]["cells"] == {"E28": "=D28/($B$5*B28)"}
    assert found["parts:C28"]["fix"]["cells"] == {"C28": "=SUM(C16:C27)"}
    assert found["pattern:E46"]["fix"]["cells"] == {"E46": "=B46/$D$28*1000"}


def test_the_clean_sheet_stays_quiet():
    found = _mine(_hints(_infekt(False)))
    assert found == {}, {k: h["text"] for k, h in found.items()}


def test_a_long_row_name_does_not_cut_the_parts_of_a_total():
    """„Klebsiella pneumoniae (Carbapenem-resistent)“ hat 44 Zeichen: Die Summe darunter zählt trotzdem alle Erreger."""
    found = _mine(_hints(_infekt(False)))
    assert not any(k.startswith("parts:") for k in found)


def test_a_z_score_divides_by_the_standard_deviation_when_the_sheet_has_one():
    name, t = _infekt(False)
    t = dict(t)
    t.update({"A60": "Mittelwert HAI je Monat", "B60": "=AVERAGE(F16:F27)", "A61": "HAI im Dezember", "B61": "=F27",
              "A64": "Standardabweichung HAI je Monat", "B64": "=STDEV(F16:F27)"})
    for label in ("z-Score", "Z-Wert Dezember", "zScore", "z score"):
        t["A62"], t["B62"] = label, "=(B61-B60)/B60"
        z = _mine(_hints((name, t)))["zscore:B62"]
        assert z["fix"]["cells"] == {"B62": "=(B61-B60)/B64"}, (label, z)
        assert "die Standardabweichung in B64 („Standardabweichung HAI je Monat“)" in z["text"]


def test_a_relative_deviation_without_a_spread_is_only_checked():
    name, t = _infekt(False)
    t = dict(t)
    t.update({"A58": "Vergleich", "A59": "", "B59": "", "A60": "Mittelwert HAI je Monat", "B60": "=AVERAGE(F16:F27)",
              "A61": "HAI im Dezember", "B61": "=F27", "A62": "z-Score Dezember", "B62": "=(B61-B60)/B60"})
    z = _mine(_hints((name, t)))["zscore:B62"]
    assert z["level"] == "pruefen" and not z.get("fix")


def test_a_mean_of_rates_outside_a_total_row_is_only_checked():
    name, t = _infekt(False)
    t = dict(t)
    t["A28"], t["E28"] = "Ø Monat", ("=AVERAGE(E16:E27)", PERCENT)
    found = _mine(_hints((name, t)))
    assert found["ratio:E28"]["level"] == "pruefen", found["ratio:E28"]
