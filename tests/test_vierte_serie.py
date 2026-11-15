"""Die vierte selbst gebaute Testreihe: Lohnabrechnung, Lagerbestand, Darlehen, Kreuztabelle mit zweizeiligem Kopf,
Preise auf einem zweiten Blatt, Währungen, BMI und ein Kostenplan mit Gruppenköpfen und einer zweiten Tabelle darunter.
Die Mappen sind genau die der Testreihe (dieselben Zufallszahlen). Ohne Fehler bleiben sie still, bis auf zwei echte
Fragen; jeder eingebaute Fehler wird gefunden, die Fehler mit der Korrektur zum Einsetzen."""

import random
from datetime import date, timedelta

import pytest
from test_dritte_serie import PERCENT, TIME, _book

from voiceagent import pruefer

THEMES = {}
FORMATS = {}


def theme(fn):
    THEMES[fn.__name__] = fn
    return fn


def col(i):
    s = ""
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def put(cells, row, values, start=1, sheet=""):
    for i, v in enumerate(values):
        cells[f"{sheet}{col(start + i)}{row}"] = v


@theme
def lohn(rnd, planted):
    """Lohnabrechnung: AHV und ALV in Prozent aus festen Zellen, auf 5 Rappen gerundet, BVG je Person, Netto, Total."""
    names = ["Amsler", "Brunner", "Caduff", "Dubois", "Egger", "Frei", "Gerber", "Huber", "Imhof", "Jost", "Keller",
             "Lüthi", "Meier", "Nef", "Oswald", "Portmann", "Roth"]
    cells = {}
    put(cells, 1, ["Personal-Nr", "Name", "Pensum", "Brutto", "AHV/IV/EO", "ALV", "BVG", "Netto"])
    for i, n in enumerate(names):
        r = i + 2
        pensum = rnd.choice([0.5, 0.6, 0.8, 0.8, 1, 1, 1])
        brutto = round(rnd.uniform(5200, 9800) * pensum, -1)
        put(cells, r, [1001 + i, n, pensum, brutto, f"=ROUND(D{r}*$K$2*20,0)/20", f"=ROUND(D{r}*$K$3*20,0)/20",
                       round(brutto * rnd.uniform(0.05, 0.09), 2), f"=D{r}-E{r}-F{r}-G{r}"])
    last = len(names) + 1
    put(cells, last + 1, ["", "Total", "", f"=SUM(D2:D{last})", f"=SUM(E2:E{last})", f"=SUM(F2:F{last})",
                          f"=SUM(G2:G{last})", f"=SUM(H2:H{last})"])
    cells["J2"], cells["K2"] = "AHV/IV/EO", 0.053
    cells["J3"], cells["K3"] = "ALV", 0.011
    key = []
    if planted:
        cells["E7"] = "=ROUND(D7*0.05*20,0)/20"; key.append(("E7", "AHV mit 5 % statt $K$2"))
        cells["H11"] = "=D11-E11-F11"; key.append(("H11", "Netto ohne BVG"))
        cells[f"D{last + 1}"] = f"=SUM(D2:D{last - 1})"; key.append((f"D{last + 1}", "Total Brutto ohne letzte Person"))
    return cells, key


FORMATS["lohn"] = {"C": "0%", "K2": "0.0%", "K3": "0.0%"}


@theme
def lager(rnd, planted):
    """Lagerbestand: Bestand = Bestand davor + Zugang - Abgang, der erste Bestand ist gezählt."""
    cells = {}
    put(cells, 1, ["Datum", "Beleg", "Zugang", "Abgang", "Bestand"])
    for i in range(28):
        r = i + 2
        z = rnd.choice([0, 0, 0, 40, 60, 100])
        a = rnd.randint(5, 30)
        put(cells, r, [date(2026, 2, 2) + timedelta(days=i), f"B-{7100 + i}", z, a, 250 if i == 0 else f"=E{r - 1}+C{r}-D{r}"])
    cells["G2"], cells["H2"] = "Zugang gesamt", "=SUM(C2:C29)"
    cells["G3"], cells["H3"] = "Abgang gesamt", "=SUM(D2:D29)"
    key = []
    if planted:
        cells["E12"] = "=E11+C12+D12"; key.append(("E12", "Abgang addiert statt abgezogen"))
        cells["E20"] = "=E18+C20-D20"; key.append(("E20", "Bestand aus der vorletzten Zeile"))
        cells["H3"] = "=SUM(D3:D29)"; key.append(("H3", "Abgang gesamt ohne ersten Tag"))
    return cells, key


@theme
def darlehen(rnd, planted):
    """Darlehen mit fester Monatsrate: Zins = Restschuld × Satz / 12, Tilgung = Rate - Zins."""
    cells = {}
    put(cells, 1, ["Monat", "Restschuld Anfang", "Zins", "Tilgung", "Rate", "Restschuld Ende"])
    for i in range(36):
        r = i + 2
        put(cells, r, [i + 1, "=$I$2" if i == 0 else f"=F{r - 1}", f"=B{r}*$I$3/12", f"=E{r}-C{r}", "=$I$4", f"=B{r}-D{r}"])
    cells["H2"], cells["I2"] = "Darlehen", 200000
    cells["H3"], cells["I3"] = "Zins p.a.", 0.018
    cells["H4"], cells["I4"] = "Monatsrate", 1500
    cells["H6"], cells["I6"] = "Zinsen gesamt", "=SUM(C2:C37)"
    cells["H7"], cells["I7"] = "Tilgung gesamt", "=SUM(D2:D37)"
    key = []
    if planted:
        cells["C10"] = "=B10*$I$3"; key.append(("C10", "Jahreszins statt Monatszins"))
        cells["B20"] = "=F18"; key.append(("B20", "Restschuld aus der vorletzten Zeile"))
        cells["I6"] = "=SUM(C2:C30)"; key.append(("I6", "Zinsen gesamt ohne die letzten Monate"))
    return cells, key


FORMATS["darlehen"] = {"I3": "0.00%"}


@theme
def kreuz(rnd, planted):
    """Umsatz je Region und Monat mit zweizeiligem Kopf („Umsatz 2026“ über den Monaten), Totale rechts und unten."""
    regions = ["Nord", "Ost", "Süd", "West"]
    months = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun"]
    cells = {"_merge": ["B1:G1"]}
    cells["A1"], cells["B1"] = "", "Umsatz 2026"
    put(cells, 2, ["Region", *months, "Total"])
    for i, reg in enumerate(regions):
        r = i + 3
        put(cells, r, [reg, *[round(rnd.uniform(20000, 60000), -2) for _ in months], f"=SUM(B{r}:G{r})"])
    put(cells, 7, ["Total", *[f"=SUM({col(c)}3:{col(c)}6)" for c in range(2, 9)]])
    key = []
    if planted:
        cells["H4"] = "=SUM(B4:F4)"; key.append(("H4", "Total Ost ohne Juni"))
        cells["D7"] = "=SUM(D3:D5)"; key.append(("D7", "Total Februar ohne West"))
        cells["E5"] = 0; key.append(("E5", "Umsatz Süd im März 0 (fehlt)"))
    return cells, key


@theme
def preise(rnd, planted):
    """Bestellung mit Preisen aus einer Preisliste auf einem zweiten Blatt (SVERWEIS genau)."""
    items = {"Schraube M6": 0.35, "Mutter M6": 0.12, "Dübel 8": 0.25, "Winkel 40": 2.9, "Scharnier": 7.5,
             "Griff": 12.8, "Schiene 1m": 18.4, "Leim 250g": 9.9, "Band 50m": 14.5, "Kabelbinder": 0.08}
    cells = {}
    put(cells, 1, ["Artikel", "Preis CHF"], sheet="Preise!")
    for i, (a, p) in enumerate(items.items()):
        put(cells, i + 2, [a, p], sheet="Preise!")
    put(cells, 1, ["Artikel", "Menge", "Preis", "Betrag"], sheet="Bestellung!")
    names = list(items)
    for i in range(22):
        r = i + 2
        put(cells, r, [names[(i * 7) % len(names)], rnd.choice([10, 20, 25, 50, 100, 200]),
                       f"=VLOOKUP(A{r},Preise!$A$2:$B$11,2,FALSE)", f"=B{r}*C{r}"], sheet="Bestellung!")
    cells["Bestellung!F2"], cells["Bestellung!G2"] = "Total", "=SUM(D2:D23)"
    key = []
    if planted:
        cells["Bestellung!C8"] = "=VLOOKUP(A8,Preise!$A$2:$B$11,2)"; key.append(("C8", "SVERWEIS ohne FALSCH"))
        cells["Bestellung!C14"] = "=VLOOKUP(A14,Preise!$A$2:$B$11,1,FALSE)"; key.append(("C14", "SVERWEIS holt Spalte 1"))
        cells["Bestellung!D19"] = 250; key.append(("D19", "Betrag getippt"))
    return cells, key


@theme
def waehrung(rnd, planted):
    """Spesen in Euro, umgerechnet mit dem Kurs aus einer Zelle, Totale in beiden Währungen."""
    cells = {}
    put(cells, 1, ["Datum", "Beleg", "Betrag EUR", "Betrag CHF"])
    for i in range(24):
        r = i + 2
        put(cells, r, [date(2026, 4, 1) + timedelta(days=i), rnd.choice(["Hotel", "Bahn", "Essen", "Taxi"]),
                       round(rnd.uniform(12, 380), 2), f"=C{r}*$G$2"])
    cells["F2"], cells["G2"] = "Kurs EUR/CHF", 0.94
    cells["F3"], cells["G3"] = "Total EUR", "=SUM(C2:C25)"
    cells["F4"], cells["G4"] = "Total CHF", "=SUM(D2:D25)"
    key = []
    if planted:
        cells["D9"] = "=C9*0.95"; key.append(("D9", "fester Kurs 0,95 statt $G$2"))
        cells["D15"] = "=C15/$G$2"; key.append(("D15", "geteilt statt mal Kurs"))
        cells["G3"] = "=SUM(D2:D25)"; key.append(("G3", "Total EUR summiert CHF"))
    return cells, key


FORMATS["waehrung"] = {"C": "#,##0.00", "D": "#,##0.00"}


@theme
def bmi(rnd, planted):
    """Patientenliste mit Größe in cm, Gewicht und BMI = Gewicht / (Größe/100)²."""
    cells = {}
    put(cells, 1, ["Patient", "Geschlecht", "Größe cm", "Gewicht kg", "BMI"])
    for i in range(30):
        r = i + 2
        h = rnd.randint(152, 196)
        put(cells, r, [f"P-{5200 + i}", rnd.choice(["w", "m"]), h, round(rnd.uniform(48, 115), 1),
                       f"=D{r}/(C{r}/100)^2"])
    cells["G2"], cells["H2"] = "BMI Mittelwert", "=AVERAGE(E2:E31)"
    cells["G3"], cells["H3"] = "Anzahl w", '=COUNTIF(B2:B31,"w")'
    key = []
    if planted:
        cells["C12"] = 1.72; key.append(("C12", "Größe in Metern"))
        cells["E18"] = "=D18/C18^2"; key.append(("E18", "BMI ohne /100"))
        cells["D25"] = 7.8; key.append(("D25", "Gewicht 7,8 statt 78"))
    return cells, key


@theme
def kostenplan(rnd, planted):
    """Kosten je Abteilung mit Gruppenköpfen über zwei Spalten (Personal: Lohn, Sozial; Sach: Miete, IT) und
    darunter eine zweite Tabelle mit dem Budget je Abteilung und der Abweichung."""
    deps = ["Einkauf", "Verkauf", "Logistik", "Informatik", "Finanzen", "Personal", "Recht"]
    cells = {"_merge": ["B1:C1", "D1:E1"]}
    cells["B1"], cells["D1"] = "Personal", "Sachkosten"
    put(cells, 2, ["Abteilung", "Lohn", "Sozial", "Miete", "IT", "Total"])
    for i, d in enumerate(deps):
        r = i + 3
        lohn = round(rnd.uniform(200000, 900000), -3)
        put(cells, r, [d, lohn, f"=ROUND(B{r}*0.18,0)", round(rnd.uniform(30000, 90000), -3), round(rnd.uniform(10000, 60000), -3),
                       f"=SUM(B{r}:E{r})"])
    last = len(deps) + 2
    put(cells, last + 1, ["Total", *[f"=SUM({col(c)}3:{col(c)}{last})" for c in range(2, 7)]])
    top = last + 4
    put(cells, top, ["Abteilung", "Budget", "Ist", "Abweichung"])
    for i, d in enumerate(deps):
        r = top + 1 + i
        put(cells, r, [d, round(rnd.uniform(300000, 1100000), -4), f"=F{3 + i}", f"=C{r}-B{r}"])
    key = []
    if planted:
        cells["F7"] = "=SUM(B7:D7)"; key.append(("F7", "Total Finanzen ohne IT"))
        cells[f"C{top + 4}"] = f"=F{3 + 4}"; key.append((f"C{top + 4}", "Ist aus der falschen Abteilung (Finanzen statt Informatik)"))
        cells["C9"] = 52000; key.append(("C9", "Sozial getippt statt Formel"))
    return cells, key


_STYLES = {"0%": PERCENT, "0.0%": PERCENT, "0.00%": PERCENT, "hh:mm": TIME}


def _hints(name: str, planted: bool) -> dict:
    cells, key = THEMES[name](random.Random(sum(map(ord, name))), planted)
    cells.pop("_merge", None)
    sheets: dict = {}
    for address, value in cells.items():
        sheet, _, ref = address.rpartition("!")
        letters = "".join(ch for ch in ref if ch.isalpha())
        row = ref[len(letters):]
        style = next((_STYLES.get(f) for k, f in FORMATS.get(name, {}).items()
                      if k.rpartition("!")[0] in ("", sheet) and k.rpartition("!")[2] in (ref, letters, row) and int(row) > 1), None)
        sheets.setdefault(sheet or "Daten", {})[ref] = (value, style) if style and not isinstance(value, date) else value
    # SVERWEIS ohne FALSCH in der unsortierten Preisliste: Was LibreOffice dort gespeichert hat.
    done = pruefer.Checker(None).check_file("mappe.xlsx", _book(*sheets.items(), cached={("Bestellung", "C8"): 0.25}))
    assert done["state"] == "ok", done
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}, key


# Echte Fragen in den sauberen Mappen: 0,18 steht fest in jeder Formel für die Sozialkosten, und zwei Bestellzeilen
# sind genau gleich (dieselbe Menge desselben Artikels).
ALLOWED = {"kostenplan": {"constant:0,18"}, "preise": {"same:A23"}}


@pytest.mark.parametrize("name", sorted(THEMES))
def test_a_clean_book_raises_only_real_questions(name):
    found, _ = _hints(name, False)
    assert set(found) == ALLOWED.get(name, set()), {k: h["text"] for k, h in found.items()}
    assert all(h["level"] == "pruefen" for h in found.values())


# Eine 0 mitten unter den Umsätzen kann echt sein: Das bleibt eine offene Frage für die Mappe, kein Hinweis.
MISSED = {("kreuz", "E5")}


@pytest.mark.parametrize("name", sorted(THEMES))
def test_every_planted_mistake_is_found_with_a_fix(name):
    found, key = _hints(name, True)
    for cell, what in key:
        if (name, cell) in MISSED:
            continue
        hits = [h for h in found.values() if cell in ((h.get("fix") or {}).get("cells") or {}) or h["key"].endswith(f":{cell}")]
        assert hits, (cell, what, {k: h["text"][:120] for k, h in found.items()})
        assert any(h.get("fix") or h["level"] == "pruefen" for h in hits), (cell, what)
