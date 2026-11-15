"""Zweite Testrunde mit KI-erstellten Mappen: Tabelle1 (Finanzierung und Cashflow über 5 Perioden) und
Tabelle2 (Investitionsrechnung mit Darlehen, Jahr 0 bis 10), je mit 5 eingebauten Fehlern. Gefunden wurden 3 und 0.

Die Ursachen waren allgemein, nicht die einzelnen Fälle:
- Bedeutungsprüfungen kannten nur eine Form: Steuern addiert, aber nicht jede Auszahlung (Investition positiv im
  Netto-Cashflow); Gordon nur ohne (1 + g), nicht mit (Zins + Wachstum) im Nenner; ein Satz als ganze Zahl (WACC 9
  statt 9 %) gar nicht.
- Zehn gleiche Formeln mit demselben Bereich gaben zehn Hinweise; nach drei war Schluss, und die zu kurze Summe der
  Barwerte darunter fiel hinten heraus.
- Jahr 0 (fast leer, nur Investition und Darlehen) galt als Datenzeile: MAX über die Jahre 1 bis 10 „ließ eine
  Zeile aus“."""

import pytest
from test_dritte_serie import PERCENT, _hints

from voiceagent.xlsx import column_letters

C5 = "BCDEF"


def value_of(t, address):
    from voiceagent import formel
    memo = {}
    def value(pos):
        if pos not in memo:
            e = t.get(f"{column_letters(pos[1])}{pos[0]}")
            e = e[0] if isinstance(e, tuple) else e
            if isinstance(e, str) and e.startswith("="):
                memo[pos] = formel.Evaluator(value, extent=lambda s: (200, 30)).run(formel.parse(e[1:]))
            else:
                memo[pos] = float(e) if isinstance(e, (int, float)) else e
        return memo[pos]
    return value(formel.position(address))


def tabelle1(p=True):
    t = {}
    def put(a, v): t[a] = v
    put("A1", "Testmodell – Finanzierung & Cashflow"); put("A2", "Eingaben in Blau. Beträge in €.")
    put("A4", "Annahmen"); put("A5", "Annahme"); put("B5", "Wert")
    rows = [("Umsatz Jahr 1 (€)", 1000000), ("Umsatzwachstum p.a.", (0.08, PERCENT)), ("Materialquote (% Umsatz)", (0.4, PERCENT)),
            ("Personalkosten Jahr 1 (€)", 250000), ("Personalkostensteigerung p.a.", (0.03, PERCENT)),
            ("Sonstige Kosten (% Umsatz)", (0.1, PERCENT)), ("Steuersatz", (0.3, PERCENT)), ("Investition Jahr 1 (€)", 600000),
            ("Nutzungsdauer (Jahre, linear)", 5), ("WACC (Diskontierungssatz)", 9.0 if p else (0.09, PERCENT)),
            ("Zinssatz Darlehen", (0.05, PERCENT)), ("Darlehen Anfangsbestand (€)", 300000), ("Tilgung p.a. (€)", 60000)]
    for i, (a, v) in enumerate(rows, start=6):
        put(f"A{i}", a); put(f"B{i}", v)
    put("A20", "Modell – GuV, Darlehen, Cashflow (€)")
    put("A21", "Periode")
    for k, c in enumerate(C5, 1): put(f"{c}21", k)
    names = {22: "Umsatz", 23: "Materialkosten", 24: "Personalkosten", 25: "Sonstige Kosten", 26: "EBITDA",
             27: "Abschreibung (AfA)", 28: "EBIT", 29: "Zinsaufwand", 30: "Ergebnis vor Steuern", 31: "Ertragsteuern",
             32: "Jahresüberschuss", 34: "Darlehen", 35: "Anfangsbestand", 36: "Tilgung", 37: "Endbestand",
             38: "Zinsaufwand (auf Anfangsbestand)", 40: "Cashflow", 41: "Jahresüberschuss", 42: "AfA (Rückrechnung)",
             43: "Operativer Cashflow", 44: "Investition", 45: "Tilgung Darlehen", 46: "Netto-Cashflow",
             47: "Kumulierter Cashflow"}
    for r, a in names.items(): put(f"A{r}", a)
    for i, c in enumerate(C5):
        pv = C5[i - 1] if i else None
        put(f"{c}22", "=$B$6" if not pv else (1400000 if (p and c == "E") else f"={pv}22*(1+$B$7)"))
        put(f"{c}23", f"=-{c}22*$B$8")
        put(f"{c}24", "=-$B$9" if not pv else f"={pv}24*(1+$B$10)")
        put(f"{c}25", f"=-{c}22*$B$11")
        put(f"{c}26", f"=SUM({c}22:{c}24)" if p else f"=SUM({c}22:{c}25)")
        put(f"{c}27", "=-$B$13/$B$14")
        put(f"{c}28", f"={c}26+{c}27"); put(f"{c}29", f"={c}38"); put(f"{c}30", f"={c}28+{c}29")
        put(f"{c}31", f"=-MAX(0,{c}30)*$B$12"); put(f"{c}32", f"={c}30+{c}31")
        put(f"{c}35", "=$B$17" if not pv else f"={pv}37")
        put(f"{c}36", f"=-MIN($B$18,{c}35)"); put(f"{c}37", f"={c}35+{c}36")
        put(f"{c}38", "=-C35*$B$16" if (p and c == "D") else f"=-{c}35*$B$16")
        put(f"{c}41", f"={c}32"); put(f"{c}42", f"=-{c}27"); put(f"{c}43", f"={c}41+{c}42")
        put(f"{c}44", ("=$B$13" if p else "=-$B$13") if not pv else 0)
        put(f"{c}45", f"={c}36"); put(f"{c}46", f"={c}43+{c}44+{c}45")
        put(f"{c}47", f"={c}46" if not pv else f"={pv}47+{c}46")
    put("A50", "Auswertung & Kennzahlen"); put("A52", "Kennzahl"); put("B52", "Wert")
    put("A53", "EBITDA-Marge Jahr 1"); put("B53", ("=B26/B22", PERCENT))
    put("A54", "Jahresüberschuss Summe (5 J.)"); put("B54", "=SUM(B32:F32)")
    put("A55", "Darlehen Restbestand Ende J5"); put("B55", "=F37")
    put("A56", "Kumulierter Cashflow Ende J5"); put("B56", "=F47")
    put("A57", "Kapitalwert Netto-Cashflow (WACC)"); put("B57", "=NPV(B15,B46:F46)")
    put("A60", "Jahr"); put("B60", "Umsatz"); put("C60", "EBITDA"); put("D60", "Jahresüberschuss")
    for k, c in enumerate(C5, 1):
        r = 60 + k
        put(f"A{r}", f"Jahr {k}"); put(f"B{r}", f"={c}22"); put(f"C{r}", f"={c}26"); put(f"D{r}", f"={c}32")
    return "Tabelle1", t


def tabelle2(p=True):
    t = {}
    def put(a, v): t[a] = v
    put("A1", "Investitionsrechnung mit Darlehen – Annuität, Tilgungsplan, Kapitalwert & IRR")
    put("A2", "Eingaben in Blau. Beträge in €. Jahr 0 = Investitionszeitpunkt.")
    put("A4", "Eingaben")
    rows = [("Investition (€)", 500000), ("Fremdfinanzierungsquote", (0.6, PERCENT)), ("Darlehenszins", (0.06, PERCENT)),
            ("Laufzeit Darlehen (Jahre)", 10), ("Cashflow vor Schuldendienst Jahr 1 (€)", 80000),
            ("Cashflow-Wachstum p.a.", (0.04, PERCENT)), ("Nutzungsdauer AfA (Jahre, linear)", 10),
            ("Steuersatz", (0.25, PERCENT)), ("Diskontsatz (Eigenkapitalkosten)", (0.08, PERCENT)),
            ("Terminalwachstum", (0.02, PERCENT))]
    for i, (a, v) in enumerate(rows, start=5):
        put(f"A{i}", a); put(f"B{i}", v)
    put("A16", "Berechnungen")
    put("A17", "Darlehen (€)"); put("B17", "=B5*B6")
    put("A18", "Eigenkapital (€)"); put("B18", "=B5-B17")
    put("A19", "Annuität p.a. (€)"); put("B19", "=PMT(B7,B8,-B17)")
    put("A20", "Terminalwert Jahr 10 (€)"); put("B20", "=I33*(1+B14)/(B13+B14)" if p else "=I33*(1+B14)/(B13-B14)")
    heads = ["Jahr", "Anfangsschuld", "Zinsen", "Tilgung", "Endschuld", "CF vor Schuldendienst", "AfA", "Steuer",
             "CF nach Schuldendienst", "Diskontfaktor", "Barwert CF", "EK-Cashflow"]
    for i, h in enumerate(heads): put(f"{chr(65 + i)}22", h)
    put("A23", 0); put("E23", "=B17"); put("J23", "=1/(1+$B$13)^A23"); put("L23", "=-B18")
    for r in range(24, 34):
        put(f"A{r}", r - 23)
        put(f"B{r}", f"=E{r - 1}")
        put(f"C{r}", f"=B{r}*$B$8" if (p and r == 30) else f"=B{r}*$B$7")
        put(f"D{r}", f"=$B$19-C{r}"); put(f"E{r}", f"=B{r}-D{r}")
        put(f"F{r}", f"=$B$9*(1+$B$10)^(A{r}-1)"); put(f"G{r}", "=$B$5/$B$11")
        put(f"H{r}", f"=-MAX(0,F{r}-G{r}-C{r})*$B$12")
        put(f"I{r}", f"=F{r}-C{r}-D{r}+H{r}")
        put(f"J{r}", f"=1/(1+$B$13)^A{r - 1}" if (p and r == 27) else f"=1/(1+$B$13)^A{r}")
        put(f"K{r}", f"=I{r}*J{r}")
        put(f"L{r}", f"=I{r}+IF(A{r}=MAX($A$24:$A$33),$B$20,0)")
    if p:
        # H27 als feste Zahl mit dem Wert, den die Formel gerade ergibt (gerundet wie getippt)
        t["H27"] = round(value_of(t, "H27"))
    put("A36", "Ergebnisse")
    put("A37", "Barwert operativer CF (Jahre 1–10)"); put("B37", "=SUM(K24:K32)" if p else "=SUM(K24:K33)")
    put("A38", "Barwert Terminalwert"); put("B38", "=B20*J33")
    put("A39", "Kapitalwert Eigenkapital (NPV)"); put("B39", "=B37+B38-B18")
    put("A41", "IRR Eigenkapital"); put("B41", ("=IRR(L23:L33)", PERCENT))
    put("A42", "Restschuld Ende Laufzeit (soll 0)"); put("B42", "=E33")
    put("A43", "Tilgungskontrolle: Summe Tilgung – Darlehen (soll 0)"); put("B43", "=SUM(D24:D33)-B17")
    put("A44", "Minimum DSCR"); put("B44", "=MIN(F24:F33)/B19")
    return "Tabelle2", t


PLANTED = {
    "Tabelle1": {"auszahlung:B44": {"B44": "=-$B$13"}, "ganzzahl:B15": {"B15": 0.09}, "gap:B26": None,
                 "pattern:D38": None, "typed:E22": None},
    "Tabelle2": {"gordon:B20": {"B20": "=I33*(1+B14)/(B13-B14)"}, "pattern:C30": None, "pattern:J27": None,
                 "range:B37": {"B37": "=SUM(K24:K33)"}, "typedcell:H27": None},
}
BUILD = {"Tabelle1": tabelle1, "Tabelle2": tabelle2}


def _found(name, planted=True, change=None):
    sheet, t = BUILD[name](planted)
    t = {k: v for k, v in {**t, **(change or {})}.items() if v is not None}
    return {k: h for (_, k), h in _hints((sheet, t)).items()}


@pytest.mark.parametrize("name", list(PLANTED))
def test_all_five_errors_are_found(name):
    found = _found(name)
    assert set(found) == set(PLANTED[name]), sorted(found)
    for key, cells in PLANTED[name].items():
        assert found[key]["level"] == "fehler", (key, found[key])
        if cells:
            assert found[key]["fix"]["cells"] == cells, (key, found[key]["fix"])


@pytest.mark.parametrize("name", list(PLANTED))
def test_the_right_version_stays_quiet(name):
    assert _found(name, False) == {}


@pytest.mark.parametrize("name,key", [(n, k) for n, keys in PLANTED.items() for k, v in keys.items() if v])
def test_each_correction_removes_its_hint(name, key):
    change = {a: ((v, PERCENT) if isinstance(v, float) else v) for a, v in PLANTED[name][key].items()}
    after = _found(name, change=change)
    assert key not in after and set(after) == set(PLANTED[name]) - {key}, sorted(after)


def test_the_hints_say_what_is_wrong():
    one, two = _found("Tabelle1"), _found("Tabelle2")
    out = one["auszahlung:B44"]["text"]
    assert "B46 (Netto-Cashflow) addiert B44 (Investition) mit 600000" in out and "B45 (Tilgung Darlehen) -60000" in out
    assert "B46 -389500 statt 810500" in out
    rate = one["ganzzahl:B15"]["text"]
    assert "B15 (WACC (Diskontierungssatz)) steht als 9 da" in rate and "mit 900 %" in rate
    gordon = two["gordon:B20"]["text"]
    assert "Der Nenner ist (B13 + B14), richtig ist (Zins − Wachstum) = (B13 − B14)" in gordon
    assert "^ Jahr 3 statt" in two["pattern:J27"]["text"] and "Jahr Jahr" not in two["pattern:J27"]["text"]


def test_year_zero_is_no_missing_row():
    # MAX über die Jahre 1 bis 10 lässt Jahr 0 mit Absicht aus; die leere Steuer in Jahr 0 fehlt nicht.
    found = _found("Tabelle2", False)
    assert not any(k.startswith(("range:", "missing:")) for k in found)


def test_one_cause_in_many_rows_is_one_hint_and_hides_nothing():
    # Zehn gleiche zu kurze Bereiche (ohne Jahr 10) sind ein Hinweis; die zu kurze Summe in B37 bleibt sichtbar.
    rows = {f"L{r}": f"=I{r}+IF(A{r}=MAX($A$23:$A$32),$B$20,0)" for r in range(24, 34)}
    found = _found("Tabelle2", change=rows)
    ranges = [k for k in found if k.startswith("range:")]
    assert "range:B37" in ranges and len(ranges) == 2, sorted(found)
    many = next(found[k] for k in ranges if k != "range:B37")
    assert "Ebenso die 9 gleichen Formeln" in many["text"] and len(many["fix"]["cells"]) == 10


def test_what_looks_alike_but_is_right_stays_quiet():
    # Auszahlungen positiv und abgezogen, AfA als Rückrechnung positiv addiert: richtig.
    flows = {f"{c}44": ("=$B$13" if c == "B" else 0) for c in "BCDEF"} | {f"{c}45": f"=-{c}36" for c in "BCDEF"} \
        | {f"{c}46": f"={c}43-{c}44-{c}45" for c in "BCDEF"}
    assert "auszahlung:B44" not in _found("Tabelle1", False, flows)
    # Ein Darlehen, das zufließt, steht positiv im Cashflow.
    loan = {"A48": "Darlehensaufnahme", "B48": "=$B$17", "B46": "=B43+B44+B45+B48"}
    assert not any(k.startswith("auszahlung:") for k in _found("Tabelle1", False, loan))
    # Ein Satz als ganze Zahl, den jede Formel durch 100 teilt, ist gewollt.
    whole = {"B15": 9, "B57": "=NPV(B15/100,B46:F46)"}
    assert "ganzzahl:B15" not in _found("Tabelle1", False, whole)
    # Eine Laufzeit oder ein Betrag zwischen den Sätzen ist kein Satz.
    assert _found("Tabelle1", False, {"A19": "Laufzeit (Jahre)", "B19": 10}) == {}


# Ein Fehler nach dem anderen in die richtige Fassung eingebaut, aus jeder Art, die vorher blind war.
EINZELN = [
    # Periode 1 rechnet anders (Startwert): Die Formel der Folgejahre ist trotzdem das Muster.
    ("Tabelle1", {"D22": "=C22*(1+$B$8)"}, "pattern:D22", "fehler"),
    # Die Summe in der letzten Spalte eines Zeilenmodells ist keine Spaltensumme.
    ("Tabelle1", {"F26": "=SUM(F23:F25)"}, "pattern:F26", "fehler"),
    # Eine feste Zahl in der ersten Prognosespalte, wenn die Spalte sonst überall rechnet.
    ("Tabelle1", {"B23": -400000}, "typed:B23", "pruefen"),
    # Ein Satz als ganze Zahl mit einem Namen ohne „Satz“ oder „Zins“.
    ("Tabelle1", {"B10": 3}, "ganzzahl:B10", "fehler"),
    # Jahr 1 eingetippt, ab Jahr 2 rechnet die Spalte (Jahr 0 rechnet anders).
    ("Tabelle2", {"L24": 123456}, "typedcell:L24", "fehler"),
]


@pytest.mark.parametrize("name,change,key,level", EINZELN)
def test_each_kind_of_mistake_is_found_alone(name, change, key, level):
    found = _found(name, False, change)
    assert set(found) == {key} and found[key]["level"] == level, sorted(found)


def test_npv_with_year_zero_added_beside_is_right():
    # =B2+NBW(Zins;B3:B7): Jahr 0 steht unabgezinst daneben, der Bereich lässt es mit Absicht aus.
    found = _found("Tabelle2", False, {"B39": "Kapitalwert", "C39": "=K23+NPV($B$13,I24:I33)"})
    assert not any(k.startswith("range:") for k in found), sorted(found)
