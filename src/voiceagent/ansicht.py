"""Wie das Seitenfenster einen Hinweis zeigt: ein kurzer Titel, wo er steht, der erklärende Satz und, wo der Text zwei
Formeln oder zwei Werte gegenüberstellt („=E25-E23 statt =E25-E24“, „D29 zeigt darum 418,97 statt 578,57“), diese
beiden für sich, damit man sie auf einen Blick vergleicht. Der Text des Hinweises selbst bleibt, wie er ist (für den
Prüfbericht und die Prüfung einer Datei); view() gibt nur zusätzliche Felder dazu.

Die Gegenüberstellung liest den Satz des Hinweises, darum kennt sie jede Formulierung auf Deutsch und auf Englisch
(_FORMULAS, _SHOWS). Wer einen dieser Sätze umformuliert, passt die Muster hier an.
"""

from __future__ import annotations

import re

from .bericht import kind_name
from .sprache import T

# Kurz, wie eine Überschrift: was falsch ist (deutsch, englisch). Fehlt eine Art hier, steht ihr Name aus dem
# Prüfbericht da.
TITLES = {
    "pattern": ("Formel weicht vom Muster ab", "Formula breaks the pattern"),
    "error": ("Fehlerwert", "Error value"),
    "gap": ("Summe lässt Zeilen aus", "Sum leaves out rows"),
    "sum": ("Summe stimmt nicht", "Sum does not add up"),
    "round": ("Gerundete Werte gehen nicht auf", "Rounded values do not add up"),
    "typed": ("Feste Zahlen statt Formeln", "Fixed numbers instead of formulas"),
    "typedcell": ("Feste Zahl statt Formel", "Fixed number instead of formula"),
    "constant": ("Feste Zahl in mehreren Formeln", "Fixed number in several formulas"),
    "copied": ("Abgetippte Zahl statt Bezug", "Typed-in number instead of reference"),
    "circle": ("Kreisbezug", "Circular reference"),
    "double": ("Zwischensumme doppelt gezählt", "Subtotal counted twice"),
    "recalc": ("Wert passt nicht zur Formel", "Value does not match the formula"),
    "emptyref": ("Formel liest eine leere Zelle", "Formula reads an empty cell"),
    "empty": ("Formel fehlt in einzelnen Zeilen", "Formula missing in some rows"),
    "complex": ("Schwer prüfbare Formel", "Formula hard to check"),
    "approx": ("Ungefähre Suche, Tabelle unsortiert", "Approximate lookup, table unsorted"),
    "nearest": ("Suche nimmt still einen anderen Eintrag", "Lookup silently takes another entry"),
    "hidden": ("WENNFEHLER verdeckt einen Fehler", "IFERROR hides an error"),
    "twice": ("Suchschlüssel doppelt", "Lookup key appears twice"),
    "sliding": ("Suchtabelle rutscht mit", "Lookup table shifts"),
    "outlier": ("Auffälliger Wert", "Unusual value"),
    "scale": ("Komma oder Einheit verrutscht", "Decimal point or unit shifted"),
    "rowunit": ("Andere Einheit als die übrigen", "Different unit from the rest"),
    "formeltext": ("Formel steht als Text da", "Formula stored as text"),
    "extern": ("Werte aus einer anderen Datei", "Values from another file"),
    "copylink": ("Bezug zeigt noch auf die alte Datei", "Reference still points to the old file"),
    "mixed": ("Anteile gemischt geschrieben", "Shares written in mixed ways"),
    "share": ("Anteile ergeben nicht 100 %", "Shares do not add up to 100%"),
    "text": ("Text in einer Zahlenspalte", "Text in a number column"),
    "dupe": ("Name doppelt", "Name appears twice"),
    "missing": ("Fehlende Werte", "Missing values"),
    "placeholder": ("Platzhalter statt Zahl", "Placeholder instead of number"),
    "code": ("Code für „fehlt“ zählt mit", "Code for “missing” is counted"),
    "impossible": ("Unmöglicher Wert", "Impossible value"),
    "regel": ("Eigene Regel verletzt", "Your own rule broken"),
    "order": ("Ende vor dem Anfang", "End before the start"),
    "future": ("Datum in der Zukunft", "Date in the future"),
    "id": ("Nummer mehrmals vergeben", "Number used more than once"),
    "same": ("Zeile doppelt", "Duplicate row"),
    "units": ("Zwei Einheiten gemischt", "Two units mixed"),
    "rate": ("Satz weicht ab", "Rate differs"),
    "range": ("Bereich lässt Zeilen aus", "Range leaves out rows"),
    "group": ("Bereich passt nicht zur Gruppe", "Range does not match the group"),
    "label": ("Bereich passt nicht zum Namen", "Range does not match the name"),
    "shift": ("Bereiche gegeneinander verschoben", "Ranges shifted against each other"),
    "partial": ("Nur ein Teil der Tabelle", "Only part of the table"),
    "unfilled": ("Formeln fehlen in neuen Zeilen", "Formulas missing in new rows"),
    "criterion": ("Bedingung findet nichts", "Criterion finds nothing"),
    "count": ("Feste Anzahl passt nicht", "Fixed count does not fit"),
    "total": ("Summenzeile zählt mit", "Total row is counted"),
    "slide": ("Bezug rutscht beim Herunterziehen", "Reference shifts when filled down"),
    "astext": ("Zahl als Text", "Number stored as text"),
    "ratio": ("Mittelwert von Quoten", "Average of ratios"),
    "quantil": ("Quartil mit falscher Nummer", "Quartile with the wrong number"),
    "bonf": ("Bonferroni ohne Korrektur", "Bonferroni without correction"),
    "cagr": ("Wachstumsrate mit falscher Jahreszahl", "Growth rate with the wrong number of years"),
    "wmean": ("Gewichtetes Mittel durch die Anzahl", "Weighted average divided by the count"),
    "avgmeans": ("Mittel von Gruppenmitteln", "Average of group averages"),
    "corrshift": ("Verschobene Bereiche", "Shifted ranges"),
    "counttext": ("ANZAHL über Text", "COUNT over text"),
    "nocrit": ("Bedingung der Zeile fehlt", "Row criterion missing"),
    "partmonth": ("Angefangener Zeitraum verglichen", "Incomplete period compared"),
    "over100": ("Anteil über 100 %", "Share over 100%"),
    "winstart": ("Gleitendes Mittel am Anfang kürzer", "Moving average shorter at the start"),
    "effekte": ("Preis- und Mengeneffekt gehen nicht auf", "Price and volume effects do not add up"),
    "bruecke": ("Brücke ohne einen Treiber", "Bridge leaves out a driver"),
    "negbasis": ("Veränderung gegen eine negative Basis", "Change against a negative base"),
    "vorzeichen": ("Vorzeichen gegen „abzüglich“", "Sign against “less”"),
    "ausgleich": ("Soll und Haben gehen nicht auf", "Debits and credits do not balance"),
    "restwert": ("AfA ohne Restwert", "Depreciation without residual value"),
    "buchwert": ("Buchwert unter null", "Book value below zero"),
    "zinsperiode": ("Jahreszins mit Monaten", "Annual rate with months"),
    "mape": ("Mittlerer Fehler ohne Betrag", "Mean error without absolute value"),
    "nbw0": ("NBW zinst Jahr 0 ab", "NPV discounts year 0"),
    "basis": ("Anteil an der falschen Zeile", "Share of the wrong row"),
    "steuern": ("Steuern addiert statt abgezogen", "Taxes added instead of subtracted"),
    "gordon": ("Terminal Value falsch nach Gordon", "Gordon terminal value wrong"),
    "tvfaktor": ("Terminal Value mit falschem Faktor abgezinst", "Terminal value discounted with the wrong factor"),
    "anfang": ("Endbestand statt Anfangsbestand", "Closing instead of opening balance"),
    "annahme": ("Falsche Annahme gelesen", "Wrong assumption read"),
    "auszahlung": ("Auszahlung als Zufluss addiert", "Outflow added as inflow"),
    "barwert": ("Barwert abgezogen statt addiert", "Present value subtracted instead of added"),
    "ganzzahl": ("Satz als ganze Zahl statt Prozent", "Rate as a whole number instead of percent"),
    "startmonat": ("Startmonat nicht mitgezählt", "Start month not counted"),
    "zscore": ("z-Score ohne Streuung", "z-score without standard deviation"),
    "means": ("Mittelwert von Mittelwerten", "Average of averages"),
    "growth": ("Wachstum auf falscher Basis", "Growth on the wrong base"),
    "window": ("Gleitender Bereich festgehalten", "Moving range fixed"),
    "nenner": ("Geteilt durch die falsche Spalte", "Divided by the wrong column"),
    "unit": ("Einheit im Spaltenkopf passt nicht", "Unit in the column header does not match"),
    "sign": ("Vorzeichen umgekehrt", "Sign reversed"),
    "negcost": ("Negative Kosten abgezogen", "Negative costs subtracted"),
    "rabatt": ("Rabatt zugezählt", "Discount added"),
    "gross": ("Brutto ohne MwSt.", "Gross without VAT"),
    "percent2": ("Prozent doppelt", "Percent twice"),
    "percentfmt": ("Prozent ohne Prozentformat", "Percent without percent format"),
    "timeunit": ("Zeit in falscher Einheit", "Time in the wrong unit"),
    "period": ("Monat und Jahr verwechselt", "Month and year mixed up"),
    "indexbase": ("Index auf falscher Basis", "Index on the wrong base"),
    "span": ("Fenster mit falscher Länge", "Window with the wrong length"),
    "markup": ("Marge durch die Kosten", "Margin divided by costs"),
    "extreme": ("MIN statt MAX", "MIN instead of MAX"),
    "statcol": ("Kennzahl über falsche Spalte", "Figure over the wrong column"),
    "shifted": ("Bezug eine Zeile verschoben", "Reference shifted by one row"),
    "swapped": ("Spalten vertauscht?", "Columns swapped?"),
    "ytd": ("„Bis heute“ mit festem Datum", "“To date” with a fixed date"),
    "currency": ("Währungen ohne Kurs", "Currencies without exchange rate"),
    "title": ("Titel und Kopf widersprechen sich", "Title and header contradict each other"),
    "filter": ("Weggefilterte Zeilen zählen mit", "Filtered-out rows are counted"),
    "ausgeblendet": ("Ausgeblendete Zeilen zählen mit", "Hidden rows are counted"),
    "lookcol": ("Falsche Spalte geholt", "Wrong column returned"),
    "lookshort": ("Suchtabelle hört zu früh auf", "Lookup table ends too early"),
    "listshort": ("Liste wird nicht ganz gelesen", "List not read to the end"),
    "colname": ("Falsche Spalte der Tabelle", "Wrong column of the table"),
    "dupkey": ("Doppelter Schlüssel zählt doppelt", "Duplicate key counted twice"),
    "unitmix": ("Zeilen in verschiedener Einheit", "Rows in different units"),
    "rateshift": ("Satz um Faktor 10 verschoben?", "Rate off by a factor of 10?"),
    "monthend": ("Falsches Monatsende", "Wrong month end"),
    "fiscal": ("Falsches Geschäftsjahr", "Wrong fiscal year"),
    "fquarter": ("Kalender- statt Geschäftsquartal", "Calendar instead of fiscal quarter"),
    "lenient": ("Prüfung mit Toleranz", "Check with tolerance"),
    "pivotfilter": ("Pivot-Filter anders als verlangt", "Pivot filter differs from what is asked"),
    "lookexact": ("Ungefähre statt genaue Suche", "Approximate instead of exact lookup"),
    "lookslide": ("Suchtabelle rutscht mit", "Lookup table shifts"),
    "looknear": ("Suchwert wird nicht gefunden", "Lookup value not found"),
    "parts": ("Summe passt nicht zu ihren Teilen", "Sum does not match its parts"),
    "blocks": ("Block lässt eine Zeile aus", "Block leaves out a row"),
    "clipped": ("Zahl im Uhrzeitformat", "Number in time format"),
    "zeroshown": ("Abweichung, die das Format als 0 zeigt", "Difference the format shows as 0"),
    "groups": ("Teile erfassen nicht alle Zeilen", "Parts do not cover all rows"),
    "statcode": ("Code für „fehlt“ zählt mit", "Code for “missing” is counted"),
    "statzero": ("Nullen aus WENNFEHLER zählen mit", "Zeros from IFERROR are counted"),
    "statblank": ("Leere Zellen zählen als 0", "Empty cells count as 0"),
    "spread": ("Falsche Streuung", "Wrong spread"),
    "spreadmix": ("Streuung anders als die Nachbarn", "Spread differs from its neighbors"),
    "base": ("Anteil von der falschen Summe", "Share of the wrong total"),
    "ttest": ("T.TEST falsch aufgesetzt", "T.TEST set up wrong"),
    "confidence": ("KONFIDENZ falsch aufgesetzt", "CONFIDENCE set up wrong"),
    "interval": ("Konfidenzintervall falsch", "Confidence interval wrong"),
    "points": ("Prozentpunkte und Prozent verwechselt", "Percentage points and percent mixed up"),
    "age": ("Alter von Hand gerechnet", "Age calculated by hand"),
    "months": ("Monate von Hand gerechnet", "Months calculated by hand"),
    "today": ("HEUTE() in einer Fall-Liste", "TODAY() in a list of cases"),
    "days": ("Tage falsch gezählt", "Days counted wrong"),
    "pivot": ("Pivot-Tabelle veraltet", "Pivot table out of date"),
    "pivotsource": ("Pivot-Tabelle nimmt nicht alle Zeilen", "Pivot table misses rows"),
    "pivotcount": ("Pivot-Tabelle zählt statt zu summieren", "Pivot table counts instead of summing"),
    "chart": ("Diagramm zeigt nicht alles", "Chart does not show everything"),
    "chartname": ("Diagrammreihe falsch benannt", "Chart series named wrong"),
    "line": ("Summe lässt ein Jahr aus", "Sum leaves out a year"),
    "year": ("Falsches Jahr in der Rechnung", "Wrong year in the calculation"),
    "walk": ("Bezug auf das falsche Jahr", "Reference to the wrong year"),
    "drift": ("Annahme rutscht beim Kopieren mit", "Assumption shifts when copied"),
    "npv": ("Barwert falsch abgezinst", "Present value discounted wrong"),
    "frozen": ("Annahme mit $ festgehalten", "Assumption fixed with $"),
    "sensitivity": ("Sensitivität passt nicht zum Modell", "Sensitivity does not match the model"),
    "sensaxis": ("Achse der Sensitivität falsch", "Sensitivity axis wrong"),
    "hundred": ("Anteile ergeben nicht 100 %", "Shares do not add up to 100%"),
    "discount": ("Abgezinst mit der falschen Annahme", "Discounted with the wrong assumption"),
    "selfcheck": ("Eigene Prüfung zeigt einen Fehler", "Own check shows an error"),
    "spelling": ("Anders geschrieben", "Spelled differently"),
    "typo": ("Wohl ein Tippfehler", "Probably a typo"),
    "mapping": ("Passt nicht zum Rest der Zeile", "Does not match the rest of the row"),
    "pair": ("Zwei Angaben widersprechen sich", "Two entries contradict each other"),
    "merged": ("Verbundene Zellen in der Liste", "Merged cells in the list"),
}

_CELL = r"[A-Z]{1,3}[1-9]\d*"
# Ein Wert im Text: eine Zahl (auch mit %), ein Text in Anführungszeichen oder ein Fehlerwert.
# Englisch in “…” (die Zahlen schreibt sprache.output erst zum Schluss um, hier stehen sie noch deutsch).
_VALUE = r"(?:-?\d[\d.,]*(?: ?%)?|„[^“]*“|“[^”]*”|#[A-ZÄÖÜ/0!?]+)"
# Zwei Formeln: „…: =E25 - E23 statt =E25 - E24.“, englisch „…: =E25 - E23 instead of =E25 - E24.“ Die Formeln enden an
# einem Punkt mit Leerzeichen dahinter.
_FORMULAS = re.compile(r"(?P<pre>:\s*|\s)(?P<now>=\S.*?) (?:statt|instead of) (?P<want>=\S.*?)(?=\.(?:\s|$))")
# Je Muster erst Deutsch, dann Englisch; die deutschen Muster stehen vorne, damit deutsche Sätze wie bisher gelesen werden.
_SHOWS = [
    # „D29 zeigt darum 418,97 statt 578,57.“ (auch „(so viel wie D28)“ dazwischen)
    re.compile(rf"(?P<cell>{_CELL}) zeigt darum (?P<now>{_VALUE})(?: \(so viel wie {_CELL}\))? statt (?P<want>{_VALUE})\."),
    # „So ergibt C10 32487758,22 statt 36997349,84.“, „Darum zeigt C7 0 % statt 0,2.“
    re.compile(rf"(?:So|Darum) (?:ergibt|zeigt) (?P<cell>{_CELL}) (?P<now>{_VALUE}) statt (?P<want>{_VALUE})\."),
    # „Richtig wäre 713236,38 statt 686803,03.“
    re.compile(rf"Richtig wäre (?P<want>{_VALUE}) statt (?P<now>{_VALUE})\."),
    # “So D29 shows 418,97 instead of 578,57.” (also “(the same as D28)” in between; “D29 therefore shows …” too)
    re.compile(rf"(?:(?:So|That is why|Therefore|This is why|As a result,) (?P<cell>{_CELL})|(?P<cell2>{_CELL}) therefore) "
               rf"(?:shows|gives|returns) (?P<now>{_VALUE})(?: \((?:the same|as much) as {_CELL}\))? instead of (?P<want>{_VALUE})\."),
    # “Should be 713236,38 instead of 686803,03.”, auch „It should be“, „In C5, it should be“, „The correct value would be“
    re.compile(rf"(?:(?:It s|S)hould be|(?:The c|C)orrect (?:value )?would be) (?P<want>{_VALUE}) "
               rf"instead of (?P<now>{_VALUE})\."),
]


def view(hint: dict) -> dict:
    """{"title": …, "where": …, "lead": …, "compare": {…}} für das Seitenfenster; leer, was nicht passt."""
    kind = hint["key"].partition(":")[0]
    text = hint.get("text") or ""
    cell = (hint.get("cell") or "").split(",")[0].strip()
    title = TITLES.get(kind)
    out: dict = {"title": T(*title) if title else kind_name(kind)}
    # Zwei Zellen („F40,G40“): Der Satz nennt beide, nichts davon gehört nach oben.
    where, lead = _where(text, cell) if "," not in (hint.get("cell") or "") else ("", text)
    compare = None
    match = _FORMULAS.search(lead)
    if match:
        now, want = _tidy(match.group("now")), _tidy(match.group("want"))
        compare = {"kind": "formel", "now": now, "want": want,
                   "now_label": f"In {cell}" if cell else T("Jetzt", "Now"),
                   "want_label": T("Wie die übrigen", "Like the others") if kind == "pattern" else T("Richtig", "Correct")}
        lead = lead[:match.start()] + "." + lead[match.end() + 1:]
    shown = None
    for way in _SHOWS:
        shown = way.search(lead)
        if shown:
            break
    if shown:
        at = shown.groupdict().get("cell") or shown.groupdict().get("cell2") or cell
        values = {"now": shown.group("now"), "want": shown.group("want"),
                  "now_label": (T(f"{at} zeigt", f"{at} shows") if at else T("Jetzt", "Now")),
                  "want_label": T("Richtig wäre", "Should be")}
        lead = (lead[:shown.start()] + lead[shown.end():]).strip()
        if compare is None:
            compare = {"kind": "wert", **values}
        else:
            compare["values"] = values
    out["lead"] = re.sub(r"\s{2,}", " ", lead).replace(" .", ".").strip()
    if where:
        out["where"] = where
    if compare:
        out["compare"] = compare
    return out


def _where(text: str, cell: str) -> tuple[str, str]:
    """„C23 (B, Jahr 2) rechnet …“ → („B · Jahr 2“, „Rechnet …“): Die Zelle steht schon vorne im Fenster."""
    if not cell or not text.startswith(f"{cell} (") or ":" in cell:
        return "", text
    depth, end = 0, None
    for i in range(len(cell) + 1, len(text)):
        if text[i] == "(":
            depth += 1
        elif text[i] == ")":
            depth -= 1
            if depth == 0:
                end = i
                break
    if end is None or end + 1 >= len(text) or text[end + 1] != " ":
        return "", text
    inside = text[len(cell) + 2:end]
    rest = text[end + 2:]
    if not rest or not (rest[0].islower() or rest[0] in "„“"):
        return "", text
    parts = _split(inside)
    return " · ".join(parts), rest[0].upper() + rest[1:]


def _split(inside: str) -> list[str]:
    """„Neu-Ausfallanteil, Jahr 4“ → zwei Teile; Kommas in Klammern („(TEUR, real)“) trennen nicht."""
    parts, depth, start = [], 0, 0
    for i, ch in enumerate(inside):
        depth += ch == "("
        depth -= ch == ")"
        if ch == "," and depth == 0 and inside[i + 1:i + 2] == " ":
            parts.append(inside[start:i].strip())
            start = i + 1
    parts.append(inside[start:].strip())
    return [p for p in parts if p]


def _tidy(formula: str) -> str:
    """Wie in der Bearbeitungsleiste: =E25-E23 statt =E25 - E23. Mit Namen („=Umsatz 2026 * Marge“) bleiben die
    Leerzeichen, sonst liest man die Namen nicht mehr."""
    if re.search(r"[a-zäöüß]", re.sub(r'"[^"]*"|„[^“]*“|“[^”]*”', "", formula)):
        return formula
    return re.sub(r"\s*(<=|>=|<>|[-+*/^&<>;=])\s*", r"\1", formula)
