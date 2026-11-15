"""Prüfbericht zum Weitergeben: alle Hinweise der Mappe als eigenständige HTML-Seite, zum Drucken oder als PDF.

Zwei Arten: mit den Texten der Hinweise (darin stehen Werte und Namen aus der Mappe) oder ohne Daten. Ohne Daten
stehen nur die Art des Hinweises, seine Stufe und die Zelle da, die Blätter heißen „Blatt 1“, „Blatt 2“, und
der Name der Mappe fehlt. So lässt sich der Bericht weitergeben, auch wenn in der Mappe Patientendaten stehen.
"""

from __future__ import annotations

import re
import secrets
import time
from datetime import datetime
from html import escape
from pathlib import Path

from . import __version__
from .sprache import T, english, formula, lang, sentence

# Die Art eines Hinweises (Schlüssel vor dem Doppelpunkt) in Worten, ohne Werte aus der Mappe: (deutsch, englisch).
KINDS = {
    "error": ("Fehlerwert in einer Zelle", "Error value in a cell"),
    "gap": ("Summe lässt Zeilen aus", "Sum leaves out rows"),
    "sum": ("Summe ergibt nachgerechnet etwas anderes", "Sum recalculates to a different result"),
    "round": ("Gerundete Werte gehen zusammen nicht auf", "Rounded values do not add up"),
    "typed": ("Feste Zahl statt Formel", "Fixed number instead of a formula"),
    "constant": ("Dieselbe feste Zahl in mehreren Formeln", "Same fixed number in several formulas"),
    "copied": ("Abgetippte Zahl statt Zellbezug", "Typed-in number instead of a cell reference"),
    "pattern": ("Formel fällt aus dem Muster ihrer Spalte", "Formula breaks the pattern of its column"),
    "circle": ("Kreisbezug", "Circular reference"),
    "double": ("Zwischensumme doppelt gezählt", "Subtotal counted twice"),
    "recalc": ("Wert in Excel weicht vom nachgerechneten Ergebnis ab", "Value in Excel differs from the recalculated result"),
    "emptyref": ("Formel liest eine leere Zelle", "Formula reads an empty cell"),
    "empty": ("Formelspalte mit leeren Zeilen", "Formula column with empty rows"),
    "complex": ("Schwer prüfbare Formel", "Formula that is hard to check"),
    "approx": ("Ungefähre Suche in unsortierter Tabelle", "Approximate lookup in an unsorted table"),
    "nearest": ("Suche ohne FALSCH nimmt still einen anderen Eintrag", "Lookup without FALSE silently takes a different entry"),
    "hidden": ("WENNFEHLER überdeckt nicht gefundene Suchwerte", "IFERROR hides lookup values that were not found"),
    "twice": ("Suchschlüssel steht doppelt in der Tabelle", "Lookup key appears twice in the table"),
    "sliding": ("Suchtabelle rutscht beim Kopieren mit", "Lookup table shifts when copied"),
    "outlier": ("Wert weit weg von den übrigen", "Value far away from the others"),
    "scale": ("Zahl um Faktor 10, 100 oder 1000 daneben", "Number off by a factor of 10, 100 or 1000"),
    "rowunit": ("Wert in einer anderen Einheit als seine Gruppe", "Value in a different unit than its group"),
    "formeltext": ("Formel als Text: rechnet nicht", "Formula stored as text: it does not calculate"),
    "extern": ("Werte aus einer anderen Datei, nur der letzte Stand geprüft", "Values from another file, only the last saved state checked"),
    "copylink": ("Kopiertes Blatt: Bezug zeigt noch auf die alte Datei", "Copied sheet: reference still points to the old file"),
    "mixed": ("Anteile gemischt als Bruch und in Prozent", "Shares mixed as fractions and percentages"),
    "share": ("Anteile ergeben nicht 100 %", "Shares do not add up to 100%"),
    "text": ("Text in einer Zahlenspalte", "Text in a number column"),
    "dupe": ("Name steht doppelt", "Name appears twice"),
    "missing": ("Fehlende Werte", "Missing values"),
    "placeholder": ("Platzhalter statt Zahl", "Placeholder instead of a number"),
    "code": ("Code für „fehlt“ wird mitgerechnet", "Code for “missing” is counted"),
    "impossible": ("Unmöglicher Wert", "Impossible value"),
    "regel": ("Eigene Regel für eine Spalte verletzt", "Your own rule for a column is broken"),
    "order": ("Ende vor dem Anfang", "End before the start"),
    "future": ("Datum in der Zukunft", "Date in the future"),
    "id": ("Nummer steht mehrmals", "Number appears more than once"),
    "same": ("Zeile doppelt", "Duplicate row"),
    "units": ("Zwei Einheiten in einer Spalte", "Two units in one column"),
    "rate": ("Ein Satz weicht von den übrigen ab", "One rate differs from the others"),
    "range": ("Bereich lässt Zeilen aus", "Range leaves out rows"),
    "group": ("Bereich passt nicht zur Gruppe", "Range does not match the group"),
    "label": ("Bereich oder Zahl passt nicht zum Namen der Kennzahl", "Range or number does not match the name of the figure"),
    "shift": ("Bereiche von SUMMEWENN gegeneinander verschoben", "SUMIF ranges shifted against each other"),
    "partial": ("SUMMEWENN prüft nur einen Teil der Tabelle", "SUMIF checks only part of the table"),
    "unfilled": ("Formeln fehlen in neuen Zeilen", "Formulas missing in new rows"),
    "criterion": ("Bedingung findet nicht, was in der Tabelle steht", "Criterion does not find what is in the table"),
    "count": ("Summe geteilt durch eine feste Anzahl, die nicht passt", "Sum divided by a fixed count that does not fit"),
    "total": ("Bereich nimmt die Summenzeile mit", "Range includes the total row"),
    "slide": ("Bezug rutscht beim Herunterziehen mit ($ fehlt)", "Reference shifts when filled down ($ missing)"),
    "astext": ("Zahl oder Datum als Text, Excel rechnet nicht damit", "Number or date stored as text, Excel does not calculate with it"),
    "ratio": ("Mittelwert von Quoten statt Quote über alle Zeilen", "Average of ratios instead of the ratio over all rows"),
    "quantil": ("Quartil oder Perzentil mit einer anderen Nummer, als der Name sagt", "Quartile or percentile with a different number than the name says"),
    "bonf": ("Bonferroni im Namen, die Schwelle aber nicht durch die Zahl der Tests geteilt", "Bonferroni in the name, but the threshold is not divided by the number of tests"),
    "cagr": ("Jährliche Wachstumsrate mit einem Jahr zu viel oder zu wenig", "Annual growth rate with one year too many or too few"),
    "wmean": ("Gewichtetes Mittel, das durch die Zahl der Zeilen statt durch die Gewichte teilt", "Weighted average that divides by the number of rows instead of the weights"),
    "avgmeans": ("Mittel von Gruppenmitteln, obwohl die Gruppen verschieden groß sind", "Average of group averages although the groups differ in size"),
    "corrshift": ("Zwei Bereiche in KORREL & Co., die gegeneinander verschoben sind", "Two ranges in CORREL & co. shifted against each other"),
    "counttext": ("ANZAHL über Texte, die immer 0 ergibt", "COUNT over text, which always gives 0"),
    "nocrit": ("ZÄHLENWENNS ohne die Bedingung der Zeile, in jeder Zeile dasselbe", "COUNTIFS without the row's criterion, the same in every row"),
    "partmonth": ("Angefangener letzter Zeitraum im Vergleich mit ganzen", "Incomplete last period compared with full ones"),
    "over100": ("Anteil über 100 %, Zähler größer als Nenner", "Share over 100%, numerator larger than denominator"),
    "winstart": ("Gleitendes Mittel, das am Anfang weniger Werte nimmt", "Moving average that takes fewer values at the start"),
    "effekte": ("Preis- und Mengeneffekt, die zusammen nicht die Abweichung ergeben", "Price and volume effects that together do not give the variance"),
    "bruecke": ("Brücke, die einen Treiber direkt über dem Ergebnis auslässt", "Bridge that leaves out a driver directly above the result"),
    "negbasis": ("Veränderung in %, geteilt durch eine negative Basis ohne ABS", "Change in % divided by a negative base without ABS"),
    "vorzeichen": ("Zeile „abzüglich“ addiert oder „zuzüglich“ abgezogen", "Row “less” added or “plus” subtracted"),
    "ausgleich": ("Soll und Haben oder Aktiva und Passiva, die nicht gleich sind", "Debits and credits or assets and liabilities that are not equal"),
    "restwert": ("Abschreibung über die ganzen Anschaffungskosten, obwohl ein Restwert dasteht", "Depreciation over the full acquisition cost although a residual value is given"),
    "buchwert": ("Buchwert unter null, die Abschreibung läuft zu lange", "Book value below zero, depreciation runs too long"),
    "zinsperiode": ("Jahreszins mit einer Laufzeit in Monaten", "Annual interest rate with a term in months"),
    "mape": ("Mittlerer absoluter Fehler mit Vorzeichen gemittelt", "Mean absolute error averaged with signs"),
    "nbw0": ("NBW, das auch die Zahlung in Jahr 0 abzinst", "NPV that also discounts the payment in year 0"),
    "basis": ("Anteil an einer Zeile („% vom Umsatz“), deren Bezug mitgerutscht ist", "Share of a row (“% of revenue”) whose reference has shifted"),
    "steuern": ("Steuern zum Ergebnis vor Steuern addiert statt abgezogen", "Taxes added to earnings before taxes instead of subtracted"),
    "gordon": ("Terminal Value nach Gordon mit dem Cashflow des letzten Jahres statt des Folgejahrs", "Gordon terminal value with the last year's cash flow instead of the next year's"),
    "tvfaktor": ("Terminal Value mit dem Diskontierungsfaktor eines anderen Jahres abgezinst", "Terminal value discounted with another year's discount factor"),
    "anfang": ("Rechnet mit dem Endbestand, obwohl der Name den Anfangsbestand nennt", "Uses the closing balance although the name says opening balance"),
    "annahme": ("Liest die Annahme mit dem anderen Namen", "Reads the assumption with the other name"),
    "auszahlung": ("Auszahlung (Investition, Tilgung, Kosten) positiv addiert, wo die übrigen negativ stehen", "Outflow (investment, repayment, costs) added as positive where the others are negative"),
    "barwert": ("Barwert im Kapitalwert abgezogen statt addiert", "Present value subtracted instead of added in the net present value"),
    "ganzzahl": ("Zinssatz oder Quote als ganze Zahl (9 statt 9 %)", "Interest rate or ratio as a whole number (9 instead of 9%)"),
    "startmonat": ("Monate ab Startmonat ohne den Startmonat (12 statt 13 minus Monat)", "Months from the start month without the start month (12 instead of 13 minus month)"),
    "zscore": ("z-Score, der durch den Mittelwert statt durch die Streuung teilt", "z-score that divides by the mean instead of the standard deviation"),
    "means": ("Mittelwert von Mittelwerten verschieden großer Gruppen", "Average of averages of groups of different sizes"),
    "growth": ("Wachstum auf der falschen Basis", "Growth on the wrong base"),
    "window": ("Gleitender Mittelwert mit $ festgehalten", "Moving average fixed with $"),
    "nenner": ("Geteilt durch eine andere Spalte, als der Name sagt", "Divided by a different column than the name says"),
    "unit": ("Tausend und Einer ohne Umrechnung zusammengezählt", "Thousands and units added without conversion"),
    "sign": ("Vorzeichen umgekehrt gegen den Namen (Plan − Ist statt Ist − Plan)", "Sign reversed against the name (plan − actual instead of actual − plan)"),
    "negcost": ("Negativ erfasste Kosten noch einmal abgezogen", "Costs recorded as negative subtracted again"),
    "rabatt": ("Rabatt zugezählt statt abgezogen", "Discount added instead of subtracted"),
    "gross": ("Brutto oder inkl. MwSt., aber ohne MwSt. gerechnet", "Gross or incl. VAT, but calculated without VAT"),
    "percent2": ("Mal 100 und Prozentformat: Prozent doppelt", "Times 100 and percent format: percent twice"),
    "percentfmt": ("Anteil ohne Prozentformat", "Share without percent format"),
    "timeunit": ("Zeitdifferenz in der falschen Einheit", "Time difference in the wrong unit"),
    "period": ("Pro Monat und pro Jahr verwechselt", "Per month and per year mixed up"),
    "indexbase": ("Index teilt durch eine andere Basis, als der Name sagt", "Index divides by a different base than the name says"),
    "span": ("Gleitendes Fenster mit einer anderen Länge, als der Name sagt", "Moving window with a different length than the name says"),
    "markup": ("Marge durch die Kosten statt durch den Umsatz", "Margin divided by costs instead of revenue"),
    "extreme": ("MIN statt MAX oder umgekehrt", "MIN instead of MAX or the other way round"),
    "statcol": ("Kennzahl rechnet über eine andere Spalte, als der Name sagt", "Figure calculates over a different column than the name says"),
    "shifted": ("Bezug aus der Zeile darüber", "Reference from the row above"),
    "swapped": ("Umsatz und Kosten vielleicht vertauscht", "Revenue and costs possibly swapped"),
    "ytd": ("„Bis heute“ mit einem festen Datum", "“To date” with a fixed date"),
    "currency": ("Zwei Währungen ohne Wechselkurs zusammengezählt", "Two currencies added without an exchange rate"),
    "title": ("Satz über der Tabelle und Spaltenkopf nennen verschiedene Einheiten", "Text above the table and column header name different units"),
    "filter": ("Summe über gefilterte Tabelle zählt ausgeblendete Zeilen", "Sum over a filtered table counts hidden rows"),
    "ausgeblendet": ("Summe zählt ausgeblendete Zeilen mit", "Sum counts hidden rows"),
    "lookcol": ("SVERWEIS holt die falsche Spalte", "VLOOKUP returns the wrong column"),
    "lookshort": ("Suchtabelle hört zu früh auf", "Lookup table ends too early"),
    "listshort": ("Eine Formel liest ihre Liste nicht bis zum letzten Eintrag", "A formula does not read its list to the last entry"),
    "colname": ("Eine Kennzahl fasst eine andere Spalte zusammen, als ihr Name sagt", "A figure summarizes a different column than its name says"),
    "dupkey": ("Ein doppelter Schlüssel bekommt seine Summe zweimal", "A duplicate key gets its sum twice"),
    "unitmix": ("Eine Summe zählt Zeilen in verschiedener Einheit zusammen", "A sum adds up rows in different units"),
    "rateshift": ("Ein Satz ist genau ein Zehntel oder das Zehnfache der übrigen Jahre", "A rate is exactly a tenth or ten times that of the other years"),
    "monthend": ("Unter „Monatsende“ steht das Ende eines anderen Monats", "Under “month end” is the end of another month"),
    "fiscal": ("Die Formel teilt die Monate einem anderen Geschäftsjahr zu, als das Blatt sagt", "The formula assigns the months to a different fiscal year than the sheet says"),
    "fquarter": ("Unter einem Quartal des Geschäftsjahrs steht das Kalenderquartal", "Under a fiscal quarter is the calendar quarter"),
    "lenient": ("Eine Prüfung ohne Toleranz lässt eine Abweichung durch", "A check without tolerance lets a difference through"),
    "pivotfilter": ("Der Berichtsfilter einer Pivot-Tabelle steht anders, als darüber verlangt", "The report filter of a pivot table differs from what is asked above it"),
    "lookexact": ("Ungefähre Suche, wo genau gesucht werden muss", "Approximate lookup where an exact match is needed"),
    "lookslide": ("Suchtabelle ohne $ rutscht beim Herunterziehen mit", "Lookup table without $ shifts when filled down"),
    "looknear": ("Suchwert fehlt wegen Leerzeichen oder Zahl als Text", "Lookup value missing because of spaces or a number stored as text"),
    "parts": ("Summe passt nicht zu ihren Teilen", "Sum does not match its parts"),
    "blocks": ("Block lässt eine Zeile aus oder nimmt sie doppelt", "Block leaves out a row or takes it twice"),
    "clipped": ("Zahl im Uhrzeitformat", "Number in time format"),
    "zeroshown": ("Abweichung, die das Format als 0 zeigt", "Difference that the format shows as 0"),
    "groups": ("Teile erfassen nicht alle Zeilen der Tabelle", "Parts do not cover all rows of the table"),
    "statcode": ("Statistik rechnet Codes für „fehlt“ (999) mit", "Statistic includes codes for “missing” (999)"),
    "statzero": ("Statistik zählt Nullen aus WENNFEHLER als Werte", "Statistic counts zeros from IFERROR as values"),
    "statblank": ("Mittelwert zählt leere Zellen oder Text als 0", "Average counts empty cells or text as 0"),
    "spread": ("Falsche Streuung (STABW.N/STABW.S, Varianz, Standardfehler)", "Wrong spread (STDEV.P/STDEV.S, variance, standard error)"),
    "spreadmix": ("Standardabweichung rechnet anders als ihre Nachbarn", "Standard deviation calculated differently from its neighbors"),
    "base": ("Anteil von der falschen Summe", "Share of the wrong total"),
    "ttest": ("T.TEST falsch aufgesetzt (gepaart, Varianzen, Seiten)", "T.TEST set up wrong (paired, variances, tails)"),
    "confidence": ("KONFIDENZ falsch aufgesetzt", "CONFIDENCE set up wrong"),
    "interval": ("Konfidenzintervall von Hand falsch gerechnet", "Confidence interval calculated wrong by hand"),
    "points": ("Prozentpunkte und Prozent verwechselt", "Percentage points and percent mixed up"),
    "age": ("Alter von Hand gerechnet, um den Geburtstag ein Jahr daneben", "Age calculated by hand, off by a year around the birthday"),
    "months": ("Ganze Monate von Hand gerechnet", "Whole months calculated by hand"),
    "today": ("HEUTE() in einer Tabelle von Fällen", "TODAY() in a table of cases"),
    "days": ("Tage mit oder ohne den ersten Tag verwechselt", "Days counted with or without the first day mixed up"),
    "pivot": ("Pivot-Tabelle nicht aktualisiert", "Pivot table not refreshed"),
    "pivotsource": ("Pivot-Tabelle nimmt nicht alle Zeilen der Daten", "Pivot table does not take all rows of the data"),
    "pivotcount": ("Pivot-Tabelle zählt statt zu summieren", "Pivot table counts instead of summing"),
    "chart": ("Diagramm zeigt nicht die ganze Tabelle", "Chart does not show the whole table"),
    "chartname": ("Diagrammreihe heißt nach einer anderen Spalte", "Chart series named after a different column"),
    "typedcell": ("Feste Zahl, wo die Spalte sonst mit Formel rechnet", "Fixed number where the column otherwise uses a formula"),
    "line": ("Summe einer Zeile lässt ein Jahr aus", "Sum of a row leaves out a year"),
    "year": ("Rechnung nimmt ein falsches Jahr", "Calculation takes a wrong year"),
    "walk": ("Formel liest das Jahr davor oder danach", "Formula reads the year before or after"),
    "drift": ("Bezug auf eine Annahme ohne $ rutscht beim Kopieren mit", "Reference to an assumption without $ shifts when copied"),
    "npv": ("Barwert zinst das erste Jahr falsch ab", "Present value discounts the first year wrong"),
    "frozen": ("Annahme mit $ für alle Jahre festgehalten", "Assumption fixed with $ for all years"),
    "sensitivity": ("Sensitivitätstabelle passt nicht zum Modell", "Sensitivity table does not match the model"),
    "sensaxis": ("Achse der Sensitivitätstabelle falsch", "Axis of the sensitivity table is wrong"),
    "hundred": ("Anteile einer Zeile ergeben nicht 100 %", "Shares of a row do not add up to 100%"),
    "discount": ("Abgezinst mit einer Annahme, die kein Zinssatz ist", "Discounted with an assumption that is not an interest rate"),
    "selfcheck": ("Die eigene Prüfung des Modells zeigt einen Fehler", "The model's own check shows an error"),
    "spelling": ("Derselbe Name anders geschrieben", "Same name spelled differently"),
    "typo": ("Wohl ein Tippfehler in einem Namen", "Probably a typo in a name"),
    "mapping": ("Ein Wert passt nicht zum Rest der Zeile", "A value does not match the rest of the row"),
    "pair": ("Zwei Angaben einer Zeile widersprechen sich", "Two entries in a row contradict each other"),
    "merged": ("Verbundene Zellen: Formeln sehen nur die erste Zeile", "Merged cells: formulas see only the first row"),
}
LEVELS = {"fehler": ("Fehler", "Error"), "pruefen": ("Prüfen", "Review"), "info": ("Hinweis", "Note")}
KEEP = 20
MAX_AGE = 24 * 3600  # Sekunden: so lange liegt ein Prüfbericht im Datenordner
NAME = re.compile(r"^[0-9]{8}-[0-9]{6}-[0-9a-f]{8}-pruefbericht(-ohne-daten)?\.html$")


def kind_name(kind: str, default: str = "") -> str:
    """Die Art eines Hinweises in Worten, in der Sprache der Anfrage."""
    pair = KINDS.get(kind)
    return T(*pair) if pair else default


def _level(level: str) -> str:
    return T(*LEVELS.get(level, LEVELS["info"]))


def _plural(n: int, one: str, many: str) -> str:
    return one if n == 1 else many


def render(book: str, sheets: list[tuple[str, dict]], with_data: bool, when: datetime | None = None,
           mappe: dict | None = None, rules: list[dict] | None = None) -> str:
    """sheets: je Blatt (Name, Ergebnis wie pruefer.result). mappe: die Mappe in Worten (inworten.describe), nur mit
    Daten. rules: die eigenen Regeln, die in der Mappe gelten (spaltenregeln.used); ohne Daten nur ihre Zahl, denn
    darin stehen Spaltennamen und Werte. Gibt die ganze HTML-Seite zurück."""
    when = when or datetime.now()
    count = {level: sum(1 for _, r in sheets for h in r.get("hints", []) if h["level"] == level) for level in LEVELS}
    report = T("Prüfbericht", "Check report")
    title = f"{report} {escape(book)}" if with_data else report
    stamp = when.strftime(T("%d.%m.%Y, %H:%M", "%Y-%m-%d, %H:%M"))
    notes, n_sheets = count["info"], len(sheets)
    parts = [
        f"<h1>{title}</h1>",
        T(f"<p class=meta>{stamp} · Gegenprobe Prüfer {escape(__version__)} · ohne KI, auf diesem Rechner nachgerechnet"
          f"{'' if with_data else ' · ohne Daten: nur Art, Stufe und Zelle, Blätter nummeriert'}</p>",
          f"<p class=meta>{stamp} · Gegenprobe checker {escape(__version__)} · no AI, recalculated on this computer"
          f"{'' if with_data else ' · without data: only type, level and cell, sheets numbered'}</p>"),
        T(f"<p class=sum><b>{count['fehler']}</b> Fehler · <b>{count['pruefen']}</b> zum Prüfen · <b>{notes}</b> {_plural(notes, 'Hinweis', 'Hinweise')}"
          f" · {n_sheets} {_plural(n_sheets, 'Blatt', 'Blätter')}</p>",
          f"<p class=sum><b>{count['fehler']}</b> {_plural(count['fehler'], 'error', 'errors')} · <b>{count['pruefen']}</b> to review · "
          f"<b>{notes}</b> {_plural(notes, 'note', 'notes')} · {n_sheets} {_plural(n_sheets, 'sheet', 'sheets')}</p>"),
    ]
    if rules:
        parts.append(_rules(rules, with_data))
    words = {}
    if with_data and mappe and mappe.get("state") == "ok":
        parts.append(_mappe(mappe))
        words = {d["name"]: d for d in mappe["sheets"]}
    for i, (name, found) in enumerate(sheets, start=1):
        parts.append(_sheet(name if with_data else T(f"Blatt {i}", f"Sheet {i}"), found, with_data, words.get(name)))
    parts.append(T(
        "<p class=foot>Der Prüfer rechnet jede Formel mit den Werten aus Excel nach und sucht typische Fehler in Formeln und Daten. "
        "Er sieht nicht, ob eine Formel fachlich das Richtige rechnet, und nichts außerhalb der Mappe (andere Dateien, Makros). "
        "„Prüfen“ heißt: Das sieht ungewöhnlich aus, kann aber richtig sein.</p>",
        "<p class=foot>The checker recalculates every formula with the values from Excel and looks for typical mistakes in "
        "formulas and data. It cannot see whether a formula calculates the right thing for your subject, nor anything outside "
        "the workbook (other files, macros). “Review” means: this looks unusual but may be correct.</p>"
    ))
    return _page(title, "\n".join(parts))


def _rules(rules: list[dict], with_data: bool) -> str:
    """Mit welchen eigenen Regeln je Spalte geprüft wurde, damit man den Bericht nachvollziehen kann."""
    if not with_data:
        n = len(rules)
        return T(f"<p>Geprüft auch mit {n} {_plural(n, 'eigenen Regel', 'eigenen Regeln')} je Spalte.</p>",
                 f"<p>Also checked with {n} {_plural(n, 'rule', 'rules')} of your own per column.</p>")
    items = "".join(f"<li><b>{escape(r['spalte'])}</b> ({escape(', '.join(r['blaetter']))}): {escape(sentence(r['text']))}</li>"
                    for r in rules)
    return f"<section><h2>{T('Eigene Regeln je Spalte', 'Your own rules per column')}</h2><ul>{items}</ul></section>"


def _mappe(mappe: dict) -> str:
    """Was die ganze Mappe rechnet: die ersten Sätze und die Ergebnisse mit dem, woraus sie entstehen."""
    lines = "".join(f"<p>{escape(sentence(line))}</p>" for line in mappe["summary"])
    rows = []
    for r in mappe["results"]:
        name = f" {escape(r['name'])}" if r["name"] else ""
        check = {"anders": T(" <span class=bad>nachgerechnet anders</span>", " <span class=bad>recalculates differently</span>"),
                 "unbekannt": T(" (nicht nachgerechnet)", " (not recalculated)")}.get(r["check"], "")
        rows.append(f"<tr><td class=cell>{escape(r['sheet'])}!{escape(r['cell'])}</td><td>{name.strip()}</td>"
                    f"<td class=num>{escape(sentence(r['value'] or ''))}</td><td>{escape(sentence(' '.join(r['text'])))}{check}</td></tr>")
    table = (T("<table><thead><tr><th>Zelle</th><th>Ergebnis</th><th>Wert</th><th>Woraus</th></tr></thead><tbody>",
               "<table><thead><tr><th>Cell</th><th>Result</th><th>Value</th><th>Made from</th></tr></thead><tbody>")
             + "".join(rows) + "</tbody></table>") if rows else ""
    return f"<section><h2>{T('Was die Mappe rechnet', 'What the workbook calculates')}</h2>{lines}{table}</section>"


def _words(described: dict) -> str:
    """Was ein Blatt rechnet: die gerechneten Spalten und die einzelnen Formeln, mit ihrer Formel in Worten."""
    items = []
    for c in described["columns"]:
        words = f" = {escape(sentence(c['words']))}" if c["words"] and c["words"] != c["formula"] else ""
        items.append(f"<li><b>{escape(c['name'])}</b> <span class=cell>{escape(c['area'])}</span>: "
                     f"<span class=cell>{escape(formula(c['formula']))}</span>{words}. {escape(sentence(' '.join(c['text'])))}</li>")
    for f in described["figures"]:
        name = f"<b>{escape(f['name'])}</b> " if f["name"] else ""
        words = f" {escape(sentence(f['words']))}" if f["words"] else ""
        bad = T(" <span class=bad>Nachgerechnet anders.</span>", " <span class=bad>Recalculates differently.</span>") \
            if f["check"] == "anders" else ""
        items.append(f"<li>{name}<span class=cell>{escape(f['cell'])}</span> = {escape(sentence(f['value'] or ''))}: "
                     f"<span class=cell>{escape(formula(f['formula']))}</span>{words}{bad}</li>")
    if described.get("figures_more"):
        more = described["figures_more"]
        items.append(T("<li>und eine weitere Formel</li>", "<li>and one more formula</li>") if more == 1
                     else T(f"<li>und {more} weitere Formeln</li>", f"<li>and {more} more formulas</li>"))
    for p in described["parts"]:
        items.append(f"<li>{escape(sentence(p['text']))}</li>")
    return f"<h3>{T('Was dieses Blatt rechnet', 'What this sheet calculates')}</h3><ul>{''.join(items)}</ul>" if items else ""


def _sheet(name: str, found: dict, with_data: bool, described: dict | None = None) -> str:
    stats = found.get("stats") or {}
    formulas = stats.get("formeln", 0)
    unknown = stats.get("nicht_nachrechenbar", 0)
    line = (T(f"{formulas} Formeln, davon {formulas - unknown} nachgerechnet",
              f"{formulas} {_plural(formulas, 'formula', 'formulas')}, {formulas - unknown} of them recalculated")
            if formulas else T("keine Formeln", "no formulas"))
    hints = found.get("hints") or []
    rows = []
    for h in hints:
        kind = h["key"].split(":", 1)[0]
        what = escape(sentence(h["text"])) if with_data else escape(kind_name(kind, T("Hinweis", "Note")))
        tip = f"<div class=tip>{escape(sentence(h['tip']))}</div>" if with_data and h.get("tip") else ""
        rows.append(f"<tr class={h['level']}><td>{_level(h['level'])}</td><td class=cell>{escape(h.get('cell') or '–').replace(',', ', ')}</td>"
                    f"<td>{what}{tip}</td></tr>")
    nothing = (T("Mir fällt nichts weiter auf.", "Nothing else stands out to me.") if found.get("passt")
               else T("Mir fällt nichts auf.", "Nothing stands out to me."))
    table = (T("<table><thead><tr><th>Stufe</th><th>Zelle</th><th>Was</th></tr></thead><tbody>",
               "<table><thead><tr><th>Level</th><th>Cell</th><th>What</th></tr></thead><tbody>") + "".join(rows) + "</tbody></table>"
             if rows else f"<p>{nothing}</p>")
    if found.get("passt"):
        # Wer den Bericht liest, soll wissen, dass jemand Hinweise angeschaut und für richtig befunden hat.
        n = found["passt"]
        table += T(f"<p class=meta>{n} {_plural(n, 'Hinweis', 'Hinweise')} mit „Passt so“ ausgeblendet.</p>",
                   f"<p class=meta>{n} {_plural(n, 'finding', 'findings')} hidden with “Looks right”.</p>")
    drivers = ""
    if with_data and found.get("drivers"):
        items = "".join(T(f"<li>{escape(d['output'])} = {escape(str(d['output_value']))}: 1 % mehr bei {escape(d['input'])} "
                          f"({escape(str(d['input_value']))}) ergibt {_percent(d['elasticity'])}</li>",
                          f"<li>{escape(d['output'])} = {escape(str(d['output_value']))}: 1% more in {escape(d['input'])} "
                          f"({escape(str(d['input_value']))}) gives {_percent(d['elasticity'])}</li>")
                        for d in found["drivers"][:8])
        drivers = f"<h3>{T('Wovon die Ergebnisse abhängen', 'What the results depend on')}</h3><ul>{items}</ul>"
    words = _words(described) if with_data and described else ""
    return f"<section><h2>{escape(name)}</h2><p class=meta>{line}</p>{table}{words}{drivers}</section>"


def _percent(value: float) -> str:
    text = f"{abs(value):.2f}"
    if english():
        return f"{'−' if value < 0 else '+'}{text}%"
    return f"{'−' if value < 0 else '+'}{text.replace('.', ',')} %"


def _page(title: str, body: str) -> str:
    return f"""<!doctype html>
<html lang="{lang()}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
:root {{ --fg: #1d1d1b; --muted: #6b6b66; --line: #e2e2dc; --bad: #b23b3b; --warn: #8a5a00; --bg: #fff; }}
@media (prefers-color-scheme: dark) {{ :root {{ --fg: #ececea; --muted: #a0a09a; --line: #34342f; --bad: #f08a8a; --warn: #e3a35b; --bg: #171716; }} }}
@media print {{ :root {{ --fg: #000; --muted: #555; --line: #ccc; --bad: #a00; --warn: #7a4a00; --bg: #fff; }} }}
body {{ font: 14px/1.5 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: var(--fg); background: var(--bg);
  max-width: 900px; margin: 24px auto; padding: 0 16px; }}
h1 {{ font-size: 22px; margin: 0 0 4px; }} h2 {{ font-size: 17px; margin: 24px 0 2px; }} h3 {{ font-size: 14px; }}
.meta, .foot, .tip {{ color: var(--muted); font-size: 12px; }}
.sum {{ font-size: 15px; }}
table {{ width: 100%; border-collapse: collapse; margin-top: 6px; }}
th, td {{ text-align: left; vertical-align: top; padding: 6px 8px; border-bottom: 1px solid var(--line); }}
td:first-child {{ white-space: nowrap; font-weight: 600; }}
th:nth-child(1), th:nth-child(2) {{ width: 90px; }}
tr.fehler td:first-child {{ color: var(--bad); }} tr.pruefen td:first-child {{ color: var(--warn); }}
.cell {{ font-family: ui-monospace, SFMono-Regular, Menlo, monospace; white-space: nowrap; }}
.num {{ white-space: nowrap; font-variant-numeric: tabular-nums; }} .bad {{ color: var(--bad); }}
li {{ margin: 2px 0; }} li .cell {{ white-space: normal; overflow-wrap: anywhere; }}
section {{ break-inside: auto; }} tr {{ break-inside: avoid; }}
</style>
</head>
<body>
{body}
</body>
</html>
"""


def save(folder: Path, html: str, with_data: bool, when: datetime | None = None) -> str:
    """Den Bericht im Datenordner ablegen. Gibt den Dateinamen zurück (für /berichte/<name>)."""
    when = when or datetime.now()
    # Der Zufallsteil macht die Adresse unratbar: Mit Werten stehen darin Daten aus der Mappe.
    name = f"{when.strftime('%Y%m%d-%H%M%S')}-{secrets.token_hex(4)}-pruefbericht{'' if with_data else '-ohne-daten'}.html"
    folder.mkdir(parents=True, exist_ok=True)
    tidy(folder)
    (folder / name).write_text(html, encoding="utf-8")
    # Nur die letzten Berichte behalten.
    old = sorted(p for p in folder.iterdir() if NAME.fullmatch(p.name))[:-KEEP]
    for path in old:
        path.unlink(missing_ok=True)
    return name


def tidy(folder: Path, now: float | None = None) -> None:
    """Berichte, die älter als einen Tag sind, löschen. Ein Bericht liegt nur im Datenordner, damit der Browser ihn
    öffnen kann (das Seitenfenster in Excel kann nichts herunterladen); wer ihn behalten will, sichert ihn dort als
    PDF. So bleiben keine Werte aus einer Mappe auf der Festplatte liegen."""
    now = time.time() if now is None else now
    try:
        paths = [p for p in folder.iterdir() if NAME.fullmatch(p.name)]
    except OSError:
        return
    for path in paths:
        try:
            if now - path.stat().st_mtime > MAX_AGE:
                path.unlink(missing_ok=True)
        except OSError:
            pass
