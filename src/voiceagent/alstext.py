"""Zahlen und Daten, die als Text in der Zelle stehen.

Aus einem Import (CSV, ein anderes Programm, Kopieren aus dem Browser) kommen Zahlen und Daten oft als Text
in die Mappe: „4855.45“ oder „23.02.2026“ sehen richtig aus, aber Excel rechnet nicht mit ihnen. SUMME und
SUMMEWENN lassen eine Zahl als Text still weg, SUMMEWENNS mit Datumsgrenzen, MONAT() und Sortieren übersehen
ein Datum als Text. Der Prüfer findet solche Zellen in Zahlen- und Datumsspalten und rechnet nach, welche
Formeln dadurch etwas anderes ergeben: einmal wie Excel jetzt, einmal mit den Zellen als echte Zahl oder
echtes Datum.
"""

from __future__ import annotations

import math
import re
from datetime import date

from . import formel
from .analyse import Finding, Sheet, address, more_formulas, said
from .sprache import T

Pos = tuple[int, int]

MAX_LISTED = 3
# Mehr Formeln rechnet er für die Folgen nicht zweimal nach (große Blätter bleiben flüssig).
MAX_FORMULAS = 4000

_NUMBER_FORMS = [
    (re.compile(r"\d+"), lambda t: t),
    (re.compile(r"\d+\.(?:\d{1,2}|\d{4,})"), lambda t: t),
    (re.compile(r"\d*,\d+"), lambda t: t.replace(",", ".")),
    (re.compile(r"\d{1,3}(?:\.\d{3})+(?:,\d+)?"), lambda t: t.replace(".", "").replace(",", ".")),
    (re.compile(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?"), lambda t: t.replace(",", "")),
    (re.compile(r"\d{1,3}(?:[ '’]\d{3})+(?:[.,]\d+)?"), lambda t: re.sub(r"[ '’]", "", t).replace(",", ".")),
    # 1.5e3, 1,5E+03: so schreiben Messgeräte und Statistikprogramme.
    (re.compile(r"\d+(?:[.,]\d+)?[eE][+-]?\d{1,3}"), lambda t: t.replace(",", ".")),
]
_ISO = re.compile(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2}))?)?")
_SLASH = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{2}|\d{4})")
_DOTTED = re.compile(r"(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})(?: (\d{1,2}):(\d{2})(?::(\d{2}))?)?")
# Ein Betrag mit Währung, wie man ihn tippt oder aus einer Rechnung kopiert: „85,50 €“, „CHF 1'200.00“, „$12.99“.
_MONEY = r"(?:€|EUR|CHF|SFr\.?|Fr\.|\$|USD|£|GBP|kr\.?|SEK|NOK|DKK|PLN|zł|CZK|Kč|¥|JPY)"
_CURRENCY = re.compile(rf"^{_MONEY}\s*|\s*{_MONEY}$")
_MONTHS = {
    "jan": 1, "januar": 1, "january": 1, "jänner": 1, "feb": 2, "februar": 2, "february": 2, "mär": 3, "mrz": 3,
    "märz": 3, "mar": 3, "march": 3, "apr": 4, "april": 4, "mai": 5, "may": 5, "jun": 6, "juni": 6, "june": 6,
    "jul": 7, "juli": 7, "july": 7, "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9, "okt": 10,
    "oct": 10, "oktober": 10, "october": 10, "nov": 11, "november": 11, "dez": 12, "dec": 12, "dezember": 12,
    "december": 12,
}
_NAMED_DAY = re.compile(r"(\d{1,2})\.?[ -]([A-Za-zÄäÖöÜü]{3,9})\.?[ -](\d{4}|\d{2})")
_NAMED_FIRST = re.compile(r"([A-Za-zÄäÖöÜü]{3,9})\.? (\d{1,2}),? (\d{4})")


def number(text: str) -> float | None:
    """„4855.45“, „4.855,45“, „1 234“, „85,50 €“, „(12,50)“, „-€12.00“, „78,5 %“, „1,5E+03“ → die Zahl, die gemeint
    ist. Sonst None. Prozent wie in Excel: „78,5 %“ ist 0,785."""
    plain = re.sub(r"[\u00a0\u202f\u2009]", " ", text.strip()).replace("\u2212", "-")
    sign = 1.0
    if plain.startswith("(") and plain.endswith(")"):
        # So schreibt die Buchhaltung einen negativen Betrag: (12,50).
        plain, sign = plain[1:-1].strip(), -1.0
    for _ in range(2):
        if plain.startswith(("+", "-")):
            sign, plain = (-sign if plain[0] == "-" else sign), plain[1:].strip()
        plain = _CURRENCY.sub("", plain).strip()
    if plain.endswith("-") and sign > 0:
        sign, plain = -1.0, plain[:-1].strip()  # „12,50-“ aus SAP und anderen Buchhaltungen
    percent = plain.endswith("%")
    if percent:
        plain = plain[:-1].strip()
    for pattern, clean in _NUMBER_FORMS:
        if pattern.fullmatch(plain):
            try:
                value = sign * float(clean(plain))
            except ValueError:
                return None
            return value / 100 if percent else value
    return None


# Stückzahlen mit ihrem Wort dahinter („12 Stk“) meinen die Zahl, egal wie die Spalte heißt. Andere Einheiten nur,
# wenn der Kopf dieselbe nennt („72 kg“ unter „Gewicht kg“): „500 g“ unter „Gewicht kg“ ist eine andere Einheit.
_COUNTED = {"stk", "stk.", "stück", "st", "st.", "x", "pkt", "pkt.", "stck", "pcs", "pc"}
_UNIT = re.compile(r"(.*?\d)\s*([A-Za-zÄäÖöÜüß]{1,6}\.?)")


def with_unit(text: str, header: str) -> tuple[float, str] | None:
    """„12 Stk“ → (12, „Stk“); „72 kg“ unter „Gewicht kg“ → (72, „kg“). Sonst None."""
    match = _UNIT.fullmatch(re.sub(r"[\u00a0\u202f\u2009]", " ", text.strip()))
    if match is None:
        return None
    unit = match.group(2)
    named = {w.casefold() for w in re.findall(r"[A-Za-zÄäÖöÜüß]+", header)}
    if unit.casefold() not in _COUNTED and unit.casefold().rstrip(".") not in named:
        return None
    value = number(match.group(1))
    return None if value is None or match.group(1).rstrip().endswith("%") else (value, unit)


def _serial(year: int, month: int, day_: int, hour: str | None = None, minute: str | None = None,
            second: str | None = None) -> float | None:
    if year < 100:
        year += 2000 if year < 30 else 1900  # wie Excel: 29 ist 2029, 30 ist 1930
    try:
        days = float((date(year, month, day_) - date(1899, 12, 30)).days)
    except ValueError:
        return None
    if hour is not None:
        h, m, sec = int(hour), int(minute or 0), int(second or 0)
        if h > 23 or m > 59 or sec > 59:
            return None
        days += (h * 3600 + m * 60 + sec) / 86400
    return days


def day(text: str, order: str = "dmy") -> float | None:
    """„23.02.2026“, „23.02.26“, „2026-02-23“, „2026/02/23“, „23/02/2026“, „23. Februar 2026“, „Feb 23, 2026“ → die
    Excel-Zahl des Tages (mit Uhrzeit, wenn eine dabei ist). Sonst None. order: wie Tag und Monat mit Schrägstrich
    stehen, „dmy“ (Europa) oder „mdy“ (USA)."""
    plain = " ".join(text.strip().split())
    match = _ISO.fullmatch(plain)
    if match:
        return _serial(int(match[1]), int(match[2]), int(match[3]), match[4], match[5], match[6])
    match = _DOTTED.fullmatch(plain)
    if match:
        return _serial(int(match[3]), int(match[2]), int(match[1]), match[4], match[5], match[6])
    match = _SLASH.fullmatch(plain)
    if match:
        a, b = int(match[1]), int(match[2])
        d, m = (a, b) if order == "dmy" else (b, a)
        return _serial(int(match[3]), m, d)
    match = _NAMED_DAY.fullmatch(plain)
    if match and match[2].casefold() in _MONTHS:
        return _serial(int(match[3]), _MONTHS[match[2].casefold()], int(match[1]))
    match = _NAMED_FIRST.fullmatch(plain)
    if match and match[1].casefold() in _MONTHS:
        return _serial(int(match[3]), _MONTHS[match[1].casefold()], int(match[2]))
    return None


def slash_order(texts) -> str | None:
    """Wie eine Spalte Daten mit Schrägstrich schreibt: „dmy“, wenn ein Tag über 12 vorne steht (23/02/2026), „mdy“,
    wenn er in der Mitte steht (02/23/2026), None, wenn das nicht zu sehen ist oder beides vorkommt."""
    first = second = False
    for text in texts:
        match = _SLASH.fullmatch(" ".join(text.strip().split()))
        if match:
            first |= int(match[1]) > 12
            second |= int(match[2]) > 12
    if first == second:
        return None
    return "dmy" if first else "mdy"


def _same(a, b) -> bool:
    if isinstance(a, float) and isinstance(b, float) and not isinstance(a, bool) and not isinstance(b, bool):
        return math.isclose(a, b, rel_tol=1e-9, abs_tol=1e-9)
    return a == b


def _effects(sheet: Sheet, fixed: dict[Pos, float], col: int) -> list[tuple[Pos, object, object]]:
    """Welche Formeln anders rechnen, weil die Spalte col Text enthält: (Zelle, jetzt, richtig). Richtig heißt mit
    allen Werten als echte Zahl oder echtes Datum; anders rechnen sie, wenn nur diese Spalte Text bleibt. So
    verdeckt ein Datum als Text nicht, was eine Zahl als Text in derselben Zeile bewirkt, und umgekehrt."""
    chain = {pos for pos in sheet.nodes if pos not in sheet.unknown}
    if not chain or len(chain) > MAX_FORMULAS:
        return []
    _, now = sheet.evaluator(chain=chain)
    _, right = sheet.evaluator(overrides=fixed, chain=chain)
    _, without = sheet.evaluator(overrides={p: v for p, v in fixed.items() if p[1] != col}, chain=chain)
    changed = []
    for pos in sorted(chain):
        try:
            before, after, other = now(pos), right(pos), without(pos)
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, IndexError, RecursionError):
            continue
        if not _same(after, other):
            changed.append((pos, before, after))
    return changed


def _shown(value) -> str:
    return said(value)


def checks(sheet: Sheet, columns: dict) -> list[Finding]:
    """Ein Hinweis je Zahlen- oder Datumsspalte mit Werten als Text, mit den Formeln, die dadurch anders rechnen."""
    found = []
    texts_of = {}
    from .auswertung import is_id

    kinds = {}
    for col, info in columns.items():
        # Fallnummern stehen mit Absicht als Text da, ebenso Nummern mit Null vorn („00123“).
        if is_id(sheet.plan.headers.get(col, "")):
            continue
        filled = len(info.numbers) + len(info.dates) + len(info.texts)
        order = slash_order(info.texts.values())
        scale = _percent_scale(sheet, col, info)
        header = sheet.plan.headers.get(col, "")

        def as_number(t: str, scale=scale, header=header) -> float | None:
            value = _scaled(number(t), t, scale)
            if value is None and (unit := with_unit(t, header)) is not None:
                value = unit[0]
            return value

        reads = (("zahl", as_number, info.numbers), ("datum", lambda t: day(t, order or "dmy"), info.dates))
        for kind, read, real in reads:
            texts = {r: (t, value) for r, t in sorted(info.texts.items())
                     if not (kind == "zahl" and re.match(r"\s*-?0\d", t)) and (value := read(t)) is not None}
            # Eine Zahlen- oder Datumsspalte, auch wenn viele Werte als Text dastehen.
            if texts and real and len(real) + len(texts) >= 0.8 * filled:
                texts_of[col], kinds[col] = texts, kind
                break
    fixed = {(r, col): value for col, texts in texts_of.items() for r, (_, value) in texts.items()}
    for col, texts in texts_of.items():
        kind = kinds[col]
        header = sheet.plan.headers[col]
        cells = [(r, col) for r in texts]
        units = kind == "zahl" and all(number(t) is None for t, _ in texts.values())
        n = len(cells)
        what = (T("eine Zahl mit Einheit", "a number with a unit") if units else
                T("eine Zahl", "a number") if kind == "zahl" else T("ein Datum", "a date"))
        many = (T("Zahlen mit Einheit", "numbers with a unit") if units else
                T("Zahlen", "numbers") if kind == "zahl" else T("Daten", "dates"))
        listed = ", ".join(T(f"{address(p)} („{texts[p[0]][0].strip()}“)", f"{address(p)} (“{texts[p[0]][0].strip()}”)")
                           for p in cells[:MAX_LISTED])
        if len(cells) > MAX_LISTED:
            listed += T(f" und {len(cells) - MAX_LISTED} weitere", f" and {len(cells) - MAX_LISTED} more")
        lead = (T(f"In {header} steht bei {listed} {what} als Text.", f"In {header}, {listed} holds {what} as text.")
                if n == 1 else
                T(f"In {header} stehen {n} {many} als Text: {listed}.", f"In {header}, {n} {many} are stored as text: {listed}."))
        it = "sie" if n > 1 or kind == "zahl" else "es"
        them = "them" if n > 1 else "it"
        why = (T(f" Excel rechnet nicht damit: SUMME, SUMMEWENN und MITTELWERT lassen {it} still weg.",
                 f" Excel does not calculate with {them}: SUM, SUMIF and AVERAGE silently leave {them} out.")
               if kind == "zahl" else
               T(" Excel rechnet nicht damit: Bedingungen mit Datumsgrenzen (SUMMEWENNS, ZÄHLENWENNS), MONAT() und "
                 f"Sortieren übersehen {it}.",
                 f" Excel does not calculate with {them}: conditions with date limits (SUMIFS, COUNTIFS), MONTH() and "
                 f"sorting miss {them}."))
        effects = [e for e in _effects(sheet, fixed, col) if e[0][1] != col]
        # Zuerst Kennzahlen mit Namen außerhalb der Tabelle, dann der Rest.
        effects.sort(key=lambda e: (sheet.plan.in_table(*e[0]), e[0]))
        shown = [T(f"{sheet.label(pos)} {_shown(before)} statt {_shown(after)}",
                   f"{sheet.label(pos)} gives {_shown(before)} instead of {_shown(after)}")
                 for pos, before, after in effects[:MAX_LISTED]]
        if shown:
            more = more_formulas(len(effects) - MAX_LISTED) if len(effects) > MAX_LISTED else ""
            effect = T(f" Dadurch ergibt {', '.join(shown)}{more}.", f" As a result, {', '.join(shown)}{more}.")
        else:
            effect = T(" Bisher rechnet keine Formel damit.", f" No formula calculates with {them} yet.")
        tip = (T("Trag nur die Zahl ein; die Einheit gehört in den Spaltenkopf oder in das Zahlenformat.",
                 "Enter only the number; the unit belongs in the column header or in the number format.") if units else
               T("Markiere die Zellen, klick auf das kleine Warnzeichen daneben und wähle „In eine Zahl umwandeln“. "
                 "Für die ganze Spalte: Daten → Text in Spalten → Fertig stellen.",
                 "Select the cells, click the small warning icon next to them and choose “Convert to Number”. "
                 "For the whole column: Data → Text to Columns → Finish.")
               if kind == "zahl" else
               T("Markiere die Spalte und wähle Daten → Text in Spalten → Fertig stellen, dann sind es echte Daten. "
                 "Oder gib die Zelle neu ein.",
                 "Select the column and choose Data → Text to Columns → Finish; then they are real dates. "
                 "Or enter the cell again."))
        found.append(Finding(f"astext:{address(cells[0])}", lead + why + effect, "fehler" if effects else "pruefen", tip,
                             fix=_fix(sheet, col, texts, kind, columns[col])))
    return found


def _percent_scale(sheet: Sheet, col: int, info) -> float:
    """Was „78,5 %“ als Text in dieser Spalte meint: 0,785 wie in Excel, wenn die echten Zahlen als Prozent dastehen
    oder klein sind, 78,5 in einer Spalte, die Prozentpunkte als Zahl führt (Kopf „Anteil %“, Werte wie 78,5)."""
    values = sorted(abs(v) for v in info.numbers.values())
    if not values:
        return 1.0
    shown = [sheet.cells[(r, col)].value for r in info.numbers if (r, col) in sheet.cells]
    if sum(v.rstrip().endswith("%") for v in shown) * 2 >= len(shown):
        return 1.0
    return 100.0 if values[len(values) // 2] > 1 else 1.0


def _scaled(value: float | None, text: str, scale: float) -> float | None:
    return value * scale if value is not None and text.rstrip().endswith("%") else value


# „1,234“ und „1.234“ können tausendmal mehr oder weniger sein, je nachdem, woher sie kommen: nicht einsetzen.
_UNSURE = re.compile(r"\s*-?\d{1,3}[.,]\d{3}\s*")


def _fix(sheet: Sheet, col: int, texts: dict, kind: str, info) -> dict | None:
    """Zum Einsetzen: die Zahlen (ein Datum als seine Excel-Zahl) statt der Texte, im Zahlenformat einer echten Zahl
    oder eines echten Datums derselben Spalte."""
    from . import einsetzen

    if kind == "zahl" and any(_UNSURE.fullmatch(t) for t, _ in texts.values()):
        return None
    # 03/04/2026 ist der 3. April oder der 4. März: Steht in der Spalte kein Tag über 12, nicht einsetzen.
    if kind == "datum" and slash_order(t for t, _ in texts.values()) is None and any(_SLASH.fullmatch(t.strip()) for t, _ in texts.values()):
        return None
    real = info.numbers if kind == "zahl" else info.dates
    like = min(real) if real else None
    if len(texts) == 1:
        (t, _), = texts.values()
        said = T(f"„{t.strip()}“ als {'Zahl' if kind == 'zahl' else 'Datum'}",
                 f"“{t.strip()}” as a {'number' if kind == 'zahl' else 'date'}")
    else:
        said = T(f"{len(texts)} echte {'Zahlen' if kind == 'zahl' else 'Daten'}",
                 f"{len(texts)} real {'numbers' if kind == 'zahl' else 'dates'}")
    return einsetzen.values(sheet, {(r, col): value for r, (_, value) in texts.items()},
                            None if like is None else (like, col), said)
