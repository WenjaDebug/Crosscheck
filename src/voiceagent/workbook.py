"""Excel verstehen: Zellen, Formeln und Aufbau einer Mappe, gelesen aus dem Seitenfenster in Excel (Office.js),
aus dem laufenden Excel (excel.py) oder aus der gespeicherten Datei (xlsx.py).

Was in einer Mappe eine Entscheidung ist:
- eine Formel, neu oder geändert: „Formel Sitze geändert: keine → =Wähleranteil * 577 / 100
  (Zeilen Parteien RN bis Rest)“. Eine nach unten gezogene Formel ist ein Ereignis, nicht fünf.
- ein Wert, der schon da war und geändert oder gelöscht wird.
- Filter, Sortierung, Hervorhebung (bedingte Formatierung), Auswahllisten und andere
  Eingaberegeln, Diagramme, Pivot-Tabellen, Excel-Tabellen, Namen und Blätter.

Daten neu eintippen ist meist Abschreiben, keine Entscheidung, und gibt kein Ereignis. Ebenso
wenig, was sie in derselben Sitzung frisch getippt hat und gleich wieder ändert oder löscht
(„0.28*577“ ohne Gleichheitszeichen getippt, gelöscht, als Formel neu geschrieben).
Werte, die eine Formel ausrechnet, ändern sich mit ihren Eingaben und zählen auch nicht,
ebenso wenig, was eine Pivot-Tabelle anzeigt.

In den Formeln stehen statt Zelladressen die Spaltenköpfe: aus =B3*577/100 wird
=Wähleranteil(%) * 577 / 100. So kann der Prüfer sagen, was eine Formel rechnet.
Die Kopfzeile darf tiefer stehen, mit einem Titel darüber und Leerzeilen dazwischen.
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field

from . import formel, woerter
from .events import RECORD_SWITCH, Event
from .sprache import T, english
from .xlsx import (  # noqa: F401 - read_xlsx, read_book, shift und MAX_FILE gehören zur Schnittstelle
    ANY_REF,
    MAX_FILE,
    MAX_FORMULA,
    STRING,
    STRUCTURED,
    Area,
    Cell,
    ExcelTable,
    Rule,
    Sheet,
    area_of,
    area_text,
    column_letters,
    column_number,
    in_area,
    number_text,
    outside_literals,
    read_book,
    read_xlsx,
    said,
    sheet_of,
    shift,
)

# So weit unten wird die Kopfzeile noch gesucht.
HEADER_ROWS = 30
# Mehr Formel-Ereignisse auf einmal: Da wurde kopiert oder ein Blatt eingefügt.
MAX_FORMULA_EVENTS = 6
# Mehr Änderungen an Filtern, Diagrammen … auf einmal: eine andere Fassung der Datei.
MAX_PART_EVENTS = 6
# So lang wird die Übersicht für die KI höchstens.
MAX_OVERVIEW = 1800
NONE = "keine"
EMPTY = "(leer)"
FORMULA = "Formel"
SORTING = "Sortierung"
# Womit die Bezeichnung eines Ereignisses aus Excel anfängt (außer bei Werten).
PREFIXES = (
    FORMULA, "Filter", "Hervorhebung", "Auswahlliste", "Eingaberegel", "Diagramm", "Pivot-Tabelle",
    "Excel-Tabelle", "Name", "Blatt",
)

_ADDRESS = re.compile(r"^([A-Z]{1,3})(\d{1,7})$")
_LIKE_CELL = re.compile(r"\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})")


def like_cell(text: str) -> bool:
    """F242398, C2, Q1: Text, den man in einem Hinweis für eine Zelladresse hält."""
    match = _LIKE_CELL.fullmatch(text.strip())
    return bool(match) and column_number(match.group(1).upper()) <= 16384 and int(match.group(2)) <= 1048576
# Eine Zelladresse auf diesem Blatt, z. B. B3, $B$3, B$3. Nicht Teil eines Namens oder Funktionsnamens,
# und nicht auf anderen Blättern (Tabelle2!B3): Die haben dort andere Spaltenköpfe.
_REF = re.compile(r"(?<![A-Za-z0-9_.$!\\])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(])")
_OTHER_SHEET = re.compile(r"(?:'[^']+'|[A-Za-z0-9_.]+)!\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?")
_RANGE = re.compile(r"(?<![A-Za-z0-9_.$!\\])(\$?[A-Z]{1,3}\$?\d{1,7}):(\$?[A-Z]{1,3}\$?\d{1,7})")
_DECIMAL = re.compile(r"(?<![A-Za-z0-9_])(\d+)\.(\d+)")
_FUNCTION = re.compile(r"(?<![A-Za-z0-9_.])(?:_xlfn\.|_xlws\.)?([A-Z][A-Z0-9.]*)\(")
_NUMBER = re.compile(r"(?<![A-Za-z0-9_$.,])\d+(?:[.,]\d+)?(?![A-Za-z0-9_(])")
_YEAR = re.compile(r"^(19|20)\d\d$")

# Die englischen Namen in der Datei, wie die deutsche Excel sie zeigt.
FUNCTIONS = {
    "SUM": "SUMME", "AVERAGE": "MITTELWERT", "ROUND": "RUNDEN", "ROUNDDOWN": "ABRUNDEN", "ROUNDUP": "AUFRUNDEN",
    "IF": "WENN", "IFS": "WENNS", "IFERROR": "WENNFEHLER", "IFNA": "WENNNV", "AND": "UND", "OR": "ODER", "NOT": "NICHT",
    "COUNT": "ANZAHL", "COUNTA": "ANZAHL2", "COUNTBLANK": "ANZAHLLEEREZELLEN", "COUNTIF": "ZÄHLENWENN",
    "COUNTIFS": "ZÄHLENWENNS", "SUMIF": "SUMMEWENN", "SUMIFS": "SUMMEWENNS", "AVERAGEIF": "MITTELWERTWENN",
    "AVERAGEIFS": "MITTELWERTWENNS", "MAXIFS": "MAXWENNS", "MINIFS": "MINWENNS", "SUMPRODUCT": "SUMMENPRODUKT",
    "SUBTOTAL": "TEILERGEBNIS", "VLOOKUP": "SVERWEIS", "HLOOKUP": "WVERWEIS", "XLOOKUP": "XVERWEIS",
    "MATCH": "VERGLEICH", "XMATCH": "XVERGLEICH", "CHOOSE": "WAHL", "INDIRECT": "INDIREKT", "OFFSET": "BEREICH.VERSCHIEBEN",
    "INT": "GANZZAHL", "MOD": "REST", "LARGE": "KGRÖSSTE", "SMALL": "KKLEINSTE", "RANK": "RANG",
    "RANK.EQ": "RANG.GLEICH", "TODAY": "HEUTE", "NOW": "JETZT", "DATE": "DATUM", "YEAR": "JAHR", "MONTH": "MONAT",
    "DAY": "TAG", "HOUR": "STUNDE", "WEEKDAY": "WOCHENTAG", "EOMONTH": "MONATSENDE", "EDATE": "EDATUM",
    "NETWORKDAYS": "NETTOARBEITSTAGE", "YEARFRAC": "BRTEILJAHRE", "STDEV": "STABW", "STDEV.S": "STABW.S",
    "STDEV.P": "STABW.N", "STDEVP": "STABWN", "VAR": "VARIANZ", "VARP": "VARIANZEN", "AVERAGEA": "MITTELWERTA", "CORREL": "KORREL", "PERCENTILE": "QUANTIL", "PERCENTILE.INC": "QUANTIL.INKL",
    "QUARTILE": "QUARTILE", "TRUNC": "KÜRZEN", "CONCATENATE": "VERKETTEN", "CONCAT": "TEXTKETTE",
    "TEXTJOIN": "TEXTVERKETTEN", "LEFT": "LINKS", "RIGHT": "RECHTS", "MID": "TEIL", "LEN": "LÄNGE", "FIND": "FINDEN",
    "SEARCH": "SUCHEN", "SUBSTITUTE": "WECHSELN", "TRIM": "GLÄTTEN", "UPPER": "GROSS", "LOWER": "KLEIN", "VALUE": "WERT",
    "FLOOR": "UNTERGRENZE", "CEILING": "OBERGRENZE", "SQRT": "WURZEL", "POWER": "POTENZ", "ROW": "ZEILE",
    "ROWS": "ZEILEN", "COLUMN": "SPALTE", "COLUMNS": "SPALTEN", "ISBLANK": "ISTLEER", "ISNUMBER": "ISTZAHL",
    "ISERROR": "ISTFEHLER", "ISEVEN": "ISTGERADE", "ISODD": "ISTUNGERADE", "SWITCH": "ERSTERWERT",
    "UNIQUE": "EINDEUTIG", "SORT": "SORTIEREN", "SEQUENCE": "SEQUENZ", "TRANSPOSE": "MTRANS", "RAND": "ZUFALLSZAHL",
    "QUOTIENT": "QUOTIENT", "FLOOR.MATH": "UNTERGRENZE.MATHEMATIK",
    "ABS": "ABS", "T.TEST": "T.TEST", "TTEST": "TTEST", "CONFIDENCE.NORM": "KONFIDENZ.NORM", "CONFIDENCE.T": "KONFIDENZ.T",
    "CONFIDENCE": "KONFIDENZ", "NORM.S.INV": "NORM.S.INV", "NORMSINV": "STANDNORMINV", "NORM.S.DIST": "NORM.S.VERT",
    "NORMSDIST": "STANDNORMVERT", "T.INV.2T": "T.INV.2S", "TINV": "TINV", "T.INV": "T.INV", "T.DIST": "T.VERT",
    "T.DIST.RT": "T.VERT.RE", "T.DIST.2T": "T.VERT.2S", "TDIST": "TVERT",
    "NA": "NV", "PI": "PI", "SIN": "SIN", "COS": "COS", "TAN": "TAN", "ASIN": "ARCSIN", "ACOS": "ARCCOS", "ATAN": "ARCTAN",
    "ATAN2": "ARCTAN2", "DEGREES": "GRAD", "RADIANS": "BOGENMASS", "FACT": "FAKULTÄT", "LN": "LN", "LOG10": "LOG10",
    "EXP": "EXP", "TRUE": "WAHR", "FALSE": "FALSCH",
    "STEYX": "STFEHLERYX", "COVAR": "KOVAR", "GEOMEAN": "GEOMITTEL", "HARMEAN": "HARMITTEL", "DEVSQ": "SUMQUADABW",
    "AVEDEV": "MITTELABW", "SKEW": "SCHIEFE", "KURT": "KURT", "STANDARDIZE": "STANDARDISIERUNG",
    "NPV": "NBW", "XNPV": "XKAPITALWERT", "IRR": "IKV", "PMT": "RMZ", "PV": "BW", "FV": "ZW", "IPMT": "ZINSZ",
    "PPMT": "KAPZ", "NPER": "ZZR", "MMULT": "MMULT",
}

# Seltenere Funktionen (englisch → deutsch): damit ein Name in der Formel als bekannt gilt, ein Tippfehler einen
# Vorschlag bekommt (fehlerwerte.py) und Hinweise =ARBEITSTAG(…) schreiben statt =WORKDAY(…).
FUNCTIONS.update({
    "ADDRESS": "ADRESSE", "AGGREGATE": "AGGREGAT", "CEILING.MATH": "OBERGRENZE.MATHEMATIK", "CELL": "ZELLE",
    "CHAR": "ZEICHEN", "CHISQ.TEST": "CHIQU.TEST", "CHOOSECOLS": "SPALTENWAHL", "CHOOSEROWS": "ZEILENWAHL",
    "CLEAN": "SÄUBERN", "CODE": "CODE", "COVARIANCE.P": "KOVARIANZ.P", "COVARIANCE.S": "KOVARIANZ.S",
    "DATEDIF": "DATEDIF", "DATEVALUE": "DATWERT", "DAYS": "TAGE", "DAYS360": "TAGE360", "DROP": "WEGLASSEN",
    "EVEN": "GERADE", "EXACT": "IDENTISCH", "EXP": "EXP", "F.TEST": "F.TEST", "FACT": "FAKULTÄT", "FILTER": "FILTER",
    "FIXED": "FEST", "FORECAST": "PROGNOSE", "FORECAST.LINEAR": "PROGNOSE.LINEAR", "FORMULATEXT": "FORMELTEXT",
    "FV": "ZW", "GCD": "GGT", "GROWTH": "VARIATION", "HSTACK": "HSTAPELN", "HYPERLINK": "HYPERLINK", "INDEX": "INDEX",
    "INTERCEPT": "ACHSENABSCHNITT", "IRR": "IKV", "ISERR": "ISTFEHL", "ISFORMULA": "ISTFORMEL", "ISLOGICAL": "ISTLOG",
    "ISNA": "ISTNV", "ISOWEEKNUM": "ISOKALENDERWOCHE", "ISTEXT": "ISTTEXT", "LAMBDA": "LAMBDA", "LCM": "KGV",
    "LET": "LET", "LINEST": "RGP", "LN": "LN", "LOG": "LOG", "LOG10": "LOG10", "LOOKUP": "VERWEIS", "MAX": "MAX",
    "MAXA": "MAXA", "MEDIAN": "MEDIAN", "MIN": "MIN", "MINA": "MINA", "MINUTE": "MINUTE", "MODE": "MODALWERT",
    "MODE.SNGL": "MODUS.EINF", "MROUND": "VRUNDEN", "N": "N", "NA": "NV", "NETWORKDAYS.INTL": "NETTOARBEITSTAGE.INTL",
    "NORM.DIST": "NORM.VERT", "NORM.INV": "NORM.INV", "NPV": "NBW", "NUMBERVALUE": "ZAHLENWERT", "ODD": "UNGERADE",
    "PEARSON": "PEARSON", "PERCENTILE.EXC": "QUANTIL.EXKL", "PERCENTRANK": "QUANTILSRANG", "PI": "PI", "PMT": "RMZ",
    "PRODUCT": "PRODUKT", "PROPER": "GROSS2", "PV": "BW", "QUARTILE.EXC": "QUARTILE.EXKL",
    "QUARTILE.INC": "QUARTILE.INKL", "RANDARRAY": "ZUFALLSMATRIX", "RANDBETWEEN": "ZUFALLSBEREICH",
    "RANK.AVG": "RANG.MITTELW", "RATE": "ZINS", "REPT": "WIEDERHOLEN", "RSQ": "BESTIMMTHEITSMASS", "SECOND": "SEKUNDE",
    "SHEET": "BLATT", "SIGN": "VORZEICHEN", "SLOPE": "STEIGUNG", "SORTBY": "SORTIERENNACH", "STDEVA": "STABWA",
    "SUMSQ": "QUADRATESUMME", "T": "T", "TAKE": "ÜBERNEHMEN", "TEXT": "TEXT", "TEXTAFTER": "TEXTNACH",
    "TEXTBEFORE": "TEXTVOR", "TEXTSPLIT": "TEXTTEILEN", "TIME": "ZEIT", "TIMEVALUE": "ZEITWERT", "TOCOL": "ZUSPALTE",
    "TOROW": "ZUZEILE", "TREND": "TREND", "VAR.P": "VAR.P", "VAR.S": "VAR.S", "VSTACK": "VSTAPELN",
    "WEEKNUM": "KALENDERWOCHE", "WORKDAY": "ARBEITSTAG", "WORKDAY.INTL": "ARBEITSTAG.INTL", "XNPV": "XKAPITALWERT",
    "XOR": "XODER", "Z.TEST": "G.TEST",
})


# --- Tabelle verstehen ---


def _is_word(value: str) -> bool:
    return any(ch.isalpha() for ch in value) and len(value) <= 40


def _filled(cell: Cell) -> bool:
    return bool(cell.value) or cell.formula is not None


@dataclass
class Layout:
    """Kopfzeile und Zeilennamen eines Blattes, damit Formeln und Ereignisse Namen bekommen."""

    header_row: int | None = None
    headers: dict[int, str] = field(default_factory=dict)
    # Spalte der Zeilennamen (meist die erste) und Name je Zeile, z. B. 3 → „RN“.
    name_column: int | None = None
    names: dict[int, str] = field(default_factory=dict)
    last_row: int = 0
    # Die Excel-Tabelle, wenn die Kopfzeile von ihr stammt.
    table: str = ""
    # Was Pivot-Tabellen anzeigen: keine Eingaben, keine Entscheidungen.
    derived: list[Area] = field(default_factory=list)
    # Ein Zeilenmodell (Finanzplan, Budget): jede Zeile eine eigene Größe, die Spalten Jahre oder Monate (_across).
    across: bool = False

    def column(self, col: int) -> str:
        return self.headers.get(col) or column_letters(col)

    def row(self, row: int) -> str:
        if row in self.names and self.name_column is not None:
            return f"{self.headers[self.name_column]} {self.names[row]}"
        return f"Zeile {row}"

    def word(self, row: int) -> str | None:
        """Der Name der Zeile zum Anzeigen („NFP“). Sieht er wie eine Zelladresse aus (eine Fall-Nr F242398, eine
        Station C2), steht der Kopf seiner Spalte davor: „Fall-Nr F242398“. Sonst liest man F242398 als Zelle."""
        name = self.names.get(row)
        if not name:
            return None
        # Eine bloße Zahl als Name („4“ in der Spalte Jahr, eine Fall-Nr 240101) heißt „Jahr 4“: Sonst liest man
        # „Barwert 4“ oder „Liegedauer 240101“ als Betrag.
        bare = name.strip().isdigit() and not _YEAR.match(name.strip())
        if (like_cell(name) or bare) and self.name_column is not None and self.name_column in self.headers:
            return f"{self.headers[self.name_column]} {name}"
        return name

    def counted(self, row: int) -> str:
        """Der Name der Zeile in einer vorgelesenen Formel („D (Ausfall) Jahr 4“), auch wenn er wie eine Zelle aussieht."""
        name = self.names[row]
        return self.word(row) if name.strip().isdigit() else name

    def in_table(self, row: int, col: int) -> bool:
        return self.header_row is not None and col in self.headers and self.header_row < row <= self.last_row

    def is_derived(self, row: int, col: int) -> bool:
        return any(in_area(area, row, col) for area in self.derived)


def _where(plan: Layout, rows: list[int]) -> str:
    """„Zeile Parteien RN“, „Zeile 12“, „Zeilen Parteien RN bis Rest“ oder „Zeilen 12 bis 14“."""
    first, last = rows[0], rows[-1]
    if len(rows) == 1:
        name = plan.row(first)
        return name if name.startswith("Zeile ") else f"Zeile {name}"
    if first in plan.names and plan.name_column is not None:
        return f"Zeilen {plan.row(first)} bis {plan.word(last) or last}"
    return f"Zeilen {first} bis {last}"


def _by_row(cells: dict[tuple[int, int], Cell]) -> dict[int, dict[int, Cell]]:
    rows: dict[int, dict[int, Cell]] = {}
    for (row, col), cell in cells.items():
        if _filled(cell):
            rows.setdefault(row, {})[col] = cell
    return rows


def layout(sheet: Sheet) -> Layout:
    derived = [p.area for p in sheet.pivots if p.area is not None]
    cells = {pos: c for pos, c in sheet.cells.items() if not any(in_area(a, *pos) for a in derived)}
    rows = _by_row(cells)
    tables = [t for t in sheet.tables if t.header and t.columns]
    if tables:
        biggest = max(tables, key=lambda t: (t.area[2] - t.area[0] + 1) * (t.area[3] - t.area[1] + 1))
        result = _table_layout(biggest, rows)
    else:
        result = _guess_layout(rows)
    result.derived = derived
    result.across = _across(cells, result)
    return result


# Weitere Tabellen untereinander sucht others() nur auf kleinen Blättern (Modelle, keine Fall-Listen), höchstens so viele.
BLOCK_CELLS = 5000
MAX_BLOCKS = 6


def others(sheet: Sheet, plan: Layout) -> list[Layout]:
    """Die übrigen Tabellen eines Blatts neben der von layout(): Ein Modell hat oft mehrere untereinander, jede mit
    eigener Kopfzeile (Übergangsmatrix, Verteilung je Jahr, Verlust je Jahr). Jede wird mit ihren eigenen Spalten
    geprüft. Leer bei Excel-Tabellen, großen Blättern und wenn es nur eine Tabelle gibt."""
    if sheet.tables or plan.header_row is None or len(sheet.cells) > BLOCK_CELLS:
        return []
    derived = [p.area for p in sheet.pivots if p.area is not None]
    rows = _by_row({pos: c for pos, c in sheet.cells.items() if not any(in_area(a, *pos) for a in derived)})
    taken = [(plan.header_row, plan.last_row, range(min(plan.headers), max(plan.headers) + 1))]
    found: list[Layout] = []
    for _, row, names, span, data in sorted(_candidates(rows, len(rows)), key=lambda c: -c[0]):
        top, bottom = row, data[-1]
        # Eine Tabelle darunter (andere Zeilen) oder daneben (andere Spalten, Tabellen nebeneinander).
        if len(data) < 2 or any(top <= b and bottom >= a and span.start < cols.stop and cols.start < span.stop
                                for a, b, cols in taken):
            continue
        taken.append((top, bottom, span))
        extra = _built(rows, (row, names, span, data))
        extra.across = _across(sheet.cells, extra)
        found.append(extra)
        if len(found) >= MAX_BLOCKS:
            break
    return found


# Bis zu so vielen Zeilen kann ein Blatt ein Zeilenmodell sein; größere sind Listen von Fällen.
ACROSS_ROWS = 2000


def _across(cells: dict[tuple[int, int], Cell], plan: Layout) -> bool:
    """Ein Zeilenmodell wie ein Finanzplan: Jede Zeile ist eine eigene Größe (Umsatz, Marge in %, Steuern), die
    Spalten sind Jahre oder Monate. Die Formeln sind entlang der Zeile gleich (=C18*(1+D6) in jedem Jahr) und von
    Zeile zu Zeile verschieden. Dann sind die Spalten keine Merkmale von Fällen: Ausreißer, fehlende Werte oder
    eingetippte Zahlen je Spalte sagen dort nichts, und eine SUMME in der Spalte fasst einen Abschnitt zusammen."""
    from collections import Counter

    from .zeilenmodell import kind

    if plan.header_row is None or len(plan.headers) < 4 or plan.last_row - plan.header_row > ACROSS_ROWS:
        return False
    values = set(plan.headers) - {plan.name_column}
    # Die Spalten sind Zeiträume derselben Art (2025A, 2026E …; Jan, Feb …; Budget 2026 …), keine Merkmale wie Alter
    # und Kosten.
    common, count = Counter(kind(plan.headers[c]) for c in values).most_common(1)[0]
    if count < 3 or count < 0.6 * len(values) or not ("9" in common or "m" in common.split()):
        return False
    along = same_along = down = same_down = 0
    rows, numbers = set(), set()
    for (r, c), cell in cells.items():
        if c not in values or not plan.in_table(r, c):
            continue
        if cell.formula or _as_number(cell.value) is not None:
            numbers.add(r)
        if not cell.formula:
            continue
        rows.add(r)
        left = cells.get((r, c - 1))
        if left is not None and left.formula and c - 1 in values:
            along += 1
            same_along += pattern(cell.formula, (r, c), "Zeile") == pattern(left.formula, (r, c - 1), "Zeile")
        up = cells.get((r - 1, c))
        if up is not None and up.formula and plan.in_table(r - 1, c):
            down += 1
            same_down += pattern(cell.formula, (r, c), "Spalte") == pattern(up.formula, (r - 1, c), "Spalte")
    # Eine Fallliste mit Summe und Mittelwert darunter rechnet auch entlang der Zeile gleich, aber fast alle ihre
    # Zeilen sind Daten: In einem Finanzplan rechnet ein guter Teil der Zeilen.
    return (len(rows) >= 2 and len(rows) >= 0.3 * len(numbers) and along >= 6 and same_along >= 0.7 * along
            and same_down <= 0.3 * max(down, 1))


def _unique(names: list[tuple[int, str]]) -> dict[int, str]:
    result: dict[int, str] = {}
    seen: list[str] = []
    for col, name in names:
        name = name.replace(":", "").strip() or f"Spalte {column_letters(col)}"
        result[col] = name if name not in seen else f"{name} {seen.count(name) + 1}"
        seen.append(name)
    return result


def _finish(result: Layout, rows: dict[int, dict[int, Cell]], data: list[int]) -> Layout:
    """Spalte der Zeilennamen: die erste, in der in den meisten Zeilen etwas steht (keine Formel)."""
    if not data:
        return result
    result.last_row = data[-1]
    for col in sorted(result.headers):
        typed = [r for r in data if col in rows.get(r, {}) and rows[r][col].formula is None and rows[r][col].value]
        if len(typed) * 2 >= len(data):
            result.name_column = col
            break
    if result.name_column is None:
        result.name_column = min(result.headers)
    for row in data:
        cell = rows.get(row, {}).get(result.name_column)
        if cell is not None and cell.value:
            result.names[row] = cell.value
    return result


def _table_layout(table: ExcelTable, rows: dict[int, dict[int, Cell]]) -> Layout:
    top, left, bottom, _ = table.area
    result = Layout(header_row=top, table=table.name)
    result.headers = _unique([(left + i, name) for i, name in enumerate(table.columns)])
    data = [r for r in range(top + 1, bottom - table.totals + 1) if any(c in result.headers for c in rows.get(r, {}))]
    _finish(result, rows, data)
    result.last_row = max(result.last_row, bottom - table.totals)
    return result


# Ein Wort vor durchgezählten Perioden: „Jahr | 0 | 1 | 2 …“ ist die Kopfzeile eines Finanzplans.
_PERIOD_WORDS = {"jahr", "jahre", "periode", "perioden", "period", "periods", "year", "years", "t", "monat", "monate",
                 "month", "months", "quartal", "quarter", "zeitpunkt"}


def _period_names(row: dict[int, Cell]) -> dict[int, str]:
    """„Jahr | 0 | 1 | 2 | 3“: die durchgezählten Perioden einer Kopfzeile als „Jahr 0“, „Jahr 1“ … (auch als Formel
    =B12+1). Nur mit dem Wort davor, mindestens drei nebeneinander."""
    word = next((cell.value.strip().rstrip(":") for _, cell in sorted(row.items()) if cell.formula is None
                 and _is_word(cell.value) and cell.value.strip().rstrip(":").casefold() in _PERIOD_WORDS), None)
    if word is None:
        return {}
    numbers = {}
    for c, cell in row.items():
        value = _as_number(cell.value)
        if value is not None and value == int(value) and 0 <= value <= 400 and not _YEAR.match(cell.value.strip()):
            numbers[c] = int(value)
    best: list[int] = []
    for c in sorted(numbers):
        run = [c]
        while run[-1] + 1 in numbers and numbers[run[-1] + 1] == numbers[run[-1]] + 1:
            run.append(run[-1] + 1)
        if len(run) > len(best):
            best = run
    if len(best) < 3:
        return {}
    return {c: f"{word} {numbers[c]}" for c in best}


def _header_names(row: dict[int, Cell]) -> dict[int, str] | None:
    """Die Namen einer möglichen Kopfzeile: fast nur Wörter (oder Jahreszahlen), keine Formeln."""
    names = {c: cell.value for c, cell in row.items() if cell.formula is None and (_is_word(cell.value) or _YEAR.match(cell.value))}
    names.update(_period_names(row))
    words = sum(1 for value in names.values() if _is_word(value))
    if len(names) >= 2 and words >= 1 and len(names) * 3 >= len(row) * 2:
        return names
    return None


def _block(rows: dict[int, dict[int, Cell]], header: int, span: range) -> tuple[list[int], int]:
    """Die Datenzeilen unter einer möglichen Kopfzeile und wie viele Zellen darin unter ihr stehen.

    Bis zu zwei Leerzeilen nach der Kopfzeile und eine mitten in den Daten gehören noch dazu."""
    data: list[int] = []
    score, gap, row, last, numeric = 0, 0, header + 1, max(rows), 0
    while row <= last:
        hits = sum(1 for c in rows.get(row, {}) if c in span)
        if hits and numeric * 2 > len(data) and _new_section(rows, row, span, gap):
            break
        if hits and gap and _type_break(rows, data, row, span):
            break
        if hits:
            data.append(row)
            score += hits
            gap = 0
            numeric += any(c in span and (cell.formula is not None or _as_number(cell.value) is not None)
                           for c, cell in rows[row].items())
        else:
            gap += 1
            # Nach der Summenzeile und einer Leerzeile fängt etwas anderes an, wenn es schmaler ist (Kennzahlen darunter).
            if gap >= (2 if data else 3) or data and _closing(rows, data[-1], row + 1, last, span):
                break
        row += 1
    return data, score


def _type_break(rows: dict[int, dict[int, Cell]], data: list[int], row: int, span: range) -> bool:
    """Nach einer Leerzeile steht in einer Spalte, die bisher in jeder Zeile eine Zahl hatte, in dieser und der
    nächsten Zeile Text („Stückpreis 19,90 …“, darunter „Land | Region“, „DACH“): Dort fängt etwas anderes an."""
    if len(data) < 3:
        return False

    def number(cell) -> bool:
        return cell is not None and (cell.formula is not None or _as_number(cell.value) is not None)

    def word(cell) -> bool:
        return cell is not None and cell.formula is None and _as_number(cell.value) is None and _is_word(cell.value)

    # Ein Abschnittstitel allein in seiner Zeile („2. KENNZAHLEN“, „PIVOT-QUELLE“): Danach kommt Neues.
    line = rows.get(row, {})
    if len(line) == 1:
        text = next(iter(line.values()))
        value = text.value.strip() if text.formula is None and isinstance(text.value, str) else ""
        letters = [ch for ch in value if ch.isalpha()]
        titled = re.match(r"^\d+(\.\d+)*\.?\s+\S", value) or len(letters) >= 4 and all(ch.isupper() for ch in letters)
        # Im Finanzmodell laufen die Jahre über alle Abschnitte weiter: Nur wenn darunter andere Spalten gefüllt sind
        # als in der Tabelle, endet sie.
        shapes = {frozenset(c for c in rows[r] if c in span) for r in data}
        nxt = next((r for r in range(row + 1, row + 4) if rows.get(r)), None)
        if titled and nxt is not None and frozenset(c for c in rows[nxt] if c in span) not in shapes:
            return True
    for c in span:
        cells = [rows[r].get(c) for r in data]
        ahead = [rows.get(r, {}).get(c) for r in (row, row + 1)]
        if all(number(cell) for cell in cells) and all(word(cell) for cell in ahead):
            return True
    return False


def _wordy(row: dict[int, Cell], span: range) -> dict[int, str] | None:
    """Eine Zeile aus lauter Wörtern unter der Tabelle („Jahr | Neu-Ausfallanteil | Barwert“): die Kopfzeile einer
    weiteren Tabelle, wenn mindestens zwei ihrer Spalten so beschriftet sind und keine Zahl und keine Formel darin steht."""
    inside = {c: cell for c, cell in row.items() if c in span}
    if len(inside) < 2 or any(cell.formula is not None or not _is_word(cell.value) for cell in inside.values()):
        return None
    return {c: cell.value for c, cell in inside.items()}


def _new_section(rows: dict[int, dict[int, Cell]], row: int, span: range, gap: int) -> bool:
    """Fängt bei row eine neue Tabelle an? Mehrere Tabellen untereinander auf einem Blatt (Übergangsmatrix,
    Startverteilung, Verteilung nach Jahren …) gehören nicht zusammen: Unter einer Tabelle mit Zahlen steht ein Titel
    („Startverteilung“) über einer Kopfzeile, oder nach einer Leerzeile gleich eine neue Kopfzeile. Eine Liste aus
    lauter Wörtern (Namen, Orte) hört so nie auf: _block fragt nur, wenn die meisten Zeilen bisher Zahlen haben."""
    here = {c: cell for c, cell in rows.get(row, {}).items() if c in span}
    # Ein Titel darf lang sein („Portfolioverteilung nach Jahren (Potenzen der Übergangsmatrix)“).
    title = len(here) == 1 and all(cell.formula is None and _as_number(cell.value) is None
                                   and any(ch.isalpha() for ch in cell.value) for cell in here.values())
    if title and _heads_numbers(rows, row + 1, span):
        return True
    return gap >= 1 and _heads_numbers(rows, row, span)


def _heads_numbers(rows: dict[int, dict[int, Cell]], row: int, span: range) -> bool:
    """Ist row eine Kopfzeile aus Wörtern über Spalten mit Zahlen? Eine Abschnittszeile im Finanzmodell
    („D. Darlehen | Einheit“, darunter „Zinsen | TEUR | 504 …“) beschriftet nur die Spalten der Namen und Einheiten:
    Die Jahre darüber gelten weiter, eine neue Tabelle fängt dort nicht an."""
    names = _wordy(rows.get(row, {}), span)
    if names is None:
        return False
    below = [rows[r] for r in sorted(rows) if r > row][:3]
    return any(any(c in line and (line[c].formula is not None or _as_number(line[c].value) is not None) for line in below)
               for c in names)


_CLOSING = ("summe", "total", "gesamt", "insgesamt")


def _closing(rows: dict[int, dict[int, Cell]], total: int, after: int, last: int, span: range) -> bool:
    """Die Zeile total ist eine Summenzeile (ihr erster Text heißt Summe, Total, Gesamt, auch „Gesamtsumme“), und
    die nächste Zeile ab after füllt weniger Spalten der Tabelle als sie."""
    row = rows.get(total, {})
    texts = [row[c].value for c in sorted(row) if span.start - 1 <= c < span.stop and row[c].formula is None and row[c].value.strip()]
    if not texts or not (texts[0].strip().casefold().startswith(_CLOSING) or woerter.starts_with(texts[0], woerter.TOTAL)):
        return False
    following = after
    while following <= last and not any(c in span for c in rows.get(following, {})):
        following += 1
    if following > last:
        return False
    return sum(1 for c in rows[following] if c in span) < sum(1 for c in row if c in span)


def _guess_layout(rows: dict[int, dict[int, Cell]]) -> Layout:
    """Ohne Excel-Tabelle: die Kopfzeile, unter der die meisten Daten stehen.

    Ein Titel darüber („Wahlgang | Sitze(577)“) verliert: Unter ihm steht zuerst eine Zeile
    aus lauter Wörtern, die eigentliche Kopfzeile, und darunter Zahlen."""
    if not rows:
        return Layout()
    found = _candidates(rows, HEADER_ROWS)
    if not found:
        return Layout()
    return _built(rows, max(found, key=lambda c: c[0])[1:])


def _candidates(rows: dict[int, dict[int, Cell]], limit: int) -> list[tuple]:
    """Jede mögliche Kopfzeile unter den ersten limit Zeilen mit ihren Daten: (Punkte, Zeile, Namen, Spalten, Daten)."""
    found: list[tuple] = []
    seen: list[tuple[range, list[int]]] = []
    small = sum(len(r) for r in rows.values()) <= BLOCK_CELLS
    for row in sorted(rows)[:limit]:
        names = _header_names(rows[row])
        if names is None:
            continue
        # Eine Datenzeile einer Tabelle („12.01.2026 | Nord | Dock | 1234,5“) ist keine Kopfzeile, auch wenn rechts
        # daneben die Köpfe einer Übersicht stehen: Dann zählen nur die Wörter neben der Tabelle.
        over = _header_names(rows.get(row - 1, {}))
        tables = [range(min(over), max(over) + 1)] if over else []
        tables += [span for span, data in seen if row in data]
        for above in tables:
            if any(c not in names for c in rows[row] if c in above):
                names = {c: v for c, v in names.items() if c not in above}
        if len(names) < 2:
            continue
        for group in _side_by_side(rows, row, names, small):
            candidate = _candidate(rows, row, group, seen)
            if candidate is not None:
                found.append(candidate)
    return found


def _candidate(rows: dict[int, dict[int, Cell]], row: int, names: dict[int, str], seen: list) -> tuple | None:
    """Eine mögliche Kopfzeile mit ihren Daten (Punkte, Zeile, Namen, Spalten, Daten), oder None für einen Titel."""
    span = range(min(names), max(names) + 1)
    data, score = _block(rows, row, span)
    if not data:
        return None
    seen.append((span, data))
    sub = _subheader(rows, row, names, span, data)
    if sub is not None:
        return score, row + 1, sub, span, data[1:]
    rest = data[1:]
    # Nur unter dieser Kopfzeile (und eine Spalte links davon für „Filiale“) zählt, ob es wie eine Kopfzeile aussieht.
    inside = lambda d: {c: cell for c, cell in rows[d].items() if span.start - 1 <= c < span.stop}  # noqa: E731
    first = _header_names(inside(data[0]))
    # Eine erste Datenzeile mit einer leeren Zahl („10.01.2026 | West | Monitor | “) sieht auch aus wie eine
    # Kopfzeile. Ihre Wörter stehen aber weiter unten in derselben Spalte wieder, die einer Kopfzeile nicht.
    again = first and sum(any((cell := rows[d].get(c)) is not None and cell.value == v for d in rest) for c, v in first.items())
    # Steht in der ersten Zeile schon eine Zahl, wo in den meisten Zeilen darunter auch eine steht, ist sie eine
    # Datenzeile („Januar | 412“), auch wenn rechts daneben ein Wort steht („Summe mit Zins“).
    numbers = first is not None and any(
        cell.formula is None and not _YEAR.match(cell.value) and _as_number(cell.value) is not None
        and sum((other := rows[d].get(c)) is not None and _as_number(other.value) is not None for d in rest) * 2 >= len(rest) > 0
        for c, cell in inside(data[0]).items())
    if first is not None and not numbers and again * 2 < len(first) and \
            sum(_header_names(inside(d)) is None for d in rest) * 2 > len(rest):
        # Ein Titel: Darunter sieht nur die erste Zeile wie eine Kopfzeile aus, die meisten anderen nicht.
        return None
    return score, row, names, span, data


def _side_by_side(rows: dict[int, dict[int, Cell]], row: int, names: dict[int, str], small: bool) -> list[dict[int, str]]:
    """Zwei Tabellen nebeneinander, getrennt durch eine leere Spalte („Monat | Umsatz | Kosten“ und rechts davon
    „Filiale | Mitarbeitende | Umsatz“): je eine Kopfzeile für sich. Ein einzelnes Wort weit rechts daneben
    („MwSt | 0,081“, zwei leere Spalten dazwischen) ist eine Annahme neben der Tabelle, keine Spalte von ihr. Eine
    einzelne leere Spalte neben nur einem Kopf („Name | | Wert“) ist bloß Abstand. Nur auf kleinen Blättern: In einer
    großen Fallliste bleibt es eine Tabelle, wie bisher."""
    if not small:
        return [names]
    span = range(min(names), max(names) + 1)
    data, _ = _block(rows, row, span)
    used = {c for d in [row, *data] for c in rows.get(d, {}) if c in span}
    cols = sorted(names)
    groups: list[list[int]] = [[cols[0]]]
    gaps: list[int] = []
    for a, b in zip(cols, cols[1:]):
        empty = [c for c in range(a + 1, b) if c not in used]
        if empty:
            groups.append([b])
            gaps.append(len(empty))
        else:
            groups[-1].append(b)
    if len(groups) == 1:
        return [names]
    # Kleine Stücke (ein Kopf) gehören bei einer leeren Spalte zum Nachbarn, bei mehr Abstand fallen sie weg.
    merged: list[list[int]] = [groups[0]]
    for gap, group in zip(gaps, groups[1:]):
        if gap == 1 and (len(merged[-1]) < 2 or len(group) < 2):
            merged[-1] = merged[-1] + group
        else:
            merged.append(group)
    out = [{c: names[c] for c in group} for group in merged if len(group) >= 2]
    return out or [names]


def _subheader(rows: dict[int, dict[int, Cell]], row: int, names: dict[int, str], span: range,
               data: list[int]) -> dict[int, str] | None:
    """Ein zweizeiliger Kopf: „Umsatz“ über zwei Spalten (verbunden) und darunter „2025 | 2026“. Die zweite Zeile
    gehört zum Kopf, wenn in ihr nur Wörter und Jahreszahlen stehen, mindestens eins davon unter einer leeren Zelle der
    ersten (dort ist die erste verbunden oder reicht nach rechts), ihre Werte weiter unten nicht wieder vorkommen und
    darunter Zahlen stehen. Die Spalten heißen dann „Umsatz 2025“, „Umsatz 2026“, „Kosten 2025“ …"""
    if len(data) < 3 or data[0] != row + 1:
        return None
    below = {c: cell for c, cell in rows[row + 1].items() if c in span}
    if len(below) < 2 or any(cell.formula is not None or not (_is_word(cell.value) or _YEAR.match(cell.value.strip()))
                             for cell in below.values()):
        return None
    if all(c in names for c in below):
        return None
    rest = data[1:]
    if any((cell := rows[d].get(c)) is not None and cell.value.strip() == below[c].value.strip() for d in rest for c in below):
        return None
    numeric = sum(any((cell := rows[d].get(c)) is not None and (cell.formula is not None or _as_number(cell.value) is not None)
                      for c in below) for d in rest)
    if numeric * 2 < len(rest):
        return None
    combined: dict[int, str] = {}
    current = ""
    for c in span:
        lower = below[c].value.strip() if c in below else ""
        if c in names:
            current = names[c].strip()
            combined[c] = f"{current} {lower}" if lower else current
            # Steht unter dem Wort ein eigener Spaltenname („1. ANNAHMEN“ über „Position“), ist es ein Titel dieser
            # Spalte und reicht nicht über die Jahre rechts davon.
            if lower and not _YEAR.match(lower):
                current = ""
        elif lower:
            combined[c] = f"{current} {lower}" if current else lower
        else:
            current = ""
    return combined


def _built(rows: dict[int, dict[int, Cell]], best: tuple) -> Layout:
    row, names, span, data = best
    used = {c for d in data for c in rows[d] if c in span}
    left = min(names) - 1
    labels = [d for d in data if left in rows[d] and rows[d][left].formula is None and _is_word(rows[d][left].value)]
    if left >= 1 and left not in rows[row] and len(labels) * 2 >= len(data):
        # Zeilenbeschriftungen ohne eigenen Spaltenkopf („Seeking Employment | 532 | 35“).
        used.add(left)
    # Zweizeilige Kopfzeile: „Percent“ über „of Class“ heißt „Percent of Class“.
    over = _header_names(rows.get(row - 1, {})) or {}
    if len([c for c in names if c in over]) * 3 < len(names) * 2:
        over = {}
    columns = []
    for col in sorted(set(names) | used):
        name = names.get(col) or _above(rows, row, col) or f"Spalte {column_letters(col)}"
        if col in names and col in over and not name.startswith(over[col].strip()):
            name = f"{over[col]} {name}"
        columns.append((col, name))
    result = Layout(header_row=row, headers=_unique(columns))
    return _finish(result, rows, data)


def _above(rows: dict[int, dict[int, Cell]], row: int, col: int) -> str:
    """Ein Spaltenname eine oder zwei Zeilen über einer leeren Zelle der Kopfzeile."""
    for above in (row - 1, row - 2):
        cell = rows.get(above, {}).get(col)
        if cell is not None and cell.formula is None and _is_word(cell.value):
            return cell.value
    return ""


def _left_label(cells: dict[tuple[int, int], Cell], row: int, col: int) -> str:
    """Wie eine einzelne Zelle außerhalb der Tabelle heißt: das Wort links daneben oder darüber."""
    for left in range(col - 1, max(0, col - 3), -1):
        cell = cells.get((row, left))
        if cell is not None and cell.formula is None and _is_word(cell.value):
            return cell.value.replace(":", "").strip()
    cell = cells.get((row - 1, col))
    if cell is not None and cell.formula is None and _is_word(cell.value):
        return cell.value.replace(":", "").strip()
    return ""


# --- Formeln vorlesen ---


def _unquote(name: str) -> str:
    return re.sub(r"'(.)", r"\1", name).strip()


def _structured(match: re.Match) -> str:
    """Tabelle1[@Betrag] → „Betrag“, Tabelle1[Betrag] → „alle Betrag“, Tabelle1[[#Totals],[Betrag]] → „Ergebnis Betrag“."""
    table, inner = match.group(1) or "", match.group(2).strip()
    if inner.isdigit():
        return match.group(0)  # Bezug auf eine andere Mappe: [1]Blatt!A1
    this_row = inner.startswith("@")
    rest = inner[1:].strip() if this_row else inner
    parts = [p.strip() for p in re.findall(r"\[([^\[\]]*)\]", rest)] or ([rest] if rest else [])
    special = {p.lower() for p in parts if p.startswith("#")}
    this_row = this_row or "#this row" in special
    columns = [_unquote(p) for p in parts if p and not p.startswith("#")]
    column = (T(f"{columns[0]} bis {columns[1]}", f"{columns[0]} to {columns[1]}") if len(columns) == 2 and ":" in rest
              else (columns[0] if columns else ""))
    if english():
        name = table.replace("\\", "")
        of = f" of {name}" if name else " of the table"
        if this_row:
            return column or f"this row{of}"
        if "#totals" in special:
            return f"total {column}" if column else f"total row{of}"
        if "#headers" in special:
            return f"header {column}" if column else f"header row{of}"
        return f"all {column}" if column else (f"table {name}" if name else "the table")
    name = table.replace("\\", "") or "der Tabelle"
    if this_row:
        return column or f"diese Zeile von {name}"
    if "#totals" in special:
        return f"Ergebnis {column}" if column else f"Ergebniszeile von {name}"
    if "#headers" in special:
        return f"Kopf {column}" if column else f"Kopfzeile von {name}"
    return f"alle {column}" if column else f"Tabelle {name}"


def _outside_strings(formula: str, change) -> str:
    """Nur außerhalb von "Texten" in der Formel ändern."""
    return "".join(part if part.startswith('"') else change(part) for part in STRING.split(formula))


_TRUTH = re.compile(r"(?<![\w.$!'])(TRUE|FALSE)(?![\w(!])", re.IGNORECASE)


def readable(formula: str, here: tuple[int, int], plan: Layout, label=None) -> str:
    """„B3*577/100“ → „=Wähleranteil(%) * 577 / 100“: Spaltenköpfe statt Adressen, deutsche Namen.
    label(pos): wie eine einzelne Zelle außerhalb der Tabelle heißt („MwSt“ für die Annahme in I1), sonst ""."""
    names: list[str] = []
    to = T(" bis ", " to ")

    def hold(name: str) -> str:
        # Bis zum Schluss ein Platzhalter: In „E-Mail“ oder „Q1“ soll nichts mehr ersetzt werden.
        names.append(name)
        return f"\x00{len(names) - 1}\x00"

    def cell_range(match: re.Match) -> str:
        a, b = (_ADDRESS.match(side.replace("$", "")) for side in match.groups())
        c1, r1, c2, r2 = column_number(a.group(1)), int(a.group(2)), column_number(b.group(1)), int(b.group(2))
        if plan.across and plan.in_table(r1, c1) and plan.in_table(r2, c2) and r1 in plan.names and r2 in plan.names \
                and not (like_cell(plan.names[r1]) or like_cell(plan.names[r2])):
            # In einem Finanzplan heißt eine Zeile nach ihrer Größe, eine Spalte nach ihrem Jahr: „Jahresüberschuss bis
            # Veränderung Working Capital“ im selben Jahr, „Umsatz 2026E bis 2030E“ über die Jahre.
            if c1 == c2:
                year = "" if c1 == here[1] else f" {plan.headers[c1]}"
                return hold(f"{plan.names[r1]}{to}{plan.names[r2]}{year}")
            if r1 == r2:
                return hold(f"{plan.names[r1]} {plan.headers[c1]}{to}{plan.headers[c2]}")
        if c1 == c2 and plan.in_table(r1, c1) and plan.in_table(r2, c2):
            if r1 <= plan.header_row + 1 and r2 >= plan.last_row:
                return hold(T(f"alle {plan.headers[c1]}", f"all {plan.headers[c1]}"))
            # „Kosten CHF F240000 bis F242397“ läse man als Zellen: Dann bleibt der Bereich, wie er ist.
            if r1 in plan.names and r2 in plan.names and not (like_cell(plan.names[r1]) or like_cell(plan.names[r2])):
                return hold(f"{plan.headers[c1]} {plan.counted(r1)}{to}{plan.names[r2]}")
        # Ein Bereich bleibt ganz: In $F$2:F2400 ist F2400 kein Spaltenkopf („$F$2:Kosten CHF“).
        return hold(match.group(0))

    def single(match: re.Match) -> str:
        col, row = column_number(match.group(2)), int(match.group(4))
        if plan.in_table(row, col):
            if plan.across and row in plan.names and not like_cell(plan.names[row]) and (row, col) != here:
                # „Umsatz“ im selben Jahr, „Umsatz 2027E“ im Vorjahr.
                return hold(plan.names[row] if col == here[1] else f"{plan.names[row]} {plan.headers[col]}")
            if row == here[0]:
                return hold(plan.headers[col])
            if row in plan.names and not like_cell(plan.names[row]):
                # „D (Ausfall) Jahr 4“ statt „D (Ausfall) 4“, wenn die Zeile nur eine Zahl als Namen hat. Die Spalte
                # der Namen selbst ist nur „Jahr 3“, nicht „Jahr Jahr 3“.
                if col == plan.name_column:
                    return hold(plan.counted(row))
                return hold(f"{plan.headers[col]} {plan.counted(row)}")
        elif label is not None and (name := label((row, col))):
            # Eine Annahme neben der Tabelle: „MwSt ($I$1)“.
            return hold(f"{name} ({match.group(0)})")
        return match.group(0)

    def change(part: str) -> str:
        part = STRUCTURED.sub(lambda m: hold(_structured(m)), part)
        part = _OTHER_SHEET.sub(lambda m: hold(m.group(0)), part)
        part = _REF.sub(single, _RANGE.sub(cell_range, part))
        part = _FUNCTION.sub(lambda m: f"{FUNCTIONS.get(m.group(1), m.group(1))}(", part)
        part = _TRUTH.sub(lambda m: "WAHR" if m.group(1).upper() == "TRUE" else "FALSCH", part)
        part = _DECIMAL.sub(r"\1,\2", part.replace(",", "; "))
        return re.sub(r"\s*([*/+^&])\s*|(?<=[\w)\x00])\s*-\s*", lambda m: f" {m.group(1) or '-'} ", part)

    text = " ".join(_outside_strings(formula.lstrip("="), change).split())
    while "\x00" in text:
        text = re.sub(r"\x00(\d+)\x00", lambda m: names[int(m.group(1))], text)
    text = text if len(text) < MAX_FORMULA else text[: MAX_FORMULA - 2].rstrip() + "…"
    return f"={text}"


def pattern(formula: str, here: tuple[int, int], along: str = "") -> str:
    """Die Formel relativ zu ihrer Zelle, wie Excel sie herunterzieht (Z1S1-Schreibweise).
    along="Spalte": für Formeln untereinander. Dort zeigt B7 und $B7 auf dieselbe Spalte, das $ ändert nichts,
    darum zählt die Spalte dann immer fest. along="Zeile" ebenso für Formeln nebeneinander."""
    return formel.remember("\x00pattern", (formula, here, along), lambda: _pattern(formula, here, along))


def _pattern(formula: str, here: tuple[int, int], along: str = "") -> str:
    def relative(match: re.Match) -> str:
        dollar_col, letters, dollar_row, digits = match.groups()
        col = f"S{column_number(letters)}" if dollar_col or along == "Spalte" else f"S[{column_number(letters) - here[1]}]"
        row = f"Z{digits}" if dollar_row or along == "Zeile" else f"Z[{int(digits) - here[0]}]"
        return row + col

    return outside_literals(formula.replace(" ", ""), lambda part: ANY_REF.sub(relative, part))


# --- Aufbau der Mappe: Filter, Diagramme, Pivot-Tabellen … ---


@dataclass(frozen=True)
class Part:
    """Ein Teil der Mappe, das kein Zellwert ist, mit dem, was er gerade tut."""

    key: str
    label: str
    text: str
    sheet: str
    # Die Spalte, um die es geht (für das Schwärzen), sonst „“.
    column: str = ""


_COMPARE = {
    "equal": "gleich", "notEqual": "ungleich", "greaterThan": "größer als", "greaterThanOrEqual": "mindestens",
    "lessThan": "kleiner als", "lessThanOrEqual": "höchstens",
}
_PERIODS = {
    "today": "heute", "yesterday": "gestern", "tomorrow": "morgen", "last7Days": "in den letzten 7 Tagen",
    "thisWeek": "diese Woche", "lastWeek": "letzte Woche", "nextWeek": "nächste Woche", "thisMonth": "diesen Monat",
    "lastMonth": "letzten Monat", "nextMonth": "nächsten Monat",
}
_ICONS = (
    ("TrafficLights", "Ampeln"), ("Arrows", "Pfeile"), ("Flags", "Flaggen"), ("Stars", "Sterne"), ("Rating", "Balken"),
    ("Signs", "Zeichen"), ("Symbols", "Symbole"), ("Quarters", "Viertel"), ("Triangles", "Dreiecke"), ("Boxes", "Kästchen"),
)
_VALIDATION = {
    "whole": "ganze Zahl", "decimal": "Zahl", "date": "Datum", "time": "Uhrzeit", "textLength": "Textlänge",
}
# Dasselbe auf Englisch, wie das englische Excel es nennt.
_COMPARE_EN = {
    "equal": "equal to", "notEqual": "not equal to", "greaterThan": "greater than", "greaterThanOrEqual": "at least",
    "lessThan": "less than", "lessThanOrEqual": "at most",
}
_PERIODS_EN = {
    "today": "today", "yesterday": "yesterday", "tomorrow": "tomorrow", "last7Days": "in the last 7 days",
    "thisWeek": "this week", "lastWeek": "last week", "nextWeek": "next week", "thisMonth": "this month",
    "lastMonth": "last month", "nextMonth": "next month",
}
_ICONS_EN = {
    "TrafficLights": "traffic lights", "Arrows": "arrows", "Flags": "flags", "Stars": "stars", "Rating": "bars",
    "Signs": "signs", "Symbols": "symbols", "Quarters": "quarters", "Triangles": "triangles", "Boxes": "boxes",
}
_VALIDATION_EN = {
    "whole": "whole number", "decimal": "number", "date": "date", "time": "time", "textLength": "text length",
}


def _compare(operator: str) -> str:
    return T(_COMPARE.get(operator, operator), _COMPARE_EN.get(operator, operator))


class _Book:
    """Was zum Beschreiben gebraucht wird: alle Blätter, ihre Kopfzeilen und die Namen."""

    def __init__(self, sheets: list[Sheet], plans: dict[str, Layout], names: dict[str, str]) -> None:
        self.sheets = {s.name: s for s in sheets}
        self.plans = plans
        self.names = names

    def header(self, sheet: str, row: int, col: int) -> str:
        """Name einer Spalte: aus einer Excel-Tabelle, aus der Kopfzeile, sonst ihr Buchstabe."""
        for table in self.sheets[sheet].tables if sheet in self.sheets else ():
            top, left, bottom, right = table.area
            if left <= col <= right and top <= row <= bottom and col - left < len(table.columns):
                return table.columns[col - left]
        plan = self.plans.get(sheet)
        # Unter der Tabelle (eine zweite Tabelle weiter unten) gilt ihr Spaltenkopf nicht mehr.
        if plan is not None and col in plan.headers and (plan.header_row or 0) <= row <= plan.last_row + 1:
            return plan.headers[col]
        return T(f"Spalte {column_letters(col)}", f"Column {column_letters(col)}")

    def area_name(self, sheet: str, area: Area) -> str:
        top, left, bottom, right = area
        if (top, left) == (bottom, right):
            plan = self.plans.get(sheet)
            if plan is not None and plan.in_table(top, left):
                return plan.headers[left]
            cells = self.sheets[sheet].cells if sheet in self.sheets else {}
            label = _left_label(cells, top, left)
            return f"{label} ({area_text(area)})" if label else area_text(area)
        first, last = self.header(sheet, top + 1, left), self.header(sheet, top + 1, right)
        if left == right and first.startswith(T("Spalte ", "Column ")):
            # Ein Bereich außerhalb der Tabelle heißt nach dem Wort darüber („Monat“ über A21:A32).
            above = self.sheets[sheet].cells.get((top - 1, left)) if sheet in self.sheets else None
            if above is not None and above.formula is None and _is_word(above.value):
                return above.value
        return first if left == right else T(f"{first} bis {last}", f"{first} to {last}")

    def ref_name(self, sheet: str, ref: str) -> str:
        """„Tabelle1!$D$6:$D$10“ → „Anzahl Sitze“, auch über Namen und Excel-Tabellen."""
        ref = ref.strip().lstrip("=")
        if ref in self.names:
            ref = self.names[ref]
        structured = STRUCTURED.fullmatch(ref)
        if structured:
            return _structured(structured)
        area = area_of(ref)
        if area is None:
            return ref
        return self.area_name(sheet_of(ref) or sheet, area)

    def values(self, sheet: str, ref: str, limit: int = 8) -> list[str]:
        """Die Werte eines Bereichs, z. B. für eine Auswahlliste aus Zellen."""
        ref = ref.strip().lstrip("=")
        if ref in self.names:
            ref = self.names[ref].lstrip("=")
        structured = STRUCTURED.fullmatch(ref)
        area = None
        if structured:
            sheet, area = self._table_column(structured)
        else:
            sheet = sheet_of(ref) or sheet
            area = area_of(ref)
        if area is None or sheet not in self.sheets:
            return []
        cells = self.sheets[sheet].cells
        found = []
        for row in range(area[0], area[2] + 1):
            for col in range(area[1], area[3] + 1):
                cell = cells.get((row, col))
                if cell is not None and cell.value and cell.value not in found:
                    found.append(cell.value)
                if len(found) > limit:
                    return found
        return found

    def _table_column(self, match: re.Match) -> tuple[str, Area | None]:
        name = (match.group(1) or "").replace("\\", "")
        column = _unquote(match.group(2).strip("[]@ "))
        for sheet in self.sheets.values():
            for table in sheet.tables:
                if table.name.replace("\\", "") == name and column in table.columns:
                    col = table.area[1] + table.columns.index(column)
                    return sheet.name, (table.area[0] + 1, col, table.area[2] - table.totals, col)
        return "", None


def _plain(formula: str, sheet: str, area: Area, book: _Book) -> str:
    """Eine Formel in einer Regel: Zahlen und Texte wie sie sind, sonst lesbar."""
    formula = formula.strip()
    if re.fullmatch(r"-?\d+(\.\d+)?", formula):
        return number_text(float(formula))
    if re.fullmatch(r'"(?:[^"]|"")*"', formula):
        return T(f"„{formula[1:-1]}“", f"“{formula[1:-1]}”")
    plan = book.plans.get(sheet) or Layout()
    return readable(formula, (area[0], area[1]), plan).lstrip("=")


def _quoted(values: list[str], limit: int = 8) -> str:
    shown = values[:limit]
    more = T(f" und {len(values) - limit} weitere", f" and {len(values) - limit} more") if len(values) > limit else ""
    return ", ".join(shown) + more


def _between(operator: str, formulas: list[str]) -> str:
    """„zwischen 1 und 5“ oder „nicht zwischen …“."""
    inside = operator != "notBetween"
    return T(f"{'zwischen' if inside else 'nicht zwischen'} {formulas[0]} und {formulas[1]}",
             f"{'between' if inside else 'not between'} {formulas[0]} and {formulas[1]}")


def _format_text(rule: Rule, sheet: str, area: Area, book: _Book) -> str:
    formulas = [_plain(f, sheet, area, book) for f in rule.formulas]
    if rule.kind == "cellIs":
        if rule.operator in ("between", "notBetween") and len(formulas) == 2:
            return _between(rule.operator, formulas)
        return f"{_compare(rule.operator)} {formulas[0] if formulas else ''}".strip()
    if rule.kind == "expression":
        return T(f"wenn {formulas[0]}", f"if {formulas[0]}") if formulas else T("nach Formel", "by formula")
    if rule.kind == "iconSet":
        if english():
            return f"icons ({next((name for key, name in _ICONS_EN.items() if key in rule.extra), rule.extra)})"
        return f"Symbole ({next((name for key, name in _ICONS if key in rule.extra), rule.extra)})"
    if rule.kind == "top10":
        where, count = (rule.extra.split(" ", 1) + ["10"])[:2]
        if count.endswith("%"):
            return T(f"die {'obersten' if where == 'top' else 'untersten'} {count}",
                     f"the {'top' if where == 'top' else 'bottom'} {count}")
        return T(f"die {count} {'größten' if where == 'top' else 'kleinsten'}",
                 f"the {count} {'largest' if where == 'top' else 'smallest'}")
    if english():
        return {
            "colorScale": "color scale", "dataBar": "data bars", "duplicateValues": "duplicate values",
            "uniqueValues": "values that occur only once", "containsBlanks": "empty cells",
            "notContainsBlanks": "non-empty cells", "containsErrors": "errors", "notContainsErrors": "no errors",
            "aboveAverage": "above average" if rule.extra != "below" else "below average",
            "containsText": f"contains “{rule.extra}”", "notContainsText": f"does not contain “{rule.extra}”",
            "beginsWith": f"begins with “{rule.extra}”", "endsWith": f"ends with “{rule.extra}”",
            "timePeriod": f"date {_PERIODS_EN.get(rule.extra, rule.extra)}",
        }.get(rule.kind, rule.kind)
    texts = {
        "colorScale": "Farbskala", "dataBar": "Datenbalken", "duplicateValues": "doppelte Werte",
        "uniqueValues": "Werte, die nur einmal vorkommen", "containsBlanks": "leere Zellen",
        "notContainsBlanks": "nicht leere Zellen", "containsErrors": "Fehler", "notContainsErrors": "keine Fehler",
        "aboveAverage": "über dem Durchschnitt" if rule.extra != "below" else "unter dem Durchschnitt",
        "containsText": f"enthält „{rule.extra}“", "notContainsText": f"enthält nicht „{rule.extra}“",
        "beginsWith": f"beginnt mit „{rule.extra}“", "endsWith": f"endet mit „{rule.extra}“",
        "timePeriod": f"Datum {_PERIODS.get(rule.extra, rule.extra)}",
    }
    return texts.get(rule.kind, rule.kind)


def _validation_text(rule: Rule, sheet: str, area: Area, book: _Book) -> str:
    if rule.kind == "list":
        source = rule.formulas[0] if rule.formulas else ""
        if source.startswith('"'):
            values = [v.strip() for v in source.strip('"').split(",") if v.strip()]
        else:
            values = book.values(sheet, source)
        return _quoted(values) if values else book.ref_name(sheet, source)
    if rule.kind == "custom":
        if not rule.formulas:
            return T("nach Formel", "by formula")
        formula = _plain(rule.formulas[0], sheet, area, book)
        return T(f"nur wenn {formula}", f"only if {formula}")
    what = T(_VALIDATION.get(rule.kind, rule.kind), _VALIDATION_EN.get(rule.kind, rule.kind))
    formulas = [_plain(f, sheet, area, book) for f in rule.formulas]
    if rule.operator in ("", "between", "notBetween") and len(formulas) == 2:
        return f"{what} {_between(rule.operator, formulas)}"
    return f"{what} {_compare(rule.operator)} {formulas[0] if formulas else ''}".strip()


def _sort_text(keys, sheet: str, row: int, book: _Book) -> str:
    parts = []
    for key in keys:
        name = book.header(sheet, row, key.column)
        by = f" ({said(key.by)})" if key.by else ""
        parts.append(T(f"nach {name}{by} {'absteigend' if key.descending else 'aufsteigend'}",
                       f"by {name}{by} {'descending' if key.descending else 'ascending'}"))
    return T(", dann ", ", then ").join(parts)


def _chart_text(chart, sheet: str, book: _Book) -> str:
    text = said(chart.kind) + (T(f" „{chart.title}“", f" “{chart.title}”") if chart.title else "")
    names = []
    for series in chart.series:
        name = series.name or (book.values(sheet, series.name_ref, 1) or [""])[0] if series.name_ref else series.name
        names.append(name or (book.ref_name(sheet, series.values) if series.values else ""))
    names = [n for n in dict.fromkeys(names) if n]
    if names:
        text += T(" mit ", " with ") + _quoted(names, 5)
    categories = next((s.categories for s in chart.series if s.categories), "")
    if categories:
        text += T(" nach ", " by ") + book.ref_name(sheet, categories)
    return text


def _pivot_text(pivot, book: _Book) -> str:
    from .pivotdiagramm import _VALUES_FIELD

    parts = []
    words = {"Zeilen": "rows", "Spalten": "columns", "Werte": "values", "Filter": "filters"}
    for word, items in (("Zeilen", pivot.rows), ("Spalten", pivot.columns), ("Werte", pivot.values), ("Filter", pivot.filters)):
        if word in ("Zeilen", "Spalten") and pivot.values:
            # „Werte“ bei den Spalten (so meldet Excel das Feld „Σ Werte“) ist kein Feld der Daten.
            items = [n for n in items if n.strip().lstrip("Σ∑ ").strip().casefold() not in _VALUES_FIELD]
        if items:
            parts.append(f"{T(word, words[word])} {', '.join(said(i) for i in items)}")
    if pivot.hidden:
        parts.append("; ".join(said(h) for h in pivot.hidden))
    if pivot.source:
        source = pivot.source
        if "!" in source and area_of(source) is not None:
            source = f"{sheet_of(source)} {area_text(area_of(source))}"
        parts.append(T(f"aus {source}", f"from {source}"))
    return "; ".join(parts) or T("leer", "empty")


def parts_of(sheets: list[Sheet], plans: dict[str, Layout], names: dict[str, str] | None = None) -> dict[str, Part]:
    """Alle Teile der Mappe, die keine Zellwerte sind, mit einem Schlüssel zum Vergleichen."""
    book = _Book(sheets, plans, names or {})
    found: dict[str, Part] = {}

    def add(key: str, label: str, text: str, sheet: str, column: str = "") -> None:
        while key in found:
            key += "+"
        found[key] = Part(key, label, text, sheet, column)

    for sheet in sheets:
        if not sheet.parts:
            continue
        name = sheet.name
        for table in sheet.tables:
            columns = _quoted(list(table.columns), 8)
            add(f"table:{name}:{table.name}", T(f"Excel-Tabelle {table.name}", f"Excel table {table.name}"),
                T(f"{area_text(table.area)} mit Spalten {columns}", f"{area_text(table.area)} with columns {columns}"), name)
        for auto in sheet.filters:
            where = auto.table or area_text(auto.area)
            for column in auto.filters:
                header = book.header(name, auto.area[0] + 1, column.column)
                add(f"filter:{name}:{where}:{column.column}", f"Filter {header}", said(column.condition), name, header)
            if auto.sort:
                text = _sort_text(auto.sort, name, auto.area[0] + 1, book)
                add(f"sort:{name}:{where}", T(SORTING, "Sorting"), f"{auto.table + ' ' if auto.table else ''}{text}".strip(), name)
        for rule in sheet.formats:
            area = rule.areas[0]
            header = book.area_name(name, area)
            sqref = " ".join(area_text(a) for a in rule.areas)
            add(f"format:{name}:{sqref}:{rule.kind}:{rule.operator}", T(f"Hervorhebung {header}", f"Highlight {header}"),
                _format_text(rule, name, area, book), name, header)
        for rule in sheet.validations:
            area = rule.areas[0]
            header = book.area_name(name, area)
            label = T("Auswahlliste", "Drop-down list") if rule.kind == "list" else T("Eingaberegel", "Validation rule")
            sqref = " ".join(area_text(a) for a in rule.areas)
            add(f"valid:{name}:{sqref}", f"{label} {header}", _validation_text(rule, name, area, book), name, header)
        for chart in sheet.charts:
            add(f"chart:{name}:{chart.name}", T(f"Diagramm {chart.title or chart.name}", f"Chart {chart.title or chart.name}"),
                _chart_text(chart, name, book), name)
        for pivot in sheet.pivots:
            add(f"pivot:{name}:{pivot.name}", T(f"Pivot-Tabelle {pivot.name}", f"Pivot table {pivot.name}"),
                _pivot_text(pivot, book), name)
    if any(s.parts for s in sheets):
        for name, ref in (names or {}).items():
            add(f"name:{name}", f"Name {name}", _name_text(ref, book), "")
    return found


def _name_text(ref: str, book: _Book) -> str:
    ref = ref.lstrip("=")
    area = area_of(ref)
    sheet = sheet_of(ref)
    if area is not None and sheet in book.sheets:
        text = f"{sheet}!{area_text(area)}"
        if area[:2] == area[2:]:
            cell = book.sheets[sheet].cells.get(area[:2])
            if cell is not None and cell.value:
                text += T(f" (zurzeit {cell.value})", f" (currently {cell.value})")
        return text
    return ref


# --- Änderungen erkennen ---


_BLANK = Cell("")


@dataclass
class _State:
    cells: dict[tuple[int, int], Cell]
    plan: Layout


# So schreibt der Prüfer Zahlen selbst (number_text): schnell erkannt, ohne die Sonderfälle unten.
_PLAIN_NUMBER = re.compile(r"-?\d+(?:,\d+)?")


def _as_number(value: str) -> float | None:
    """Die Zahl in einem Zellwert, ob deutsch (1.234,5), schweizerisch (1'234.5) oder englisch (1,234.5)."""
    if _PLAIN_NUMBER.fullmatch(value):
        return float(value.replace(",", "."))
    text = re.sub(r"[\s'’]", "", value).rstrip("%")
    if "," in text and "." in text:
        thousands = "." if text.rfind(",") > text.rfind(".") else ","
        text = text.replace(thousands, "")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3}){2,}", text):
        text = text.replace(".", "")
    text = text.replace(",", ".")
    if re.fullmatch(r"-?\d+(\.\d+)?", text):
        return float(text)
    return None


def _order(values: list[str]) -> str | None:
    """„aufsteigend“ oder „absteigend“, wenn die Werte so sortiert sind und nicht alle gleich."""
    if len(set(values)) < 2:
        return None
    numbers = [_as_number(v) for v in values]
    keys = numbers if all(n is not None for n in numbers) else [v.casefold() for v in values]
    if all(a <= b for a, b in zip(keys, keys[1:])):
        return "aufsteigend"
    if all(a >= b for a, b in zip(keys, keys[1:])):
        return "absteigend"
    return None


class WorkbookTracker:
    """Merkt sich den letzten Stand der Mappe und meldet, was sich daran entschieden hat."""

    def __init__(self, mask=None) -> None:
        self.name: str = ""
        self._sheets: dict[str, _State] | None = None
        self._parts: dict[str, Part] | None = None
        self._parts_name = ""
        self._sheet_names: tuple[str, ...] | None = None
        self._sheet_names_of = ""
        self.last_row: tuple[str, ...] = ()
        # Das Blatt, an dem gerade gearbeitet wird, für die Übersicht.
        self.active: str = ""
        # Schwärzt einen Wert mit seinem Spaltenkopf (privacy.line), None: nichts schwärzen.
        self._mask = mask
        # Zellen, die in dieser Sitzung frisch getippt wurden, je Blatt: Sie wieder zu ändern ist keine Entscheidung.
        self._fresh: dict[str, set[tuple[int, int]]] = {}

    @property
    def sheets(self) -> int:
        return len(self._sheet_names or self._sheets or {})

    def _masked(self, sheet: Sheet, plan: Layout) -> dict[tuple[int, int], Cell]:
        if self._mask is None:
            return dict(sheet.cells)
        result = {}
        for (row, col), cell in sheet.cells.items():
            label = plan.headers.get(col) if plan.in_table(row, col) else "Wert"
            masked = self._mask(f"{label}: {cell.value}").partition(": ")[2] if cell.value else ""
            result[(row, col)] = Cell(masked, cell.formula)
        return result

    def _masked_part(self, part: Part) -> Part:
        if self._mask is None:
            return part
        text = self._mask(f"{part.column or 'Wert'}: {part.text}").partition(": ")[2]
        return Part(part.key, part.label, text, part.sheet, part.column)

    def row_lines(self, sheet: str, row: int) -> tuple[str, ...]:
        """Die Felder einer Zeile, Zeilenname zuerst, dann Werte und Formeln."""
        state = (self._sheets or {}).get(sheet)
        if state is None:
            return ()
        plan, lines, formulas = state.plan, [], []
        columns = sorted(c for (r, c) in state.cells if r == row)
        if plan.name_column in columns:
            columns.remove(plan.name_column)
            columns.insert(0, plan.name_column)
        for col in columns:
            cell = state.cells[(row, col)]
            label = plan.headers.get(col) if plan.in_table(row, col) else f"{column_letters(col)}{row}"
            if cell.value:
                lines.append(f"{label}: {cell.value}")
            if cell.formula is not None:
                formulas.append(f"{FORMULA} {label}: {readable(cell.formula, (row, col), plan)}")
        lines += formulas
        if self.sheets > 1:
            lines.append(f"Blatt: {sheet}")
        return tuple(lines)

    def active_state(self) -> tuple[str, dict[tuple[int, int], Cell], Layout] | None:
        """Das Blatt, an dem gerade gearbeitet wird: Name, Zellen (geschwärzt) und Aufbau."""
        sheets = self._sheets or {}
        sheet = self.active if self.active in sheets else next(iter(sheets), None)
        if sheet is None:
            return None
        return sheet, sheets[sheet].cells, sheets[sheet].plan

    def focus(self, sheet: str, row: int | None) -> None:
        """Hier steht der Cursor: Diese Zeile ist gemeint, wenn sie etwas erklärt."""
        self.active = sheet
        state = (self._sheets or {}).get(sheet)
        if row is None or state is None or state.plan.header_row is None:
            return
        if state.plan.header_row < row <= state.plan.last_row:
            self.last_row = self.row_lines(sheet, row) or self.last_row

    def update(
        self, name: str, sheets: list[Sheet], t: float, *, cells: bool = True, parts: bool = True,
        partial: bool = False, names: dict[str, str] | None = None, diff: bool = True,
    ) -> list[Event]:
        """Ein neuer Stand. cells/parts: was davon verglichen wird; partial: nur diese Blätter sind neu;
        diff=False: nur den Stand merken, ohne ihn mit dem vorigen zu vergleichen (der Prüfer braucht keine Ereignisse)."""
        events: list[Event] = []
        if cells:
            events += self._update_cells(name, sheets, t, partial, diff)
        if parts and sheets and all(s.parts for s in sheets):
            plans = {s.name: layout(s) for s in sheets}
            events += self._update_parts(name, parts_of(sheets, plans, names), t)
        if cells and not partial:
            events += self.sheet_list(name, [s.name for s in sheets], t)
        return events

    def _update_cells(self, name: str, sheets: list[Sheet], t: float, partial: bool, diff: bool = True) -> list[Event]:
        states = {}
        for sheet in sheets:
            plan = layout(sheet)
            states[sheet.name] = _State(self._masked(sheet, plan), plan)
        previous = self._sheets
        if partial and previous is not None and name == self.name:
            self._sheets = {**previous, **states}
        else:
            self._sheets = states
        if previous is None or name != self.name or not diff:
            # Erster Stand, eine andere Mappe oder nur merken: Ausgangsstand, nichts melden.
            if name != self.name:
                self._fresh = {}
            self.name = name
            return []
        formulas: list[Event] = []
        values: list[Event] = []
        sorts: list[Event] = []
        for sheet, state in states.items():
            old = previous.get(sheet)
            if old is None:
                continue
            sort = self._sort_event(sheet, old, state, t)
            if sort is not None:
                # Sortiert: Alle Zeilen stehen woanders, aber geändert ist nichts.
                sorts.append(sort)
                continue
            formulas += self._formula_events(sheet, old, state, t)
            values += self._value_events(sheet, old, state, t)
            typed = [pos for pos, cell in state.cells.items() if cell.value and not (old.cells.get(pos) or _BLANK).value]
            if len(typed) < RECORD_SWITCH:
                # Von Hand getippt. Viele auf einmal sind eingefügt oder geladen: Das sind die Daten, mit denen sie arbeitet.
                self._fresh.setdefault(sheet, set()).update(typed)
        if len(values) >= RECORD_SWITCH or len(formulas) > MAX_FORMULA_EVENTS:
            # Eingefügt oder ein ganzes Blatt neu: keine einzelne Entscheidung.
            first = (formulas + values)[0]
            return sorts + [Event(t=t, kind="record_changed", text=f"Viele Zellen auf einmal geändert ({len(values) + len(formulas)})",
                                  label=first.label, after=first.after, context=first.context)]
        events = formulas + values
        if events:
            self.last_row = events[-1].context or ()
        return sorts + events

    def _update_parts(self, name: str, found: dict[str, Part], t: float) -> list[Event]:
        found = {key: self._masked_part(part) for key, part in found.items()}
        previous, self._parts = self._parts, found
        if previous is None or name != self._parts_name:
            self._parts_name = name
            return []
        events = []
        for key in list(previous) + [k for k in found if k not in previous]:
            before, after = previous.get(key), found.get(key)
            if before is not None and after is not None and before.text == after.text:
                continue
            part = after or before
            old, new = (before.text if before else NONE), (after.text if after else NONE)
            verb = "neu" if before is None else "entfernt" if after is None else "geändert"
            context = (f"Blatt: {part.sheet}",) if part.sheet else None
            events.append(Event(
                t=t, kind="value_changed", text=f"{part.label} {verb}: {old} → {new}" + (f" (Blatt {part.sheet})" if part.sheet else ""),
                label=part.label, before=old, after=new, context=context,
            ))
        if len(events) > MAX_PART_EVENTS:
            return [Event(t=t, kind="record_changed", text=f"Viele Änderungen am Aufbau der Mappe ({len(events)})",
                          label=events[0].label, after=events[0].after)]
        return events

    def sheet_list(self, name: str, sheet_names: list[str], t: float) -> list[Event]:
        """Blätter neu, gelöscht oder umbenannt."""
        current = tuple(sheet_names)
        previous, self._sheet_names = self._sheet_names, current
        if previous is None or name != self._sheet_names_of:
            self._sheet_names_of = name
            return []
        added = [s for s in current if s not in previous]
        removed = [s for s in previous if s not in current]
        if len(added) == 1 and len(removed) == 1:
            return [Event(t=t, kind="value_changed", text=f"Blatt {removed[0]} umbenannt in {added[0]}",
                          label=f"Blatt {added[0]}", before=removed[0], after=added[0])]
        events = [Event(t=t, kind="value_changed", text=f"Blatt {s} neu", label=f"Blatt {s}", before=NONE, after="neu")
                  for s in added]
        events += [Event(t=t, kind="value_changed", text=f"Blatt {s} gelöscht", label=f"Blatt {s}", before="da", after=NONE)
                   for s in removed]
        return events[:MAX_PART_EVENTS]

    def _sort_event(self, sheet: str, old: _State, new: _State, t: float) -> Event | None:
        plan = new.plan
        if plan.header_row is None or plan.header_row != old.plan.header_row or plan.headers != old.plan.headers:
            return None
        columns = sorted(plan.headers)
        rows = range(plan.header_row + 1, max(plan.last_row, old.plan.last_row) + 1)

        def signature(state: _State, row: int) -> tuple:
            cells = [state.cells.get((row, col), _BLANK) for col in columns]
            return tuple((c.value, pattern(c.formula, (row, col)) if c.formula is not None else "") for c, col in zip(cells, columns))

        before = [signature(old, r) for r in rows]
        after = [signature(new, r) for r in rows]
        filled = [s for s in after if any(v for v, _ in s)]
        if before == after or len(filled) < 2 or sorted(before) != sorted(after):
            return None
        order = []
        for index, col in enumerate(columns):
            now = _order([s[index][0] for s in filled])
            was = _order([s[index][0] for s in before if any(v for v, _ in s)])
            if now is not None:
                order.append((was == now, index, col, now))
        if order:
            _, _, col, direction = min(order)
            text = f"nach {plan.headers[col]} {direction}"
        else:
            text = "von Hand umgestellt"
        return Event(t=t, kind="value_changed", text=f"{SORTING}: {text} (Blatt {sheet})", label=SORTING,
                     before=NONE, after=text, context=(f"Blatt: {sheet}",))

    def _formula_events(self, sheet: str, old: _State, new: _State, t: float) -> list[Event]:
        groups: dict[tuple, list[int]] = {}
        for pos in sorted(set(old.cells) | set(new.cells)):
            before, after = old.cells.get(pos), new.cells.get(pos)
            was = pattern(before.formula, pos) if before and before.formula is not None else None
            now = pattern(after.formula, pos) if after and after.formula is not None else None
            if was == now:
                continue
            groups.setdefault((pos[1], was, now), []).append(pos[0])
        events = []
        cells = new.cells
        for (col, was, now), rows in groups.items():
            first = rows[0]
            plan = new.plan
            if plan.in_table(first, col):
                label = f"{FORMULA} {plan.column(col)}"
            else:
                word = _left_label(cells, first, col)
                label = f"{FORMULA} {word} ({column_letters(col)}{first})" if word else f"{FORMULA} {column_letters(col)}{first}"
            old_cell, new_cell = old.cells.get((first, col)), new.cells.get((first, col))
            before = readable(old_cell.formula, (first, col), old.plan) if was is not None else NONE
            after = readable(new_cell.formula, (first, col), plan) if now is not None else NONE
            events.append(Event(
                t=t, kind="value_changed", text=f"{label} geändert: {before} → {after} ({_where(plan, rows)})",
                label=label, before=before, after=after, context=self.row_lines(sheet, first) or None,
            ))
        return events

    def _value_events(self, sheet: str, old: _State, new: _State, t: float) -> list[Event]:
        events = []
        fresh = self._fresh.get(sheet, set())
        for pos, before in old.cells.items():
            if before.formula is not None or not before.value or new.plan.is_derived(*pos) or pos in fresh:
                continue
            after = new.cells.get(pos)
            if after is not None and after.formula is not None:
                continue  # Jetzt eine Formel: steht schon bei den Formeln.
            value = after.value if after is not None and after.value else EMPTY
            if value == before.value:
                continue
            row, col = pos
            plan = new.plan
            if plan.in_table(row, col):
                if col == plan.name_column:
                    continue  # Ein neuer Zeilenname ist eine andere Zeile, kein geänderter Wert.
                label, where = plan.headers[col], _where(plan, [row])
            else:
                word = _left_label(new.cells, row, col)
                label, where = (word or f"{column_letters(col)}{row}"), f"Blatt {sheet}"
            events.append(Event(
                t=t, kind="value_changed", text=f"{label} geändert: {before.value} → {value} ({where})",
                label=label, before=before.value, after=value, context=self.row_lines(sheet, row) or None,
            ))
        return events

    # --- Übersicht für die KI ---

    def overview(self) -> str:
        """Die Mappe in wenigen Zeilen: Blätter, Tabellen, Formeln je Spalte, Filter, Diagramme …"""
        if not self._sheets and not self._parts:
            return ""
        names = list(self._sheet_names or self._sheets or ())
        head = f"Excel-Mappe „{self.name or self._parts_name}“"
        if len(names) > 1:
            head += f" mit den Blättern {', '.join(names[:10])}"
        lines = [head + (f", gerade auf Blatt {self.active}" if self.active and len(names) > 1 else "")]
        order = sorted((self._sheets or {}).items(), key=lambda item: item[0] != self.active)
        for sheet, state in order[:4]:
            lines += self._sheet_overview(sheet, state)
        for part in list((self._parts or {}).values())[:15]:
            where = f" (Blatt {part.sheet})" if part.sheet and len(names) > 1 else ""
            lines.append(f"{part.label}: {part.text}{where}")
        text = "\n".join(lines)
        return text if len(text) <= MAX_OVERVIEW else text[: MAX_OVERVIEW - 1].rstrip() + "…"

    def _sheet_overview(self, sheet: str, state: _State) -> list[str]:
        plan, cells = state.plan, state.cells
        lines = []
        if plan.header_row is not None:
            rows = [r for r in plan.names if plan.header_row < r <= plan.last_row]
            table = f"Excel-Tabelle {plan.table}" if plan.table else "Tabelle"
            text = f"Blatt {sheet}: {table} mit Kopfzeile in Zeile {plan.header_row}, Spalten {', '.join(plan.headers.values())}"
            if rows:
                shown = [plan.names[r] for r in rows[:8]]
                more = f" und {len(rows) - 8} weitere" if len(rows) > 8 else ""
                text += f"; {len(rows)} Zeilen ({', '.join(shown)}{more})"
            lines.append(text)
            totals = []
            for col in sorted(plan.headers):
                found = []
                for (r, cc), c in sorted(cells.items()):
                    if cc != col or c.formula is None or not plan.in_table(r, col):
                        continue
                    if r in plan.names:
                        found.append((r, pattern(c.formula, (r, col))))
                    else:
                        # Eine Zeile ohne Namen, meist die Summe darunter.
                        totals.append(f"{column_letters(col)}{r} {plan.headers[col]} {readable(c.formula, (r, col), plan)}")
                if not found:
                    continue
                common, count = Counter(p for _, p in found).most_common(1)[0]
                first = min(r for r, p in found if p == common)
                formula = readable(cells[(first, col)].formula, (first, col), plan)
                if count == len(found):
                    same = "in allen Zeilen gleich" if count > 1 else f"nur bei {plan.names[first]}"
                elif count == 1:
                    same = f"bei {plan.names[first]}; jede Zeile hat eine eigene Formel"
                else:
                    same = f"in {count} von {len(found)} Zeilen gleich"
                lines.append(f"Formel {plan.headers[col]}: {formula} ({same})")
            if totals:
                lines.append(f"Zeilen ohne Namen (Ergebnisse): {'; '.join(totals[:6])}")
        elif any(not plan.is_derived(*pos) for pos in cells):
            lines.append(f"Blatt {sheet}: keine Kopfzeile erkannt")
        outside = []
        for (row, col), cell in sorted(cells.items()):
            if plan.in_table(row, col) or plan.is_derived(row, col) or (row == plan.header_row and col in plan.headers):
                continue
            value = readable(cell.formula, (row, col), plan) if cell.formula is not None else cell.value
            if value:
                outside.append(f"{column_letters(col)}{row} {value}")
            if len(outside) >= 10:
                break
        if outside:
            lines.append(f"Außerhalb der Tabelle: {'; '.join(outside)}")
        return lines
