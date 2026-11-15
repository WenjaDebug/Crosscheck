"""Eine .xlsx-Datei lesen: Zellen, Formeln und was Excel sonst darin speichert.

Gelesen wird mit Python allein (zipfile und XML), ohne zusätzliches Paket. Neben den Zellen:
Excel-Tabellen, Filter und Sortierung, bedingte Formatierung, Auswahllisten und andere
Eingaberegeln, Diagramme, Pivot-Tabellen und Namen. Hier steht alles noch mit Zelladressen;
was die Spaltenköpfe sind und wie es sich vorlesen lässt, weiß workbook.py.

Jedes dieser Teile ist für sich gelesen: Versteht Gegenprobe eines nicht (eine neue Excel-Version,
eine andere Tabellenkalkulation), fehlt nur dieses Teil, die Zellen kommen trotzdem.
"""

from __future__ import annotations

import io
import logging
import posixpath
import re
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from decimal import Decimal
from xml.etree import ElementTree

from .sprache import T

logger = logging.getLogger(__name__)

MAX_FILE = 100_000_000
# Ausgepackt höchstens so groß je Teil der Datei (gegen gepackte Riesendateien). Ein Blatt wird beim Lesen
# verarbeitet und darf größer sein: Was über MAX_CELLS hinausgeht, wird nur überlesen.
MAX_PART = 80_000_000
MAX_SHEET_PART = 1_000_000_000
MAX_SHEETS = 20
# So viele Zellen liest der Prüfer je Blatt und in der ganzen Mappe. Ein größeres Blatt liest er bis zur letzten
# ganzen Zeile darunter (Cells.cut), ein Blatt nach der Grenze gar nicht mehr.
MAX_CELLS = 500_000
MAX_BOOK_CELLS = 1_500_000
MAX_VALUE = 200
MAX_FORMULA = 300
# Mehr Diagramme, Pivot-Tabellen, Formatierungen … je Blatt liest niemand vor.
MAX_PARTS = 40

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_X14 = "{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}"
_XM = "{http://schemas.microsoft.com/office/excel/2006/main}"
_C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
_CX = "{http://schemas.microsoft.com/office/drawing/2014/chartex}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_XDR = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"

ADDRESS = re.compile(r"^([A-Z]{1,3})(\d{1,7})$")
# Eine Zelladresse in einer Formel, auch auf anderen Blättern: B3, $B$3, Tabelle2!B3.
ANY_REF = re.compile(r"(?<![A-Za-z0-9_.$\\])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(])")
STRING = re.compile(r'("(?:[^"]|"")*")')
# Ein Bezug auf eine Excel-Tabelle: Tabelle1[Betrag], Tabelle1[@Betrag], [@[Anzahl Sitze]],
# Tabelle1[[#This Row],[Betrag]]. Darin wird nichts verschoben und nichts ersetzt.
STRUCTURED = re.compile(r"(?:(?<![A-Za-z0-9_.\\])([A-Za-z_\\][A-Za-z0-9_.\\]*))?\[((?:[^\[\]]|\[[^\[\]]*\])*)\]")
_ESCAPED = re.compile(r"_x([0-9A-Fa-f]{4})_")
_ISO_DATE = re.compile(r"^(\d{4})-(\d\d)-(\d\d)T00:00:00$")
_LITERALS = re.compile(rf"{STRING.pattern}|({STRUCTURED.pattern})")

# Eingebaute Zahlenformate für Datum und Uhrzeit.
_DATE_FORMATS = set(range(14, 23)) | set(range(45, 48)) | set(range(27, 37)) | set(range(50, 59))
_PERCENT_FORMATS = {9, 10}
# Ganze Zahlen ohne Nachkommastellen: 0, #,##0 und die Buchhaltungsformate ohne Dezimalen.
_WHOLE_FORMATS = {1, 3, 37, 38}
_WHOLE = re.compile(r"[0#]")

# Bereich: oberste Zeile, linke Spalte, unterste Zeile, rechte Spalte.
Area = tuple[int, int, int, int]


@dataclass(frozen=True)
class Cell:
    value: str
    # Die Formel wie in der Datei, z. B. „B3*577/100“ (ohne „=“), sonst None.
    formula: str | None = None
    # Der genaue Text einer Textzelle, wenn er anders rechnet, als er aussieht: Ziffern als Text („123“)
    # oder Leerzeichen am Rand („A “). SVERWEIS findet solche Werte nicht, darum merkt sich der Prüfer sie.
    text: str | None = None
    # Die Formel, wie Excel sie zeigt, wenn formula sie mit Zellen statt Tabellenbezügen und Namen schreibt
    # (=[@Menge]*MwSt als C5*$I$1, bezuege.py), sonst None.
    typed: str | None = None


_NUMERIC = re.compile(r"[-+−]?[\d.,'’ ]*\d[\d.,'’ ]*%?")


def kept_text(raw: str) -> str | None:
    """Den Text behalten, wenn er nach Zahl aussieht oder andere Leerzeichen hat, als man sieht (am Rand, doppelt),
    sonst None."""
    plain = raw.strip()
    if plain and (raw != " ".join(plain.split()) or _NUMERIC.fullmatch(plain)) and "\n" not in raw and len(raw) <= MAX_VALUE:
        return raw
    return None


@dataclass(frozen=True)
class ExcelTable:
    """Eine Excel-Tabelle (Einfügen → Tabelle): Kopfzeile, Spaltennamen, Ergebniszeile."""

    name: str
    area: Area
    columns: tuple[str, ...]
    header: bool = True
    totals: int = 0
    # Formeln, die Excel für ganze Spalten einträgt, je Spaltenname.
    calculated: tuple[tuple[str, str], ...] = ()


@dataclass(frozen=True)
class FilterColumn:
    column: int
    # Was der Filter zeigt, schon als Text: „nur ja, nein“, „> 5000“, „die 10 größten“.
    condition: str


@dataclass(frozen=True)
class SortKey:
    column: int
    descending: bool = False
    # Nach Farbe oder Symbol statt nach Wert.
    by: str = ""


@dataclass(frozen=True)
class AutoFilter:
    area: Area
    filters: tuple[FilterColumn, ...] = ()
    sort: tuple[SortKey, ...] = ()
    # Gehört er zu einer Excel-Tabelle, ihr Name.
    table: str = ""


@dataclass(frozen=True)
class Rule:
    """Bedingte Formatierung oder Eingaberegel (Datenüberprüfung) für einen Bereich."""

    areas: tuple[Area, ...]
    kind: str
    operator: str = ""
    formulas: tuple[str, ...] = ()
    # Was sonst dazugehört: Symbolsatz, gesuchter Text, Rang, Zeitraum.
    extra: str = ""


@dataclass(frozen=True)
class Series:
    name: str = ""
    # Bezüge wie in der Datei, z. B. „Tabelle1!$D$6:$D$10“.
    name_ref: str = ""
    values: str = ""
    categories: str = ""


@dataclass(frozen=True)
class Chart:
    name: str
    kind: str
    title: str = ""
    series: tuple[Series, ...] = ()


@dataclass(frozen=True)
class PivotValue:
    """Ein Wertfeld einer Pivot-Tabelle: „Summe von Umsatz“ fasst die Spalte „Umsatz“ der Daten zusammen."""

    caption: str
    field: str
    # sum, count, average, max, min, product, countNums, stdDev, stdDevp, var, varp; „“, wenn unbekannt.
    function: str = "sum"
    # Als Anteil, Differenz, laufende Summe …: Das rechnet der Prüfer nicht nach.
    shown_as: bool = False


@dataclass(frozen=True)
class Pivot:
    name: str
    area: Area | None
    rows: tuple[str, ...] = ()
    columns: tuple[str, ...] = ()
    values: tuple[str, ...] = ()
    filters: tuple[str, ...] = ()
    # Ausgeblendete Einträge je Feld: „Partei ohne Rest“.
    hidden: tuple[str, ...] = ()
    source: str = ""
    # Zum Nachrechnen (pivotdiagramm.py): die Wertfelder, die Zahlen der Pivot-Tabelle ohne Beschriftungen, und die
    # Quelle als Blatt und Bereich oder als Name einer Excel-Tabelle.
    value_fields: tuple[PivotValue, ...] = ()
    body: Area | None = None
    source_sheet: str = ""
    source_area: Area | None = None
    source_table: str = ""
    # Ausgeblendete Einträge je Feld, auch die nicht gewählten eines Berichtsfilters.
    skipped: tuple[tuple[str, tuple[str, ...]], ...] = ()
    # Gesamtergebnis für Zeilen und für Spalten.
    grand: tuple[bool, bool] = (True, True)
    # Blendet ein Beschriftungs- oder Wertfilter Einträge aus? None: unbekannt.
    filtered: bool | None = False
    # False bei berechneten oder gruppierten Feldern und Werten in den Zeilen: Dann rechnet der Prüfer nicht nach.
    plain: bool = True


class Cells(dict):
    """Die Zellen eines Blatts, wie sie gelesen sind. cut: Das Blatt ist größer, gelesen sind die Zeilen bis cut,
    darunter kann noch etwas stehen. Formeln, die dort rechnen, rechnet der Prüfer nicht nach (analyse.Sheet.partial).
    Nach dem Lesen ändert sich nichts mehr daran, darum darf sich der Prüfer Schlüssel dazu merken (key)."""

    def __init__(self, cells=(), cut: int | None = None) -> None:
        super().__init__(cells)
        self.cut = cut
        # Was der Prüfer über diese Zellen einmal ausrechnet (Schlüssel, Verweise auf andere Blätter).
        self.memo: dict = {}
        # Ausgeblendete Zeilen (aus welchem Grund auch immer), die davon, die ein Filter ausblendet, und die einer
        # eingeklappten Gruppe (Daten → Gruppieren). None: unbekannt. Einsetzen ändert daran nichts.
        self.hidden_rows: frozenset[int] | None = getattr(cells, "hidden_rows", None)
        # Zahlen mit einem Datums- oder Uhrzeitformat, das sie verdeckt (8 als hh:mm zeigt Excel als 00:00): {Zelle: Zahl}.
        self.clipped: dict = dict(getattr(cells, "clipped", {}))
        # Zahlen, die ein Format ohne Nachkommastellen als 0 zeigt (0,4 als #,##0): {Zelle: Zahl}.
        self.rounded: dict = dict(getattr(cells, "rounded", {}))
        self.filtered_rows: frozenset[int] | None = getattr(cells, "filtered_rows", None)
        self.grouped_rows: frozenset[int] | None = getattr(cells, "grouped_rows", None)


def hide_rows(sheet: "Sheet", hidden, grouped=()) -> None:
    """Die ausgeblendeten Zeilen zu den Zellen des Blatts; weggefiltert sind die in einem Bereich mit aktivem Filter."""
    if not isinstance(sheet.cells, Cells):
        sheet.cells = Cells(sheet.cells)
    hidden = frozenset(hidden)
    areas = [f.area for f in sheet.filters if f.filters]
    sheet.cells.hidden_rows = hidden
    sheet.cells.filtered_rows = frozenset(r for r in hidden if any(a[0] < r <= a[2] for a in areas))
    sheet.cells.grouped_rows = frozenset(grouped) & hidden


def cut_of(cells) -> int | None:
    """Bis zu welcher Zeile ein Blatt gelesen ist, wenn es größer ist, sonst None."""
    return getattr(cells, "cut", None)


@dataclass
class Sheet:
    name: str
    cells: dict[tuple[int, int], Cell] = field(default_factory=dict)
    # False: nur die Zellen sind bekannt (z. B. direkt aus dem laufenden Excel gelesen).
    parts: bool = False
    tables: list[ExcelTable] = field(default_factory=list)
    filters: list[AutoFilter] = field(default_factory=list)
    formats: list[Rule] = field(default_factory=list)
    validations: list[Rule] = field(default_factory=list)
    charts: list[Chart] = field(default_factory=list)
    pivots: list[Pivot] = field(default_factory=list)
    # Woran das Seitenfenster in Excel diesen Stand des Blatts erkennt (addin.sheets_from).
    stamp: str = ""
    # Wie das Blatt Tabellenbezüge und Namen schreibt (bezuege.Style), damit Korrekturen genauso aussehen.
    refs: object = None
    # Verbundene Zellen (Start → Zusammenführen): Der Wert steht nur in der ersten Zelle jedes Bereichs.
    merged: list[Area] = field(default_factory=list)
    # Andere Dateien, aus denen Formeln Werte holen: die Nummer in [1]Plan!B3 → Dateiname oder Pfad.
    links: dict[str, str] = field(default_factory=dict)
    # Ausgeblendet (Rechtsklick auf den Blattreiter → Ausblenden): Formeln anderer Blätter rechnen trotzdem damit.
    hidden: bool = False


@dataclass
class Book:
    sheets: list[Sheet] = field(default_factory=list)
    # Namen für Bereiche oder Werte: „Sitze_gesamt“ → „Tabelle1!$D$2“.
    names: dict[str, str] = field(default_factory=dict)
    # Namen, die nur auf einem Blatt gelten: Blatt → {Name → Bezug}.
    local: dict[str, dict[str, str]] = field(default_factory=dict)


# --- Adressen ---


def column_number(letters: str) -> int:
    number = 0
    for ch in letters:
        number = number * 26 + ord(ch) - 64
    return number


def column_letters(number: int) -> str:
    letters = ""
    while number:
        number, rest = divmod(number - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def cell_position(address: str) -> tuple[int, int] | None:
    match = ADDRESS.match(address.replace("$", "").strip())
    if not match:
        return None
    return int(match.group(2)), column_number(match.group(1))


def area_of(ref: str) -> Area | None:
    """„B4:D10“ → (4, 2, 10, 4), auch „Tabelle1!$B$4:$D$10“ und eine einzelne Zelle."""
    ref = ref.rsplit("!", 1)[-1].strip()
    first, _, last = ref.partition(":")
    a, b = cell_position(first), cell_position(last or first)
    if a is None or b is None:
        return None
    return min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1])


def area_text(area: Area) -> str:
    top, left, bottom, right = area
    first = f"{column_letters(left)}{top}"
    return first if (top, left) == (bottom, right) else f"{first}:{column_letters(right)}{bottom}"


def sheet_of(ref: str) -> str:
    """Das Blatt eines Bezugs wie „'Mein Blatt'!$A$1“, sonst „“."""
    if "!" not in ref:
        return ""
    return ref.rsplit("!", 1)[0].strip().strip("'").replace("''", "'")


def in_area(area: Area, row: int, col: int) -> bool:
    return area[0] <= row <= area[2] and area[1] <= col <= area[3]


# --- Teile der Datei ---


def _info(archive: zipfile.ZipFile, name: str) -> zipfile.ZipInfo | None:
    """Ein Teil der Datei. Manche Programme schreiben die Namen mit \\ und in anderer Schreibweise
    (xl\\sharedstrings.xml): Excel liest sie trotzdem, also auch wir."""
    try:
        return archive.getinfo(name)
    except KeyError:
        wanted = name.casefold()
        return next((i for i in archive.infolist() if i.filename.replace("\\", "/").casefold() == wanted), None)


def _part(archive: zipfile.ZipFile, name: str) -> ElementTree.Element | None:
    info = _info(archive, name)
    if info is None:
        return None
    if info.file_size > MAX_PART:
        raise ValueError(T("Die Excel-Datei ist zu groß.", "The Excel file is too large."))
    raw = archive.read(info)
    if b"<!DOCTYPE" in raw[:2000]:
        raise ValueError(T("Das ist keine Excel-Datei.", "This is not an Excel file."))
    return ElementTree.fromstring(raw)


def _rels(archive: zipfile.ZipFile, path: str) -> dict[str, tuple[str, str]]:
    """Die Verweise eines Teils: Id → (Art, z. B. „table“, Pfad in der Datei)."""
    folder, name = posixpath.split(path)
    root = _part(archive, posixpath.join(folder, "_rels", f"{name}.rels"))
    links = {}
    for rel in root if root is not None else ():
        if rel.tag != f"{_PKG}Relationship" or rel.get("TargetMode") == "External":
            continue
        target = rel.get("Target") or ""
        target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
        links[rel.get("Id") or ""] = ((rel.get("Type") or "").rsplit("/", 1)[-1], target)
    return links


def _text(node: ElementTree.Element | None) -> str:
    """Text einer Zelle oder eines geteilten Textes, auch mit Formatierung in Stücken."""
    if node is None:
        return ""
    return "".join(t.text or "" for t in node.iter(f"{_NS}t"))


def _clean(value: str, limit: int = MAX_VALUE) -> str:
    # Excel schreibt Zeichen wie den Zeilenumbruch in einer Zelle als „_x000a_“.
    value = " ".join(_ESCAPED.sub(lambda m: chr(int(m.group(1), 16)), value).split())
    return value if len(value) <= limit else value[: limit - 1].rstrip() + "…"


def whole_format(code: str) -> bool:
    """Ein Zahlenformat, das ganze Zahlen zeigt (#,##0 oder 0 €), ohne Text, Datum und Exponent; code klein, ohne
    Texte in Anführungszeichen."""
    first = code.split(";")[0]
    return bool(_WHOLE.search(first)) and "." not in first and "e" not in first and "general" not in first \
        and "standard" not in first and "@" not in first


def _styles(archive: zipfile.ZipFile) -> list[str]:
    """Je Zellformat: „date“, „percent“, „whole“ (ganze Zahlen) oder „“."""
    root = _part(archive, "xl/styles.xml")
    if root is None:
        return []
    custom = {}
    formats = root.find(f"{_NS}numFmts")
    for fmt in formats if formats is not None else ():
        code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", fmt.get("formatCode") or "").lower()
        if "%" in code:
            kind = "percent"
        elif re.search(r"[dy]|h+:m|m+:s", code):
            kind = "date"
        elif whole_format(code):
            kind = "whole"
        else:
            kind = ""
        custom[int(fmt.get("numFmtId") or 0)] = kind
    kinds = []
    xfs = root.find(f"{_NS}cellXfs")
    for xf in xfs if xfs is not None else ():
        number = int(xf.get("numFmtId") or 0)
        if number in custom:
            kinds.append(custom[number])
        elif number in _DATE_FORMATS:
            kinds.append("date")
        elif number in _PERCENT_FORMATS:
            kinds.append("percent")
        elif number in _WHOLE_FORMATS:
            kinds.append("whole")
        else:
            kinds.append("")
    return kinds


def number_text(value: float) -> str:
    """Wie in der deutschen Excel: 191,564 statt 191.564, ganze Zahlen ohne Komma. Nie als 1,3e-05: So läse der
    Prüfer einen kleinen p-Wert aus Excel als Text statt als Zahl."""
    if value == int(value) and abs(value) < 1e15:
        return str(int(value))
    # Zehn Stellen verdecken Rundungsreste wie 0,30000000000000004, große Beträge brauchen alle 15 von Excel.
    text = f"{value:.15g}" if abs(value) >= 1e10 else f"{value:.10g}"
    if "e" in text:
        text = format(Decimal(text), "f")
    return text.replace(".", ",")


def date_text(moment: datetime) -> str:
    if moment.second:
        return moment.strftime("%d.%m.%Y %H:%M:%S")
    if moment.hour or moment.minute:
        return moment.strftime("%d.%m.%Y %H:%M")
    return moment.strftime("%d.%m.%Y")


def _date_text(serial: float, date1904: bool) -> str:
    """Auf die Sekunde gerundet, nicht abgeschnitten: 0,395833 (9:29:59,97) ist 09:30, und eine Zeit mit Sekunden
    (08:15:42) behält sie, sonst rechnete der Prüfer mit einer Minute weniger als Excel."""
    start = datetime(1904, 1, 1) if date1904 else datetime(1899, 12, 30)
    try:
        moment = start + timedelta(seconds=round(serial * 86400))
    except (OverflowError, ValueError):
        return number_text(serial)
    if serial < 1:
        return moment.strftime("%H:%M:%S" if moment.second else "%H:%M")
    return date_text(moment)


def _as_text(raw: str) -> str | None:
    """Wie kept_text, und dazu Text, der wie ein Fehlerwert aussieht („#div/0!“ als Beschreibung): In der Datei ist
    er als Text gespeichert, also ist er kein Fehler. Aus dem Seitenfenster kommen echte Fehler auch als Text, darum
    nur hier."""
    kept = kept_text(raw)
    if kept is None and raw.strip().startswith("#") and "\n" not in raw and len(raw) <= MAX_VALUE:
        return raw
    return kept


def _value(cell: ElementTree.Element, shared: list[str], styles: list[str], date1904: bool) -> tuple[str, str | None]:
    """Der Wert, wie der Prüfer ihn zeigt, und bei Text der genaue Text (siehe Cell.text)."""
    kind = cell.get("t") or "n"
    if kind == "inlineStr":
        raw = _text(cell.find(f"{_NS}is"))
        return _clean(raw), _as_text(raw)
    raw = cell.findtext(f"{_NS}v")
    if raw is None:
        return "", None
    if kind == "s":
        try:
            raw = shared[int(raw)]
        except (ValueError, IndexError):
            return "", None
        return _clean(raw), _as_text(raw)
    if kind == "b":
        return ("WAHR" if raw.strip() == "1" else "FALSCH"), None
    if kind == "str":
        return _clean(raw), _as_text(raw)
    return _number_value(raw, cell, styles, date1904), None


def _number_value(raw: str, cell: ElementTree.Element, styles: list[str], date1904: bool) -> str:
    if (cell.get("t") or "n") == "e":
        return _clean(raw)
    try:
        number = float(raw)
    except ValueError:
        return _clean(raw)
    style = int(cell.get("s") or 0)
    look = styles[style] if style < len(styles) else ""
    if look == "date":
        return _date_text(number, date1904)
    if look == "percent":
        return f"{number_text(round(number * 100, 10))} %"
    return number_text(number)


def outside_literals(formula: str, change) -> str:
    """Nur außerhalb von "Texten" und Bezügen auf Excel-Tabellen ändern."""
    parts, last = [], 0
    for match in _LITERALS.finditer(formula):
        parts += [change(formula[last : match.start()]), match.group(0)]
        last = match.end()
    return "".join(parts + [change(formula[last:])])


def shift(formula: str, rows: int, columns: int) -> str:
    """Eine geteilte Formel an eine andere Zelle anpassen, wie Excel beim Herunterziehen."""

    def move(match: re.Match) -> str:
        dollar_col, letters, dollar_row, digits = match.groups()
        col = column_number(letters) if dollar_col else column_number(letters) + columns
        row = int(digits) if dollar_row else int(digits) + rows
        if col < 1 or row < 1:
            return "#BEZUG!"
        return f"{dollar_col}{column_letters(col)}{dollar_row}{row}"

    return outside_literals(formula, lambda part: ANY_REF.sub(move, part))


# --- Lesen ---


def read_xlsx(data: bytes) -> list[Sheet]:
    """Alle Blätter einer .xlsx-Datei mit Werten und Formeln. Wirft ValueError, wenn es keine ist."""
    return read_book(data).sheets


class Both(str):
    """Ein Text, den schon das Lesen der Datei formuliert (eine Filterbedingung, eine Diagrammart): als str deutsch,
    so wird er gemerkt und verglichen, und auf Englisch in en. said() gibt ihn in der Sprache der Anfrage."""

    en: str

    def __new__(cls, german: str, english: str | None = None) -> "Both":
        text = super().__new__(cls, german)
        text.en = german if english is None else english
        return text


def said(text: str) -> str:
    """Ein gelesener Text (Both) in der Sprache der Anfrage; ein gewöhnlicher bleibt, wie er ist."""
    return T(str(text), getattr(text, "en", str(text)))


class Locked(ValueError):
    """Eine Datei, die wir nicht lesen können, mit einem Satz, was man stattdessen tut."""


_OLE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def read_book(data: bytes) -> Book:
    """Die ganze Mappe: Blätter mit allem, was dazugehört, und die Namen."""
    if len(data) > MAX_FILE:
        raise ValueError(T("Die Excel-Datei ist zu groß.", "The Excel file is too large."))
    if not data:
        raise ValueError(T("Die Datei ist leer.", "The file is empty."))
    if data.startswith(_OLE):
        # Mit Kennwort verschlüsselte .xlsx und alte .xls stecken beide in diesem älteren Containerformat.
        if "EncryptionInfo".encode("utf-16-le") in data:
            raise Locked(T("Die Datei ist mit einem Kennwort geschützt, so kann ich sie nicht lesen. "
                           "Öffne sie in Excel und prüfe sie dort im Seitenfenster.",
                           "The file is protected with a password, so I cannot read it. "
                           "Open it in Excel and check it there in the side panel."))
        raise Locked(T("Das ist eine alte Excel-Datei (.xls). Speichere sie in Excel als .xlsx (Datei → Speichern unter) "
                       "oder öffne sie in Excel und prüfe sie dort im Seitenfenster.",
                       "This is an old Excel file (.xls). Save it in Excel as .xlsx (File → Save As) "
                       "or open it in Excel and check it there in the side panel."))
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            book = _read(archive)
    except (zipfile.BadZipFile, ElementTree.ParseError, KeyError, EOFError) as err:
        # Auch eine Datei, die Excel gerade noch schreibt.
        raise ValueError(T("Die Excel-Datei lässt sich gerade nicht lesen.", "The Excel file cannot be read right now.")) from err
    from . import bezuege  # noqa: PLC0415 - bezuege baut auf diesem Modul auf

    bezuege.resolve(book.sheets, book.names, book.local)
    return book


def _read(archive: zipfile.ZipFile) -> Book:
    book_xml = _part(archive, "xl/workbook.xml")
    if book_xml is None:
        names = set(archive.namelist())
        if "mimetype" in names and b"opendocument.spreadsheet" in archive.read("mimetype"):
            raise Locked(T("Das ist eine Tabelle aus LibreOffice oder OpenOffice (.ods). Speichere sie dort als "
                           "Excel-Datei (Datei → Speichern unter → Excel 2007-365), dann prüfe ich sie.",
                           "This is a LibreOffice or OpenOffice spreadsheet (.ods). Save it there as an Excel file "
                           "(File → Save As → Excel 2007-365), then I check it."))
        if any(n.startswith("Index/") or n == "Index.zip" for n in names):
            raise Locked(T("Das ist eine Numbers-Datei. Exportiere sie in Numbers als Excel-Datei (Ablage → Exportieren "
                           "→ Excel), dann prüfe ich sie.",
                           "This is a Numbers file. Export it from Numbers as an Excel file (File → Export To → Excel), "
                           "then I check it."))
        raise ValueError(T("Das ist keine Excel-Datei (.xlsx).", "This is not an Excel file (.xlsx)."))
    props = book_xml.find(f"{_NS}workbookPr")
    date1904 = props is not None and props.get("date1904") in ("1", "true")
    links = _rels(archive, "xl/workbook.xml")
    strings_root = _part(archive, "xl/sharedStrings.xml")
    shared = [_text(si) for si in strings_root.iter(f"{_NS}si")] if strings_root is not None else []
    styles = _styles(archive)
    book = Book()
    entries = list(book_xml.iter(f"{_NS}sheet"))
    # Ausgeblendete Blätter liest er mit (Formeln rechnen damit), die sichtbaren aber zuerst, falls die Zellen nicht
    # für alle reichen; danach wieder in der Reihenfolge der Mappe.
    ranked = sorted(range(len(entries)), key=lambda i: entries[i].get("state") in ("hidden", "veryHidden"))
    read: dict[int, tuple[Sheet, object, str]] = {}
    hidden: dict[int, _Cells] = {}
    for i in ranked:
        entry = entries[i]
        if len(read) >= MAX_SHEETS:
            break
        _, path = links.get(entry.get(f"{_REL}id") or "", ("", ""))
        if not path:
            continue
        sheet = Sheet(_clean(entry.get("name") or "Blatt", 40), parts=True,
                      hidden=entry.get("state") in ("hidden", "veryHidden"))
        left = MAX_BOOK_CELLS - sum(len(s.cells) for s, _, _ in read.values())
        reader = _Cells(sheet, shared, styles, date1904, min(MAX_CELLS, left))
        root = _sheet_part(archive, path, reader)
        if root is None:
            continue
        read[i] = (sheet, root, path)
        hidden[i] = reader
    sheets = [read[i] for i in sorted(read)]
    book.sheets = [sheet for sheet, _, _ in sheets]
    files = _external_files(archive, book_xml, links)
    for sheet, root, path in sheets:
        _read_parts(archive, root, path, sheet)
        sheet.links = files
    for i, reader in hidden.items():
        if reader.rounded:
            sheet = read[i][0]
            if not isinstance(sheet.cells, Cells):
                sheet.cells = Cells(sheet.cells)
            sheet.cells.rounded = dict(reader.rounded)
        if reader.hidden:
            hide_rows(read[i][0], reader.hidden, reader.grouped)
    order = [_clean(entry.get("name") or "Blatt", 40) for entry in book_xml.iter(f"{_NS}sheet")]
    names = book_xml.find(f"{_NS}definedNames")
    for item in names if names is not None else ():
        name = item.get("name") or ""
        if not name or name.startswith("_xlnm.") or not item.text:
            continue
        local = item.get("localSheetId")
        if local is not None:
            # Gilt nur auf einem Blatt (die Nummer zählt alle Blätter der Mappe, auch ausgeblendete).
            if local.isdigit() and int(local) < len(order):
                book.local.setdefault(order[int(local)], {})[_clean(name, 60)] = _clean(item.text, MAX_FORMULA)
        else:
            book.names[_clean(name, 60)] = _clean(item.text, MAX_FORMULA)
    return book


def _external_files(archive: zipfile.ZipFile, book_xml, links: dict[str, tuple[str, str]]) -> dict[str, str]:
    """Die Dateien hinter [1], [2] …: in der Reihenfolge von externalReferences, je mit ihrem Pfad."""
    found: dict[str, str] = {}
    refs = book_xml.find(f"{_NS}externalReferences")
    for i, ref in enumerate(refs if refs is not None else (), start=1):
        _, path = links.get(ref.get(f"{_REL}id") or "", ("", ""))
        if not path:
            continue
        folder, name = posixpath.split(path)
        root = _part(archive, posixpath.join(folder, "_rels", f"{name}.rels"))
        for rel in root if root is not None else ():
            if rel.tag == f"{_PKG}Relationship" and rel.get("TargetMode") == "External" and rel.get("Target"):
                found[str(i)] = _clean(rel.get("Target") or "", 300)
                break
    return found


class _Cells:
    """Die Zellen eines Blatts, Zeile für Zeile beim Lesen der Datei, höchstens limit viele."""

    def __init__(self, sheet: Sheet, shared, styles, date1904: bool, limit: int) -> None:
        self.sheet, self.shared, self.styles, self.date1904, self.limit = sheet, shared, styles, date1904, limit
        self.masters: dict[str, tuple[str, int, int]] = {}
        self.cut: int | None = None
        self.last = 0
        # Ausgeblendete Zeilen (<row hidden="1">) und die davon in einer eingeklappten Gruppe (outlineLevel).
        self.hidden: set[int] = set()
        self.grouped: set[int] = set()
        # Zahlen, die ihr Format als 0 zeigt (0,4 als #,##0): {Zelle: Zahl}.
        self.rounded: dict = {}

    def row(self, element) -> None:
        cells = element.findall(f"{_NS}c")
        if self.cut is not None or not cells:
            return
        if _flag(element, "hidden") and (element.get("r") or "").isdigit() and len(self.hidden) < MAX_HIDDEN:
            self.hidden.add(int(element.get("r")))
            if (element.get("outlineLevel") or "0") not in ("", "0"):
                self.grouped.add(int(element.get("r")))
        found = []
        for cell in cells:
            match = ADDRESS.match(cell.get("r") or "")
            if not match:
                continue
            col, row = column_number(match.group(1)), int(match.group(2))
            value, text = _value(cell, self.shared, self.styles, self.date1904)
            style = int(cell.get("s") or 0)
            if style < len(self.styles) and self.styles[style] == "whole" and (cell.get("t") or "n") == "n":
                try:
                    number = float(cell.findtext(f"{_NS}v") or "")
                except ValueError:
                    number = 0.0
                if 0 < abs(number) < 0.5 and len(self.rounded) < MAX_HIDDEN:
                    self.rounded[(row, col)] = number
            formula = None
            node = cell.find(f"{_NS}f")
            if node is not None:
                formula = node.text or ""
                if node.get("t") == "shared" and node.get("si") is not None:
                    if formula:
                        self.masters[node.get("si")] = (formula, row, col)
                    elif node.get("si") in self.masters:
                        base, base_row, base_col = self.masters[node.get("si")]
                        formula = shift(base, row - base_row, col - base_col)
                formula = formula.strip()[:MAX_FORMULA] or None
            if value or formula is not None:
                found.append(((row, col), Cell(value, formula, text)))
        if not found:
            return
        if len(self.sheet.cells) + len(found) > self.limit:
            # Nur ganze Zeilen: Gelesen ist bis zur Zeile davor, 0 heißt gar nichts.
            self.cut = min(pos[0] for pos, _ in found) - 1 if self.sheet.cells else 0
            self.sheet.cells = Cells(self.sheet.cells, self.cut)
            return
        self.sheet.cells.update(found)


def _sheet_part(archive: zipfile.ZipFile, name: str, cells: _Cells) -> ElementTree.Element | None:
    """Ein Blatt der Datei: die Zellen beim Lesen in cells, zurück kommt der Rest (Filter, Formate, Verweise).
    So braucht auch ein sehr großes Blatt wenig Speicher."""
    info = _info(archive, name)
    if info is None:
        return None
    if info.file_size > MAX_SHEET_PART:
        raise ValueError(T("Die Excel-Datei ist zu groß.", "The Excel file is too large."))
    parser = ElementTree.XMLPullParser(events=("start", "end"))
    root = data = None
    with archive.open(info) as stream:
        chunk, before = stream.read(1 << 20), b""
        while chunk:
            # Eine DTD (<!DOCTYPE) stünde vor dem ersten Element: bis dahin jedes Stück ansehen, auch über die Grenze.
            if root is None and b"<!DOCTYPE" in before[-16:] + chunk:
                raise ValueError(T("Das ist keine Excel-Datei.", "This is not an Excel file."))
            before = chunk
            parser.feed(chunk)
            for event, element in parser.read_events():
                if event == "start":
                    if root is None:
                        root = element
                    elif element.tag == f"{_NS}sheetData":
                        data = element
                elif element.tag == f"{_NS}row" and data is not None:
                    cells.row(element)
                    data.remove(element)
            chunk = stream.read(1 << 20)
    parser.close()
    for event, element in parser.read_events():
        if event == "end" and element.tag == f"{_NS}row" and data is not None:
            cells.row(element)
            data.remove(element)
    return root


def _read_parts(archive: zipfile.ZipFile, root, path: str, sheet: Sheet) -> None:
    """Alles außer den Zellen. Ein Teil, das sich nicht lesen lässt, fehlt eben."""
    links = _rels(archive, path)
    readers = (
        ("Tabellen", lambda: _tables(archive, root, links, sheet)),
        ("Filter", lambda: _sheet_filter(root, sheet)),
        ("bedingte Formatierung", lambda: _formats(root, sheet)),
        ("Eingaberegeln", lambda: _validations(root, sheet)),
        ("Diagramme", lambda: _charts(archive, root, links, sheet)),
        ("Pivot-Tabellen", lambda: _pivots(archive, links, sheet)),
        ("verbundene Zellen", lambda: _merges(root, sheet)),
    )
    for what, read in readers:
        try:
            read()
        except (ValueError, KeyError, IndexError, ElementTree.ParseError):
            logger.info("%s in Blatt %s nicht gelesen", what, sheet.name, exc_info=True)


# So viele verbundene Bereiche liest der Prüfer je Blatt; ein Formular mit mehr ist keine Liste.
MAX_MERGED = 5000
# So viele ausgeblendete Zeilen merkt er sich je Blatt.
MAX_HIDDEN = 50_000


def _merges(root, sheet: Sheet) -> None:
    for item in root.iter(f"{_NS}mergeCell"):
        if len(sheet.merged) >= MAX_MERGED:
            break
        area = area_of(item.get("ref") or "")
        if area is not None and area[:2] != area[2:]:
            sheet.merged.append(area)


def _flag(node, name: str, default: bool = False) -> bool:
    value = node.get(name)
    return default if value is None else value in ("1", "true")


# Filter: Vergleich in der Datei → wie man ihn sagt.
_OPERATORS = {
    "equal": "=", "notEqual": "≠", "greaterThan": ">", "greaterThanOrEqual": "≥",
    "lessThan": "<", "lessThanOrEqual": "≤",
}
_DYNAMIC = {
    "aboveAverage": "über dem Durchschnitt", "belowAverage": "unter dem Durchschnitt",
    "today": "heute", "yesterday": "gestern", "tomorrow": "morgen",
    "thisWeek": "diese Woche", "lastWeek": "letzte Woche", "nextWeek": "nächste Woche",
    "thisMonth": "dieser Monat", "lastMonth": "letzter Monat", "nextMonth": "nächster Monat",
    "thisQuarter": "dieses Quartal", "lastQuarter": "letztes Quartal", "nextQuarter": "nächstes Quartal",
    "thisYear": "dieses Jahr", "lastYear": "letztes Jahr", "nextYear": "nächstes Jahr", "yearToDate": "seit Jahresbeginn",
}
_DYNAMIC_EN = {
    "aboveAverage": "above average", "belowAverage": "below average",
    "today": "today", "yesterday": "yesterday", "tomorrow": "tomorrow",
    "thisWeek": "this week", "lastWeek": "last week", "nextWeek": "next week",
    "thisMonth": "this month", "lastMonth": "last month", "nextMonth": "next month",
    "thisQuarter": "this quarter", "lastQuarter": "last quarter", "nextQuarter": "next quarter",
    "thisYear": "this year", "lastYear": "last year", "nextYear": "next year", "yearToDate": "year to date",
}


def _quote(values: list[str], limit: int = 6) -> Both:
    shown = [f"„{v}“" for v in values[:limit]]
    english = [f"“{v}”" for v in values[:limit]]
    if len(values) > limit:
        shown.append(f"{len(values) - limit} weitere")
        english.append(f"{len(values) - limit} more")
    return Both(", ".join(shown), ", ".join(english))


def _custom(operator: str, value: str) -> Both:
    if operator in ("", "equal") and value.startswith("*") and value.endswith("*") and len(value) > 2:
        return Both(f"enthält „{value[1:-1]}“", f"contains “{value[1:-1]}”")
    if operator == "notEqual" and value.startswith("*") and value.endswith("*") and len(value) > 2:
        return Both(f"enthält nicht „{value[1:-1]}“", f"does not contain “{value[1:-1]}”")
    if operator in ("", "equal") and value.endswith("*"):
        return Both(f"beginnt mit „{value[:-1]}“", f"begins with “{value[:-1]}”")
    if operator in ("", "equal") and value.startswith("*"):
        return Both(f"endet mit „{value[1:]}“", f"ends with “{value[1:]}”")
    if operator == "notEqual" and value == " ":
        return Both("nicht leer", "not empty")
    plain = re.fullmatch(r"-?\d+(\.\d+)?", value)
    shown = value if plain else f"„{value}“"
    english = value if plain else f"“{value}”"
    sign = _OPERATORS.get(operator or "equal", operator)
    return Both(f"{sign} {shown.replace('.', ',')}", f"{sign} {english.replace('.', ',')}")


def _filter_condition(node) -> str:
    """Eine Filterspalte als Text, z. B. „nur „ja“, „nein““ oder „> 5000“ (Both: auch auf Englisch)."""
    values = node.find(f"{_NS}filters")
    if values is not None:
        chosen = [v.get("val") or "" for v in values.findall(f"{_NS}filter")]
        chosen += [_date_group(g) for g in values.findall(f"{_NS}dateGroupItem")]
        english = list(chosen)
        if _flag(values, "blank"):
            chosen.append("(leer)")
            english.append("(empty)")
        if not chosen:
            return Both("nichts", "nothing")
        return Both(f"nur {_quote(chosen)}", f"only {_quote(english).en}")
    custom = node.find(f"{_NS}customFilters")
    if custom is not None:
        parts = [_custom(c.get("operator") or "", c.get("val") or "") for c in custom.findall(f"{_NS}customFilter")]
        both = _flag(custom, "and")
        return Both((" und " if both else " oder ").join(parts), (" and " if both else " or ").join(p.en for p in parts))
    top = node.find(f"{_NS}top10")
    if top is not None:
        count = number_text(float(top.get("val") or 10))
        big = _flag(top, "top", True)
        if _flag(top, "percent"):
            return Both(f"die {'obersten' if big else 'untersten'} {count} %", f"the {'top' if big else 'bottom'} {count} %")
        return Both(f"die {count} {'größten' if big else 'kleinsten'}", f"the {count} {'largest' if big else 'smallest'}")
    dynamic = node.find(f"{_NS}dynamicFilter")
    if dynamic is not None:
        kind = dynamic.get("type") or ""
        return Both(_DYNAMIC.get(kind, kind), _DYNAMIC_EN.get(kind, kind))
    if node.find(f"{_NS}colorFilter") is not None:
        return Both("nach Farbe", "by color")
    if node.find(f"{_NS}iconFilter") is not None:
        return Both("nach Symbol", "by icon")
    # Nur der Knopf ist aus- oder eingeblendet, gefiltert wird nichts.
    return ""


def _date_group(item) -> str:
    parts = [item.get(k) for k in ("day", "month", "year") if item.get(k)]
    return ".".join(parts)


def _auto_filter(node, sort_node=None, table: str = "") -> AutoFilter | None:
    area = area_of(node.get("ref") or "")
    if area is None:
        return None
    filters = []
    for column in node.findall(f"{_NS}filterColumn")[:MAX_PARTS]:
        condition = _filter_condition(column)
        if condition:
            filters.append(FilterColumn(area[1] + int(column.get("colId") or 0), condition))
    sort_node = node.find(f"{_NS}sortState") if node.find(f"{_NS}sortState") is not None else sort_node
    return AutoFilter(area, tuple(filters), _sort_keys(sort_node), table)


def _sort_keys(node) -> tuple[SortKey, ...]:
    if node is None:
        return ()
    keys = []
    for condition in node.findall(f"{_NS}sortCondition")[:MAX_PARTS]:
        area = area_of(condition.get("ref") or "")
        if area is None:
            continue
        by = {"cellColor": Both("Zellfarbe", "cell color"), "fontColor": Both("Schriftfarbe", "font color"),
              "icon": Both("Symbol", "icon")}.get(condition.get("sortBy") or "", "")
        keys.append(SortKey(area[1], _flag(condition, "descending"), by))
    return tuple(keys)


def _sheet_filter(root, sheet: Sheet) -> None:
    node = root.find(f"{_NS}autoFilter")
    if node is not None:
        found = _auto_filter(node, root.find(f"{_NS}sortState"))
        if found is not None:
            sheet.filters.append(found)
    elif root.find(f"{_NS}sortState") is not None:
        # Sortiert ohne Filterknöpfe (Daten → Sortieren).
        state = root.find(f"{_NS}sortState")
        area = area_of(state.get("ref") or "")
        if area is not None:
            sheet.filters.append(AutoFilter(area, (), _sort_keys(state)))


def _tables(archive, root, links, sheet: Sheet) -> None:
    parts = root.find(f"{_NS}tableParts")
    for part in parts if parts is not None else ():
        kind, path = links.get(part.get(f"{_REL}id") or "", ("", ""))
        table = _part(archive, path) if kind == "table" else None
        if table is None:
            continue
        area = area_of(table.get("ref") or "")
        if area is None:
            continue
        columns, calculated = [], []
        for column in table.iter(f"{_NS}tableColumn"):
            name = _clean(column.get("name") or "", 60)
            columns.append(name)
            formula = column.findtext(f"{_NS}calculatedColumnFormula")
            if formula:
                calculated.append((name, formula.strip()[:MAX_FORMULA]))
        name = _clean(table.get("displayName") or table.get("name") or "Tabelle", 60)
        header = (table.get("headerRowCount") or "1") != "0"
        totals = int(table.get("totalsRowCount") or 0)
        sheet.tables.append(ExcelTable(name, area, tuple(columns), header, totals, tuple(calculated)))
        node = table.find(f"{_NS}autoFilter")
        if node is not None:
            found = _auto_filter(node, table.find(f"{_NS}sortState"), name)
            if found is not None:
                sheet.filters.append(found)
        if len(sheet.tables) >= MAX_PARTS:
            break


def _areas(sqref: str) -> tuple[Area, ...]:
    return tuple(a for a in (area_of(part) for part in sqref.split()) if a is not None)


def _rule(node, areas) -> Rule:
    kind = node.get("type") or ""
    formulas = []
    for child in node:
        if child.tag in (f"{_NS}formula", f"{_NS}formula1", f"{_NS}formula2", f"{_XM}f"):
            text = child.text
        elif child.tag in (f"{_X14}formula1", f"{_X14}formula2"):
            text = child.findtext(f"{_XM}f")
        else:
            continue
        if text:
            formulas.append(text.strip()[:MAX_FORMULA])
    extra = ""
    if kind == "iconSet":
        icons = node.find(f"{_NS}iconSet")
        icons = icons if icons is not None else node.find(f"{_X14}iconSet")
        extra = (icons.get("iconSet") if icons is not None else "") or "3TrafficLights1"
    elif kind in ("containsText", "notContainsText", "beginsWith", "endsWith"):
        extra = node.get("text") or ""
    elif kind == "top10":
        extra = f"{'bottom' if _flag(node, 'bottom') else 'top'} {node.get('rank') or 10}{' %' if _flag(node, 'percent') else ''}"
    elif kind == "timePeriod":
        extra = node.get("timePeriod") or ""
    elif kind == "aboveAverage":
        extra = "above" if _flag(node, "aboveAverage", True) else "below"
    return Rule(areas, kind, node.get("operator") or "", tuple(formulas), extra)


def _formats(root, sheet: Sheet) -> None:
    for block in root.findall(f"{_NS}conditionalFormatting"):
        areas = _areas(block.get("sqref") or "")
        for rule in block.findall(f"{_NS}cfRule"):
            if areas and len(sheet.formats) < MAX_PARTS:
                sheet.formats.append(_rule(rule, areas))
    # Neuere Regeln (z. B. Symbolsätze mit Sternen) stehen in einer Erweiterung.
    for block in root.iter(f"{_X14}conditionalFormatting"):
        areas = _areas(block.findtext(f"{_XM}sqref") or "")
        for rule in block.findall(f"{_X14}cfRule"):
            # Datenbalken stehen auch oben, hier nur ihre Feinheiten.
            if areas and rule.get("type") != "dataBar" and len(sheet.formats) < MAX_PARTS:
                sheet.formats.append(_rule(rule, areas))


def _validations(root, sheet: Sheet) -> None:
    block = root.find(f"{_NS}dataValidations")
    for rule in block.findall(f"{_NS}dataValidation") if block is not None else ():
        areas = _areas(rule.get("sqref") or "")
        if areas and (rule.get("type") or "") not in ("", "none") and len(sheet.validations) < MAX_PARTS:
            sheet.validations.append(_rule(rule, areas))
    # Listen von einem anderen Blatt speichert Excel in einer Erweiterung.
    for rule in root.iter(f"{_X14}dataValidation"):
        areas = _areas(rule.findtext(f"{_XM}sqref") or "")
        if areas and (rule.get("type") or "") not in ("", "none") and len(sheet.validations) < MAX_PARTS:
            sheet.validations.append(_rule(rule, areas))


# Diagrammarten in der Datei → wie Excel sie auf Deutsch nennt.
_CHARTS = {
    "lineChart": "Liniendiagramm", "line3DChart": "Liniendiagramm", "pieChart": "Kreisdiagramm",
    "pie3DChart": "Kreisdiagramm", "ofPieChart": "Kreisdiagramm", "doughnutChart": "Ringdiagramm",
    "areaChart": "Flächendiagramm", "area3DChart": "Flächendiagramm", "scatterChart": "Punktdiagramm",
    "bubbleChart": "Blasendiagramm", "radarChart": "Netzdiagramm", "stockChart": "Kursdiagramm",
    "surfaceChart": "Oberflächendiagramm", "surface3DChart": "Oberflächendiagramm",
}
_CHART_EX = {
    "waterfall": "Wasserfalldiagramm", "clusteredColumn": "Histogramm", "paretoLine": "Pareto-Diagramm",
    "treemap": "Treemap", "sunburst": "Sunburst-Diagramm", "boxWhisker": "Kastengrafik",
    "funnel": "Trichterdiagramm", "regionMap": "Kartendiagramm",
}
# Und wie das englische Excel sie nennt.
_CHART_NAMES = {
    "Liniendiagramm": "line chart", "Kreisdiagramm": "pie chart", "Ringdiagramm": "doughnut chart",
    "Flächendiagramm": "area chart", "Punktdiagramm": "scatter chart", "Blasendiagramm": "bubble chart",
    "Netzdiagramm": "radar chart", "Kursdiagramm": "stock chart", "Oberflächendiagramm": "surface chart",
    "Wasserfalldiagramm": "waterfall chart", "Histogramm": "histogram", "Pareto-Diagramm": "Pareto chart",
    "Treemap": "treemap", "Sunburst-Diagramm": "sunburst chart", "Kastengrafik": "box and whisker chart",
    "Trichterdiagramm": "funnel chart", "Kartendiagramm": "map chart", "Balkendiagramm": "bar chart",
    "Säulendiagramm": "column chart", "Diagramm": "chart",
}


def chart_kind(kinds: list[str]) -> Both:
    """„Säulendiagramm (gestapelt) und Liniendiagramm“, auch auf Englisch."""
    def english(kind: str) -> str:
        base, _, rest = kind.partition(" (")
        stacked = {"gestapelt)": " (stacked)", "100 % gestapelt)": " (100 % stacked)"}.get(rest, f" ({rest}" if rest else "")
        return _CHART_NAMES.get(base, base) + stacked

    kinds = kinds or ["Diagramm"]
    return Both(" und ".join(kinds), " and ".join(english(k) for k in kinds))


def _charts(archive, root, links, sheet: Sheet) -> None:
    for drawing in root.iter(f"{_NS}drawing"):
        kind, path = links.get(drawing.get(f"{_REL}id") or "", ("", ""))
        tree = _part(archive, path) if kind == "drawing" else None
        if tree is None:
            continue
        drawing_links = _rels(archive, path)
        for anchor in tree:
            names = [n.get("name") for n in anchor.iter(f"{_XDR}cNvPr") if n.get("name")]
            for node in anchor.iter():
                if node.tag.rsplit("}", 1)[-1] != "chart" or node.get(f"{_REL}id") is None:
                    continue
                chart_kind, chart_path = drawing_links.get(node.get(f"{_REL}id"), ("", ""))
                chart = _part(archive, chart_path) if chart_kind in ("chart", "chartEx") else None
                if chart is not None and len(sheet.charts) < MAX_PARTS:
                    name = _clean(names[0] if names else f"Diagramm {len(sheet.charts) + 1}", 60)
                    sheet.charts.append(_chart(chart, name))
                break


def _chart(root, name: str) -> Chart:
    if root.tag == f"{_CX}chartSpace":
        layouts = [s.get("layoutId") or "" for s in root.iter(f"{_CX}series")]
        kind = chart_kind([_CHART_EX.get(layouts[0], "Diagramm") if layouts else "Diagramm"])
        title = root.find(f".//{_CX}title")
        series = tuple(Series(_clean("".join(v.text or "" for v in s.iter(f"{_CX}v")), 60)) for s in root.iter(f"{_CX}series"))
        return Chart(name, kind, _clean("".join(t.text or "" for t in title.iter(f"{_A}t")), 80) if title is not None else "", series)
    plot = root.find(f"{_C}chart/{_C}plotArea")
    kinds, series = [], []
    for group in plot if plot is not None else ():
        tag = group.tag.rsplit("}", 1)[-1]
        if tag in ("barChart", "bar3DChart"):
            direction = group.find(f"{_C}barDir")
            kind = "Balkendiagramm" if direction is not None and direction.get("val") == "bar" else "Säulendiagramm"
        elif tag in _CHARTS:
            kind = _CHARTS[tag]
        else:
            continue
        grouping = group.find(f"{_C}grouping")
        if grouping is not None and grouping.get("val") in ("stacked", "percentStacked"):
            kind += " (gestapelt)" if grouping.get("val") == "stacked" else " (100 % gestapelt)"
        if kind not in kinds:
            kinds.append(kind)
        for ser in group.findall(f"{_C}ser"):
            series.append(_series(ser))
    title = root.find(f"{_C}chart/{_C}title")
    text = ""
    if title is not None:
        text = "".join(t.text or "" for t in title.iter(f"{_A}t")) or "".join(v.text or "" for v in title.iter(f"{_C}v"))
    return Chart(name, chart_kind(kinds), _clean(text, 80), tuple(series[:MAX_PARTS]))


def _series(ser) -> Series:
    def ref(*paths: str) -> str:
        for path in paths:
            found = ser.findtext(path)
            if found:
                return found.strip()
        return ""

    name = ser.findtext(f"{_C}tx/{_C}v") or ""
    if not name:
        cached = ser.find(f"{_C}tx/{_C}strRef/{_C}strCache")
        name = "".join(v.text or "" for v in cached.iter(f"{_C}v")) if cached is not None else ""
    return Series(
        _clean(name, 60),
        ref(f"{_C}tx/{_C}strRef/{_C}f"),
        ref(f"{_C}val/{_C}numRef/{_C}f", f"{_C}yVal/{_C}numRef/{_C}f"),
        ref(f"{_C}cat/{_C}strRef/{_C}f", f"{_C}cat/{_C}numRef/{_C}f", f"{_C}cat/{_C}multiLvlStrRef/{_C}f",
            f"{_C}xVal/{_C}numRef/{_C}f", f"{_C}xVal/{_C}strRef/{_C}f"),
    )


def _pivots(archive, links, sheet: Sheet) -> None:
    for kind, path in links.values():
        if kind != "pivotTable" or len(sheet.pivots) >= MAX_PARTS:
            continue
        table = _part(archive, path)
        if table is None:
            continue
        cache = None
        for cache_kind, cache_path in _rels(archive, path).values():
            if cache_kind == "pivotCacheDefinition":
                cache = _part(archive, cache_path)
        sheet.pivots.append(_pivot(table, cache))


def _item(node) -> str:
    """Ein Eintrag eines Pivot-Felds: Text, Zahl, Datum oder leer."""
    if node.tag == f"{_NS}m":
        return "(leer)"
    value = node.get("v") or ""
    date = _ISO_DATE.match(value)
    if date:
        return f"{date.group(3)}.{date.group(2)}.{date.group(1)}"
    if node.tag == f"{_NS}n":
        try:
            return number_text(float(value))
        except ValueError:
            pass
    return _clean(value, 60)


# Wie Excel die Zusammenfassung eines Wertfelds in der Datei nennt (ohne Angabe: Summe).
_SUBTOTALS = {"sum", "count", "average", "max", "min", "product", "countNums", "stdDev", "stdDevp", "var", "varp"}


def _pivot(table, cache) -> Pivot:
    fields, items = [], []
    source = ""
    special: set[int] = set()
    origin = None
    if cache is not None:
        for index, cache_field in enumerate(cache.iter(f"{_NS}cacheField")):
            fields.append(_clean(cache_field.get("name") or "", 60))
            shared = cache_field.find(f"{_NS}sharedItems")
            items.append([_item(i) for i in shared] if shared is not None else [])
            # Berechnet, gruppiert (Monate, Bereiche) oder aus einer Gruppierung entstanden.
            if cache_field.get("formula") or cache_field.find(f"{_NS}fieldGroup") is not None or \
                    cache_field.get("databaseField") in ("0", "false"):
                special.add(index)
        origin = cache.find(f"{_NS}cacheSource/{_NS}worksheetSource")
        if origin is not None:
            source = origin.get("name") or (f"{origin.get('sheet')}!{origin.get('ref')}" if origin.get("sheet") else origin.get("ref") or "")

    def field_name(index: int) -> str:
        return fields[index] if 0 <= index < len(fields) else f"Feld {index + 1}"

    def axis(tag: str, child: str, attribute: str) -> tuple[str, ...]:
        node = table.find(f"{_NS}{tag}")
        names = []
        for item in node.findall(f"{_NS}{child}") if node is not None else ():
            index = int(item.get(attribute) or 0)
            if index >= 0:  # -2 ist die Zeile oder Spalte „Werte“.
                names.append(field_name(index))
        return tuple(names)

    hidden, skipped = [], {}
    pivot_fields = table.find(f"{_NS}pivotFields")
    for index, pivot_field in enumerate(pivot_fields if pivot_fields is not None else ()):
        values = items[index] if index < len(items) else []
        off = []
        for item in pivot_field.iter(f"{_NS}item"):
            if _flag(item, "h") and item.get("x") is not None:
                position = int(item.get("x") or 0)
                off.append(values[position] if position < len(values) else f"Eintrag {position + 1}")
        if off:
            listed = _quote(off, 4)
            hidden.append(Both(f"{field_name(index)} ohne {listed}", f"{field_name(index)} without {listed.en}"))
            skipped.setdefault(field_name(index), set()).update(off)
    filters = []
    pages = table.find(f"{_NS}pageFields")
    for page in pages.findall(f"{_NS}pageField") if pages is not None else ():
        index = int(page.get("fld") or 0)
        chosen = page.get("item")
        values = items[index] if index < len(items) else []
        if chosen is not None and chosen.isdigit() and int(chosen) < len(values):
            filters.append(Both(f"{field_name(index)} = „{values[int(chosen)]}“", f"{field_name(index)} = “{values[int(chosen)]}”"))
            skipped.setdefault(field_name(index), set()).update(v for i, v in enumerate(values) if i != int(chosen))
        else:
            filters.append(field_name(index))
    data = table.find(f"{_NS}dataFields")
    data_fields = data.findall(f"{_NS}dataField") if data is not None else []
    values = tuple(_clean(d.get("name") or field_name(int(d.get("fld") or 0)), 60) for d in data_fields)
    value_fields = []
    for d, caption in zip(data_fields, values):
        index = int(d.get("fld") or 0)
        function = d.get("subtotal") or "sum"
        value_fields.append(PivotValue(caption, field_name(index), function if function in _SUBTOTALS and index not in special else "",
                                       (d.get("showDataAs") or "normal") != "normal"))
    location = table.find(f"{_NS}location")
    area = area_of(location.get("ref") or "") if location is not None else None
    body = None
    if area is not None:
        try:
            first_row, first_col = int(location.get("firstDataRow") or 0), int(location.get("firstDataCol") or 0)
        except ValueError:
            first_row, first_col = -1, -1
        if 0 <= first_row and area[0] + first_row <= area[2] and 0 <= first_col and area[1] + first_col <= area[3]:
            body = (area[0] + first_row, area[1] + first_col, area[2], area[3])
    used = {int(x.get("x") or 0) for tag in ("rowFields", "colFields") for node in table.iter(f"{_NS}{tag}")
            for x in node.findall(f"{_NS}field")} | {int(d.get("fld") or 0) for d in data_fields}
    plain = not (used & special) and not _flag(table, "dataOnRows") and \
        not any(table.find(f"{_NS}{tag}") is not None for tag in ("calculatedItems", "calculatedMembers"))
    return Pivot(
        _clean(table.get("name") or "PivotTable", 60),
        area,
        axis("rowFields", "field", "x"),
        axis("colFields", "field", "x"),
        values,
        tuple(filters),
        tuple(hidden),
        _clean(source, 80),
        tuple(value_fields),
        body,
        _clean(origin.get("sheet") or "", 40) if origin is not None and not origin.get("name") else "",
        area_of(origin.get("ref") or "") if origin is not None and not origin.get("name") and origin.get("ref") else None,
        _clean(origin.get("name") or "", 60) if origin is not None else "",
        tuple((name, tuple(sorted(off))) for name, off in skipped.items()),
        (_flag(table, "rowGrandTotals", True), _flag(table, "colGrandTotals", True)),
        table.find(f"{_NS}filters") is not None,
        plain,
    )
