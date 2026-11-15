"""Tabellen aus Dateien holen, damit der Excel-Helfer sie in ein neues Blatt schreibt.

CSV und Text (Trenner ; , Tab | werden erkannt), Excel (.xlsx, .xlsm), Word (.docx, jede Tabelle)
und einfache PDF-Tabellen (Text-PDF, keine Scans). Alles mit Python allein, ohne KI und ohne
neues Paket; die Datei bleibt im Arbeitsspeicher und wird nirgends gespeichert.

Zahlen werden Zahlen („1.234,50“, „1'234.50“, „12,5 %“, „CHF 80.–“), Datumsangaben werden Datum,
und was nur wie eine Zahl aussieht (führende Nullen wie bei Postleitzahlen und Kundennummern,
sehr lange Ziffernfolgen wie IBAN), bleibt Text.
"""

from __future__ import annotations

import csv
import io
import re
import zipfile
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from xml.etree import ElementTree

from .pivot import Draft, sheet_name, show
from .sprache import T
from .xlsx import MAX_CELLS as MAX_CHECK_CELLS
from .xlsx import Cell, Sheet, _as_text, _clean, column_letters, date_text, number_text, read_book

# Mehr Zellen schreibt AppleScript nicht in vernünftiger Zeit.
MAX_CELLS = 50_000
MAX_FILE = 30_000_000
PREVIEW_ROWS = 8
SUFFIXES = (".csv", ".tsv", ".txt", ".xlsx", ".xlsm", ".docx", ".pdf")

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_CURRENCY = re.compile(r"^(?:CHF|EUR|USD|Fr\.|SFr\.|€|\$|£)\s*|\s*(?:CHF|EUR|USD|Fr\.|SFr\.|€|\$|£)$", re.I)
_DATE_DE = re.compile(r"^(\d{1,2})\.(\d{1,2})\.(\d{4}|\d{2})$")
_DATE_ISO = re.compile(r"^(\d{4})-(\d{2})-(\d{2})(?:[T ]00:00(?::00)?)?$")
_DATE_SLASH = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$")


@dataclass
class Table:
    name: str
    rows: list[list[str]]
    # CSV, Excel, Word oder PDF
    kind: str

    @property
    def size(self) -> tuple[int, int]:
        return len(self.rows), max((len(r) for r in self.rows), default=0)


def tables_from(name: str, data: bytes) -> list[Table]:
    """Alle Tabellen der Datei, mit mindestens zwei Spalten oder zwei Zeilen."""
    if len(data) > MAX_FILE:
        raise ValueError(T("Die Datei ist zu groß (mehr als 30 MB).", "The file is too large (more than 30 MB)."))
    suffix = "." + name.rsplit(".", 1)[-1].lower() if "." in name else ""
    if suffix in (".csv", ".tsv", ".txt"):
        found = [_csv(name, data)]
    elif suffix in (".xlsx", ".xlsm"):
        found = _xlsx(data)
    elif suffix == ".docx":
        found = _docx(data)
    elif suffix == ".pdf":
        found = _pdf(data)
    elif suffix in (".xls", ".doc"):
        raise ValueError(T("Alte Office-Dateien (.xls, .doc) kann ich nicht lesen. Speichere sie als .xlsx oder .docx.",
                           "I cannot read old Office files (.xls, .doc). Save them as .xlsx or .docx."))
    else:
        raise ValueError(T("Ich lese CSV, Excel (.xlsx), Word (.docx) und PDF mit Tabellen.",
                           "I read CSV, Excel (.xlsx), Word (.docx) and PDF with tables."))
    found = [_trim(t) for t in found]
    found = [t for t in found if t.rows and (t.size[0] >= 2 or t.size[1] >= 2)]
    if not found:
        raise ValueError(T("In der Datei finde ich keine Tabelle.", "I cannot find a table in the file."))
    return found


def _decode(data: bytes) -> str:
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1")


def _csv(name: str, data: bytes) -> Table:
    text = _decode(data)
    sample = "\n".join(text.splitlines()[:30])
    counts = {d: sum(line.count(d) for line in sample.splitlines()) for d in (";", "\t", ",", "|")}
    delimiter = max(counts, key=lambda d: (counts[d], d == ";")) if any(counts.values()) else ";"
    rows = [[cell.strip() for cell in row] for row in csv.reader(io.StringIO(text), delimiter=delimiter)]
    return Table(name.rsplit(".", 1)[0], rows, "CSV")


TEXT_SUFFIXES = (".csv", ".tsv", ".txt")


def is_text(name: str) -> bool:
    """Eine CSV- oder Textdatei (Export aus KIS, SAP, Statistik), keine Excel-Mappe."""
    return name.lower().endswith(TEXT_SUFFIXES)


def csv_sheets(name: str, data: bytes) -> list[Sheet]:
    """Eine CSV- oder Textdatei als ein Blatt für den Prüfer, so wie Excel sie öffnen sollte: Zahlen als Zahlen (mit
    dem Dezimalzeichen, das in der Datei überwiegt), Daten als Daten, alles andere als Text, auch Nummern mit Null
    vorn. Formeln gibt es in einer CSV nicht; geprüft werden die Daten."""
    if len(data) > MAX_FILE:
        raise ValueError(T("Die Datei ist zu groß (mehr als 30 MB).", "The file is too large (more than 30 MB)."))
    if not data.strip():
        raise ValueError(T("Die Datei ist leer.", "The file is empty."))
    if b"\x00" in data[:4096]:
        raise ValueError(T("Das ist keine Textdatei. Ich lese CSV-Dateien und Excel-Mappen (.xlsx).",
                           "This is not a text file. I read CSV files and Excel workbooks (.xlsx)."))
    table = _trim(_csv(name, data))
    mark = decimal_mark(table.rows)
    cells: dict[tuple[int, int], Cell] = {}
    for r, row in enumerate(table.rows, start=1):
        if len(cells) + len(row) > MAX_CHECK_CELLS:
            break
        for c, text in enumerate(row, start=1):
            if text.strip():
                cells[(r, c)] = _cell(text, mark)
    if not cells:
        raise ValueError(T("In der Datei steht nichts.", "There is nothing in the file."))
    base = name.rsplit("/", 1)[-1].rsplit(".", 1)[0]
    return [Sheet(sheet_name(base, []), cells)]


def _cell(text: str, mark: str) -> Cell:
    found = number(text, mark)
    if found is not None:
        value, percent = found
        return Cell(f"{number_text(round(value * 100, 10))} %" if percent else number_text(value))
    day = date(text)
    if day is not None:
        try:
            return Cell(date_text(datetime(*day)))
        except ValueError:
            pass  # 31.02.2026: kein Tag, bleibt Text
    return Cell(_clean(text), None, _as_text(text))


def _xlsx(data: bytes) -> list[Table]:
    try:
        book = read_book(data)
    except (ValueError, zipfile.BadZipFile, KeyError) as err:
        raise ValueError(T("Die Excel-Datei lässt sich nicht lesen.", "The Excel file cannot be read.")) from err
    tables = []
    for sheet in book.sheets:
        if not sheet.cells:
            continue
        rows_used = [r for r, _ in sheet.cells]
        cols_used = [c for _, c in sheet.cells]
        top, left, bottom, right = min(rows_used), min(cols_used), max(rows_used), max(cols_used)
        rows = [[sheet.cells[(r, c)].value if (r, c) in sheet.cells else "" for c in range(left, right + 1)]
                for r in range(top, bottom + 1)]
        tables.append(Table(sheet.name, rows, "Excel"))
    return tables


def _docx(data: bytes) -> list[Table]:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            root = ElementTree.fromstring(archive.read("word/document.xml"))
    except (zipfile.BadZipFile, KeyError, ElementTree.ParseError) as err:
        raise ValueError(T("Die Word-Datei lässt sich nicht lesen.", "The Word file cannot be read.")) from err
    tables = []
    for number, table in enumerate(root.iter(f"{_W}tbl"), start=1):
        rows = []
        for row in table.findall(f"{_W}tr"):
            cells = []
            for cell in row.findall(f"{_W}tc"):
                text = " ".join("".join(t.text or "" for t in p.iter(f"{_W}t")) for p in cell.iter(f"{_W}p")).strip()
                span = cell.find(f"{_W}tcPr/{_W}gridSpan")
                cells.append(text)
                if span is not None and str(span.get(f"{_W}val", "1")).isdigit():
                    # Verbundene Zellen: Die Spalten dahinter bleiben leer, damit alles untereinander steht.
                    cells += [""] * (int(span.get(f"{_W}val")) - 1)
            rows.append(cells)
        tables.append(Table(f"Tabelle {number}", rows, "Word"))
    return tables


def _pdf(data: bytes) -> list[Table]:
    try:
        from pypdf import PdfReader
        from pypdf.errors import PdfReadError
    except ImportError as err:
        raise ValueError(T("Zum Lesen von PDF fehlt ein Paket. Starte Gegenprobe einmal neu, dann wird es installiert.",
                           "A package for reading PDF is missing. Restart Gegenprobe once, then it gets installed.")) from err
    try:
        reader = PdfReader(io.BytesIO(data))
        if reader.is_encrypted:
            raise ValueError(T("Das PDF ist mit einem Passwort geschützt.", "The PDF is protected with a password."))
        pages = [page.extract_text(extraction_mode="layout") or "" for page in reader.pages]
    except (PdfReadError, KeyError, TypeError, OSError) as err:
        raise ValueError(T("Das PDF lässt sich nicht lesen.", "The PDF cannot be read.")) from err
    tables = []
    for page_number, text in enumerate(pages, start=1):
        for block in _pdf_blocks(text):
            tables.append(Table(f"Seite {page_number} Tabelle {len(tables) + 1}", block, "PDF"))
    if not tables and not any(p.strip() for p in pages):
        raise ValueError(T("Im PDF steht kein Text. Ist es eingescannt? Dann geht es so nicht.",
                           "There is no text in the PDF. Is it scanned? Then this does not work."))
    return tables


def _pdf_blocks(text: str) -> list[list[list[str]]]:
    """Zeilen, in denen zwei oder mehr Leerzeichen Spalten trennen; mindestens drei solche Zeilen hintereinander."""
    blocks, current = [], []
    for line in text.splitlines():
        parts = [p.strip() for p in re.split(r"\s{2,}", line.strip()) if p.strip()]
        if len(parts) >= 2:
            current.append(parts)
            continue
        if len(current) >= 3:
            blocks.append(current)
        current = []
    if len(current) >= 3:
        blocks.append(current)
    result = []
    for block in blocks:
        width = Counter(len(r) for r in block).most_common(1)[0][0]
        rows = [r for r in block if len(r) == width]
        if len(rows) >= 3:
            result.append(rows)
    return result


def _trim(table: Table) -> Table:
    """Leere Zeilen und Spalten am Rand weg, alle Zeilen gleich breit."""
    rows = [list(r) for r in table.rows]
    while rows and not any(c.strip() for c in rows[-1]):
        rows.pop()
    while rows and not any(c.strip() for c in rows[0]):
        rows.pop(0)
    width = max((len(r) for r in rows), default=0)
    rows = [r + [""] * (width - len(r)) for r in rows]
    keep = [j for j in range(width) if any(r[j].strip() for r in rows)]
    if keep:
        rows = [[r[j] for j in range(keep[0], keep[-1] + 1)] for r in rows]
    return Table(table.name, rows, table.kind)


# --- Werte so, wie Excel sie bekommt ---


def decimal_mark(rows: list[list[str]]) -> str:
    """Komma oder Punkt als Dezimalzeichen, wie es in der Datei überwiegt."""
    comma = sum(1 for r in rows for c in r if re.fullmatch(r"-?[\d.' ]*\d,\d{1,2}|-?\d+,\d+", _bare(c)))
    point = sum(1 for r in rows for c in r if re.fullmatch(r"-?[\d,' ]*\d\.\d{1,2}|-?\d+\.\d+", _bare(c)) and not _DATE_DE.match(c))
    return "," if comma >= point and comma else "." if point else ","


def _bare(text: str) -> str:
    text = _CURRENCY.sub("", text.strip()).strip()
    return text[:-2] if text.endswith((".–", ".-")) else text


def number(text: str, mark: str) -> tuple[float, bool] | None:
    """Die Zahl in einem Text (und ob es Prozent sind), None wenn es keine ist oder Text bleiben soll."""
    raw = _bare(text)
    percent = raw.endswith("%")
    raw = raw.rstrip("%").strip().replace("−", "-").replace("'", "").replace("’", "").replace(" ", "").replace(" ", "")
    if not raw or not re.fullmatch(r"[-+]?[\d.,]+", raw):
        return None
    digits = raw.lstrip("+-")
    if len(digits.replace(".", "").replace(",", "")) > 15:
        return None  # IBAN, Versichertennummer …: als Zahl ginge die Genauigkeit verloren
    if re.fullmatch(r"0\d+", digits):
        return None  # Postleitzahl, Kundennummer: die Null vorne bleibt
    thousands = "." if mark == "," else ","
    if thousands in digits:
        if not re.fullmatch(rf"\d{{1,3}}(\{thousands}\d{{3}})+(\{mark}\d+)?", digits):
            return None
        digits = digits.replace(thousands, "")
    if digits.count(mark) > 1:
        return None
    try:
        value = float(digits.replace(mark, "."))
    except ValueError:
        return None
    value = -value if raw.startswith("-") else value
    return (value / 100 if percent else value), percent


def date(text: str) -> tuple[int, int, int] | None:
    text = text.strip()
    match = _DATE_DE.match(text) or _DATE_SLASH.match(text)
    if match:
        day, month, year = (int(g) for g in match.groups())
        year += 2000 if year < 100 and year < 70 else 1900 if year < 100 else 0
    else:
        match = _DATE_ISO.match(text)
        if not match:
            return None
        year, month, day = (int(g) for g in match.groups())
    if not (1 <= month <= 12 and 1 <= day <= 31 and 1900 <= year <= 2200):
        return None
    return year, month, day


def _text_entry(text: str) -> str:
    if not text:
        return ""
    if text[0] in "=+-@'#" or any(ch.isdigit() for ch in text) or text.upper() in ("TRUE", "FALSE", "WAHR", "FALSCH"):
        return "'" + text
    return text


def _number_entry(value: float) -> str:
    return str(int(value)) if value == int(value) and abs(value) < 1e15 else repr(value)


def draft(table: Table, existing: list[str], file_name: str) -> Draft:
    """Das neue Blatt für eine Tabelle: Kopfzeile fett, Zahlen als Zahlen, Datum als Datum."""
    height, width = table.size
    if height * width > MAX_CELLS:
        raise ValueError(T(f"Die Tabelle hat {height * width} Zellen, mehr als {MAX_CELLS}. So große Dateien öffnest du "
                           "besser direkt in Excel (Datei, Öffnen).",
                           f"The table has {height * width} cells, more than {MAX_CELLS}. Files this large are better "
                           "opened directly in Excel (File, Open)."))
    mark = decimal_mark(table.rows[1:])
    grid, expected, formats = [], {}, []
    numbers = dates = kept = 0
    date_columns, percent_columns = set(), set()
    for i, row in enumerate(table.rows, start=1):
        line = []
        for j, value in enumerate(row, start=1):
            value = " ".join(value.split())
            if i == 1 or not value:
                line.append(_text_entry(value))
                if value:
                    expected[(i, j)] = value
                continue
            day = date(value)
            found = number(value, mark) if day is None else None
            if day is not None:
                year, month, dom = day
                line.append(f"{year:04d}-{month:02d}-{dom:02d}")
                expected[(i, j)] = f"{dom:02d}.{month:02d}.{year:04d}"
                date_columns.add(j)
                dates += 1
            elif found is not None:
                amount, percent = found
                line.append(f"{show_entry(amount * 100)}%" if percent else _number_entry(amount))
                expected[(i, j)] = amount
                if percent:
                    percent_columns.add(j)
                numbers += 1
            else:
                if re.fullmatch(r"0\d+", value.strip()):
                    kept += 1
                line.append(_text_entry(value))
                expected[(i, j)] = value
        grid.append(line)
    last = column_letters(width)
    for j in sorted(date_columns):
        formats.append((f"{column_letters(j)}2:{column_letters(j)}{height}", "dd.mm.yyyy"))
    base = table.name if table.kind in ("Excel", "CSV") else f"{file_name.rsplit('.', 1)[0]} {table.name}"
    name = sheet_name(base, existing)
    notes = []
    if numbers:
        notes.append(T(f"{numbers} Zahlen werden Zahlen ({'Komma' if mark == ',' else 'Punkt'} als Dezimalzeichen).",
                       f"{numbers} numbers become numbers ({'comma' if mark == ',' else 'point'} as decimal separator)."))
    if dates:
        notes.append(T(f"{dates} Datumsangaben werden Datum.", f"{dates} dates become dates."))
    if kept:
        notes.append(T(f"{kept} Werte mit einer Null vorne (z. B. Postleitzahlen) bleiben Text.",
                       f"{kept} values with a leading zero (e.g. postal codes) stay text."))
    text = T(f"Ich lege das neue Blatt „{name}“ an und schreibe die {table.kind}-Tabelle „{table.name}“ hinein: "
             f"{height} Zeilen und {width} Spalten, die erste Zeile als Kopfzeile. Deine bestehenden Zellen ändere ich nicht.",
             f"I create the new sheet “{name}” and write the {table.kind} table “{table.name}” into it: "
             f"{height} rows and {width} columns, the first row as the header row. I do not change your existing cells.")
    preview = [[show(v) if isinstance(v, float) else v for v in row] for row in table.rows[:PREVIEW_ROWS]]
    if height > PREVIEW_ROWS:
        preview.append(["…"] + [""] * (width - 1))
    return Draft(name=name, title=table.name, grid=grid, expected=expected, preview=preview, text=text, notes=notes,
                 bold=[f"A1:{last}1"], formats=formats)


def show_entry(value: float) -> str:
    """Eine Zahl für Range.Formula, ohne Rundungsreste (12.5 statt 12.499999999)."""
    rounded = round(value, 10)
    return _number_entry(rounded)
