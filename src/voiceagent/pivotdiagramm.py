"""Pivot-Tabellen und Diagramme, die nicht mehr zu ihren Daten passen.

- Eine Pivot-Tabelle, die nach einer Änderung der Daten nicht aktualisiert wurde. Der Prüfer rechnet jede Zahl aus den
  Daten nach (Summe, Anzahl, Mittelwert … je Eintrag, mit den Gesamtergebnissen) und nennt, was jetzt herauskäme,
  welcher Eintrag fehlt und welcher in den Daten gar nicht mehr vorkommt.
- Eine feste Datenquelle (Daten!$A$1:$D$100), unter der die Daten weitergehen: Diese Zeilen fehlen auch nach dem
  Aktualisieren.
- „Anzahl von Umsatz“: Excel zählt statt zu summieren, sobald in der Spalte beim Anlegen eine leere Zelle oder Text
  stand.
- Ein Diagramm, dessen Reihe vor dem Ende der Tabelle aufhört (oder zu spät anfängt), die Summenzeile als eigenen
  Punkt mitnimmt oder gegen ihre Beschriftungen verschoben ist, und eine Reihe, deren Name aus einer anderen Spalte
  kommt als ihre Werte.

Im Seitenfenster in Excel aktualisiert ein Knopf die Pivot-Tabelle, nimmt die Summe statt der Anzahl oder setzt die
Bereiche der Reihen neu (mit Rückgängig). Alles ohne KI, nur aus dem, was in der Mappe steht.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, replace

from . import formel
from .analyse import Finding, address, calc_value, value_of
from .bereiche import _num, _words
from .formel import Err
from .sprache import T
from .xlsx import Cell, Chart, ExcelTable, Pivot, PivotValue, Sheet, area_text, column_letters, column_number

Pos = tuple[int, int]
Area = tuple[int, int, int, int]

MAX_LISTED = 3
MAX_ROWS = 100000
# Wörter einer Summenzeile („Summe“, „Total 2025“, „Gesamtergebnis“).
_TOTAL = {"summe", "summen", "total", "totals", "gesamt", "gesamtsumme", "gesamtergebnis", "insgesamt", "endsumme",
          "gesamtbetrag", "sum", "grand"}
# „(Leer)“: So zeigt eine Pivot-Tabelle leere Zellen der Daten.
_BLANK = {"(leer)", "(blank)", "(vide)", "(vuoto)"}
# Ein Titel, der einen Zeitraum oder einen Teil nennt: Dann kann ein kürzerer Bereich gewollt sein.
_PERIOD = {"januar", "februar", "märz", "maerz", "april", "mai", "juni", "juli", "august", "september", "oktober",
           "november", "dezember", "jan", "feb", "mär", "mrz", "apr", "jun", "jul", "aug", "sep", "sept", "okt", "nov",
           "dez", "quartal", "q1", "q2", "q3", "q4", "halbjahr", "semester", "bis", "seit", "ab", "erste", "ersten",
           "letzte", "letzten", "kw", "woche", "wochen", "auszug", "ohne", "nur", "top"}
_YEAR = re.compile(r"^(19|20)\d\d$")
# Spalten, die man zählt, nicht summiert.
_COUNTED = {"nr", "nummer", "id", "fall", "fallnummer", "fallnr", "patient", "patientin", "pid", "kennung", "code",
            "jahr", "plz", "index", "lfd", "kunde", "kundennummer", "auftrag", "auftragsnummer", "rechnung",
            "rechnungsnummer", "artikel", "artikelnummer"}
_COUNT_CAPTION = re.compile(r"^(anzahl|count)\s+(von|of)\s+(?P<field>.+)$", re.IGNORECASE)
_SUM_FUNCTIONS = {"SUM", "SUBTOTAL", "AGGREGATE"}
# So heißt in Excel das Feld „Σ Werte“: wo die Wertfelder stehen. Office.js zählt es zu den Spalten (oder Zeilen) der
# Pivot-Tabelle, auch bei nur einem Wertfeld, es ist aber keine Spalte der Daten.
_VALUES_FIELD = {"werte", "values", "valeurs", "valori", "valores", "waarden", "värden", "verdier", "værdier", "arvot",
                 "wartości", "hodnoty", "değerler", "значения"}


class _Any:
    """Was Excel bei einem leeren Eintrag zeigt, ist hier nicht sicher (0, leer oder #DIV/0!): nicht vergleichen."""


_ANY = _Any()


class _Empty:
    """Zeilen gibt es, aber keine Zahl darin: Excel zeigt nichts, 0 oder #DIV/0!, aber keine andere Zahl."""


_EMPTY = _Empty()


# --- die Mappe ---


class _Book:
    """Die Werte aller Blätter so, wie Excel mit ihnen rechnet, und die Excel-Tabellen nach Namen."""

    def __init__(self, cells_of: dict[str, dict[Pos, Cell]], tables: dict[str, tuple[str, ExcelTable]]) -> None:
        self._cells = {name.casefold(): cells for name, cells in cells_of.items()}
        self._names = {name.casefold(): name for name in cells_of}
        self.tables = {name.casefold(): item for name, item in tables.items()}
        self._typed: dict[str, dict[Pos, object]] = {}

    def has(self, sheet: str) -> bool:
        return sheet.casefold() in self._cells

    def name(self, sheet: str) -> str:
        return self._names.get(sheet.casefold(), sheet)

    def cell(self, sheet: str, pos: Pos) -> Cell | None:
        return self._cells.get(sheet.casefold(), {}).get(pos)

    def shown(self, sheet: str, pos: Pos) -> str:
        """So, wie Excel die Zelle zeigt (Daten als 01.03.2025, nicht als Zahl)."""
        cell = self.cell(sheet, pos)
        return cell.value.strip() if cell is not None else ""

    def value(self, sheet: str, pos: Pos):
        key = sheet.casefold()
        if key not in self._typed:
            self._typed[key] = {p: calc_value(value_of(c), c) for p, c in self._cells.get(key, {}).items()}
        return self._typed[key].get(pos)


def _ref(text: str, home: str) -> tuple[str, Area] | None:
    """„=Tabelle1!$B$2:$B$10“, „'Mein Blatt'!B2:B10“ oder „$B$2:$B$10“ (auf dem eigenen Blatt) → Blatt und Bereich.
    Mehrere Bereiche oder Bezüge in andere Mappen: None."""
    text = (text or "").strip().lstrip("=").strip()
    if not text or "[" in text or "," in text or ";" in text:
        return None
    sheet, bang, ref = text.rpartition("!")
    if bang:
        sheet = sheet.strip()
        if sheet.startswith("'") and sheet.endswith("'") and len(sheet) > 1:
            sheet = sheet[1:-1].replace("''", "'")
    else:
        sheet = home
    match = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?", ref.strip().upper())
    if not match or not sheet:
        return None
    r1, c1 = int(match[2]), column_number(match[1])
    r2, c2 = (int(match[4]), column_number(match[3])) if match[3] else (r1, c1)
    if min(r1, r2) < 1:
        return None
    return sheet, (min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))


def _shown_ref(sheet: str, area: Area, home: str) -> str:
    """B2:B13 auf dem eigenen Blatt, sonst Daten!B2:B13."""
    text = area_text(area)
    return text if sheet.casefold() == home.casefold() else f"{sheet}!{text}"


def _text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        return _num(value)
    if isinstance(value, Err):
        return value.code
    if isinstance(value, bool):
        return T("WAHR" if value else "FALSCH", "TRUE" if value else "FALSE")
    return str(value)


def _q(text: str) -> str:
    """„Nord“, auf Englisch “Nord”."""
    return T(f"„{text}“", f"“{text}”")


def _quoted(items: list[str]) -> str:
    shown = [_q(i) for i in items[:MAX_LISTED]]
    if len(items) > MAX_LISTED:
        return ", ".join(shown) + T(f" und {len(items) - MAX_LISTED} weitere", f" and {len(items) - MAX_LISTED} more")
    return shown[0] if len(shown) == 1 else ", ".join(shown[:-1]) + T(" und ", " and ") + shown[-1]


def _listed(parts: list[str]) -> str:
    return parts[0] if len(parts) == 1 else f"{', '.join(parts[:-1])}{T(' und ', ' and ')}{parts[-1]}"


def _blank() -> str:
    """So zeigt eine Pivot-Tabelle leere Zellen der Daten."""
    return T("(Leer)", "(blank)")


def _is_total_label(value) -> bool:
    return isinstance(value, str) and bool(set(_words(value)) & _TOTAL)


def _number(value) -> bool:
    return isinstance(value, float) and not isinstance(value, bool)


# --- Pivot-Tabellen ---


def _key(value):
    """Wie eine Pivot-Tabelle Einträge zusammenfasst: Text ohne Unterschied zwischen groß und klein, Zahlen und Daten
    als Zahl, leere Zellen als ein Eintrag „(Leer)“."""
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return ("b", value)
    if isinstance(value, float):
        return ("n", round(value, 9))
    if isinstance(value, Err):
        return ("e", value.code)
    text = str(value)
    if text.strip().casefold() in _BLANK:
        return None
    return ("t", text.casefold())


def _item_key(text: str):
    """Ein Eintrag, wie die Datei oder Excel ihn nennt („Nord“, „12“, „01.03.2025“, „(leer)“)."""
    cell = Cell(text)
    return _key(calc_value(value_of(cell), cell))


def _aggregate(function: str, values: list):
    numbers = [v for v in values if _number(v)]
    filled = [v for v in values if v is not None and v != ""]
    if any(isinstance(v, Err) for v in values):
        return _ANY
    if function == "count":
        return float(len(filled)) if filled else _EMPTY
    if function == "countNums":
        return float(len(numbers)) if numbers else _EMPTY
    if not numbers:
        return _EMPTY
    if function == "sum":
        return math.fsum(numbers)
    if function == "average":
        return math.fsum(numbers) / len(numbers)
    if function == "max":
        return max(numbers)
    if function == "min":
        return min(numbers)
    if function == "product":
        return math.prod(numbers)
    n = len(numbers)
    mean = math.fsum(numbers) / n
    squares = math.fsum((x - mean) ** 2 for x in numbers)
    if function in ("stdDev", "var"):
        if n < 2:
            return _EMPTY
        return math.sqrt(squares / (n - 1)) if function == "stdDev" else squares / (n - 1)
    if function in ("stdDevp", "varp"):
        return math.sqrt(squares / n) if function == "stdDevp" else squares / n
    return _ANY


def _same(shown, expected) -> bool:
    if expected is _ANY:
        return True
    if expected is None or expected is _EMPTY:
        # Keine Zeile dazu: Excel lässt das Feld leer (oder zeigt, was für leere Felder eingestellt ist).
        return shown is None or (isinstance(shown, str) and not shown.strip()) or shown == 0.0 or isinstance(shown, str)
    if shown is None:
        return expected == 0.0
    if not _number(shown):
        return isinstance(shown, str) and not isinstance(shown, Err)
    return abs(shown - expected) <= 1e-8 * max(1.0, abs(shown), abs(expected))


@dataclass
class _Source:
    sheet: str
    area: Area
    # Spalte je Feldname (klein geschrieben), die Datenzeilen und ob die Quelle eine Excel-Tabelle ist.
    columns: dict[str, int]
    rows: list[int]
    table: str


def _source(pivot: Pivot, home: str, book: _Book) -> _Source | None:
    table_name, sheet, area, totals = pivot.source_table, pivot.source_sheet or home, pivot.source_area, 0
    if table_name and table_name.casefold() in book.tables:
        sheet, table = book.tables[table_name.casefold()]
        area, totals = table.area, table.totals
    elif table_name and area is None:
        return None
    if area is None or not book.has(sheet):
        return None
    sheet = book.name(sheet)
    top, left, bottom, right = area
    columns = {}
    for c in range(left, right + 1):
        head = book.value(sheet, (top, c))
        if head is not None and head != "":
            columns.setdefault(_text(head).strip().casefold(), c)
    rows = list(range(top + 1, bottom - totals + 1))[:MAX_ROWS]
    return _Source(sheet, area, columns, rows, table_name)


@dataclass
class _Line:
    """Eine Zeile oder Spalte der Pivot-Tabelle: ihr Eintrag (oder das Gesamtergebnis) und die Beschriftung."""

    pos: int
    key: object
    label: str
    total: bool = False


def _without_values_field(pivot: Pivot, source: _Source) -> Pivot:
    """Die Pivot-Tabelle ohne „Σ Werte“ bei Zeilen und Spalten. Stehen mehrere Wertfelder untereinander (Werte bei den
    Zeilen), rechnet der Prüfer sie nicht nach."""

    def pseudo(name: str) -> bool:
        return name.strip().lstrip("Σ∑ ").strip().casefold() in _VALUES_FIELD and _field_column(source, name) is None

    if not any(pseudo(name) for name in pivot.rows + pivot.columns):
        return pivot
    stacked = any(pseudo(name) for name in pivot.rows) and len(pivot.value_fields) > 1
    return replace(pivot, rows=tuple(n for n in pivot.rows if not pseudo(n)),
                   columns=tuple(n for n in pivot.columns if not pseudo(n)), plain=pivot.plain and not stacked)


def _pivot_findings(pivot: Pivot, home: str, book: _Book, live: bool) -> list[Finding]:
    source = _source(pivot, home, book)
    if source is None:
        return []
    pivot = _without_values_field(pivot, source)
    found = []
    stale = _stale(pivot, home, book, source)
    if stale:
        found.append(_with_action(stale, live, {"kind": "refresh", "pivot": pivot.name}))
    short = _short_source(pivot, home, book, source)
    if short:
        # Einsetzen baut die Pivot-Tabelle mit derselben Anordnung über den ganzen Bereich neu auf.
        area = short.__dict__.pop("_area")
        table, from_table = short.__dict__.pop("_table")
        # Ausgeblendete Einträge in den Zeilen und Spalten kann der Neuaufbau nicht mitnehmen: dann nur der Tipp.
        hidden = {name.casefold() for name, _ in pivot.skipped} - {f.partition(" = ")[0].strip().casefold()
                                                                   for f in pivot.filters}
        if from_table:
            # Die Tabelle erweitern und die Pivot-Tabelle aktualisieren: Sie bleibt, wie sie ist.
            found.append(_with_action(short, live, {
                "kind": "table", "pivot": pivot.name, "table": table, "sheet": source.sheet,
                "source": area_text(area), "was": area_text(source.area)}))
        elif hidden or not pivot.value_fields or not pivot.plain or any(v.shown_as for v in pivot.value_fields):
            found.append(short)
        else:
            # Liegt über der festen Quelle schon eine Excel-Tabelle, wird sie erweitert und zur Quelle.
            found.append(_with_action(short, live, {
                "kind": "source", "pivot": pivot.name, "sheet": source.sheet, "source": area_text(area),
                "was": area_text(source.area), "table": table}))
    said = _said_filter(pivot, home, book, source)
    if said:
        field, item, visible = said.__dict__.pop("_filter")
        # Einsetzen nur mit einem Eintrag, den es in den Daten gibt (so geschrieben, wie Excel ihn nennt).
        if item:
            found.append(_with_action(said, live, {"kind": "filter", "pivot": pivot.name, "field": field,
                                                   "value": item, "was": visible}))
        else:
            found.append(said)
    for counted in _counts(pivot, home, book, source):
        found.append(_with_action(counted[0], live, {"kind": "sum", "pivot": pivot.name, "field": counted[1]}))
    return found


_SPEC = re.compile(r"\bsoll\b|filter|\bnur\b|\bonly\b|should|gewünscht|vorgabe", re.IGNORECASE)


def _filter_items(pivot: Pivot, book: _Book, source: _Source, field: str) -> tuple[dict, list[str]] | None:
    """Die Einträge eines Felds in den Daten ({Schlüssel: wie Excel ihn nennt}) und die davon sichtbaren."""
    col = _field_column(source, field)
    if col is None:
        return None
    items = {}
    for r in source.rows:
        key = _key(book.value(source.sheet, (r, col)))
        if key is not None:
            items.setdefault(key, book.shown(source.sheet, (r, col)).strip())
    off = {_item_key(i) for name, hidden in pivot.skipped if name.strip().casefold() == field.casefold() for i in hidden}
    return items, [shown for key, shown in items.items() if key not in off]


def _said_filter(pivot: Pivot, home: str, book: _Book, source: _Source | None = None) -> Finding | None:
    """Über der Pivot-Tabelle steht, was sie zeigen soll („Filter Status = Gebucht“), ihr Berichtsfilter steht aber
    auf (Alle) oder einem anderen Eintrag. Welche Einträge er zeigt, kommt aus den Daten und den ausgeblendeten
    Einträgen; aus Excel kommt nur der Feldname, aus der Datei auch „Status = Offen“."""
    if pivot.area is None or not pivot.filters:
        return None
    top = pivot.area[0]
    for f in pivot.filters:
        field, _, chosen = f.partition(" = ")
        chosen = chosen.strip("„“\"")
        known = _filter_items(pivot, book, source, field.strip()) if source is not None else None
        visible = [chosen] if chosen else []
        if known is not None:
            items, shown = known
            if len(shown) < len(items):
                visible = shown
            elif not chosen:
                visible = []
        name = re.escape(field.strip())
        wanted = re.compile(rf"(?<![\w]){name}\s*(?:=|:)\s*[„\"']?([^,;.)„“\"']+?)[“\"']?\s*(?:[,;.)]|$)", re.IGNORECASE)
        for r in range(max(1, top - 4), top):
            for c in range(1, (pivot.area[3] if pivot.area else 1) + 3):
                text = book.shown(home, (r, c))
                if not text or not _SPEC.search(text):
                    continue
                m = wanted.search(text)
                if m is None:
                    continue
                value = m.group(1).strip()
                if not value or value.casefold() in ("alle", "all", "(alle)", "(all)"):
                    continue
                if [v.casefold() for v in visible] == [value.casefold()]:
                    continue
                item = ""
                if known is not None:
                    item = known[0].get(_item_key(value), "")
                    if not item:
                        continue  # Den Eintrag gibt es in den Daten nicht: vielleicht anders gemeint.
                if not visible:
                    now, counts = T("(Alle)", "(All)"), T("alle Einträge", "all entries")
                else:
                    now = _quoted(visible) if len(visible) > 1 else _q(visible[0])
                    counts = T(f"nur {now}", f"only {now}") if len(visible) == 1 else now
                found = Finding(
                    f"pivotfilter:{address((top, pivot.area[1]))}",
                    T(f"In {address((r, c))} steht, was die Pivot-Tabelle „{pivot.name}“ zeigen soll: {field.strip()} = "
                      f"„{value}“. Ihr Berichtsfilter {field.strip()} steht aber auf {now}. So zählt sie {counts} mit.",
                      f"{address((r, c))} says what the pivot table “{pivot.name}” should show: {field.strip()} = "
                      f"“{value}”. But its report filter {field.strip()} is set to {now}. So it counts {counts}."),
                    "fehler", T(f"Stell den Berichtsfilter {field.strip()} auf „{item or value}“.",
                                f"Set the report filter {field.strip()} to “{item or value}”."))
                found._filter = (field.strip(), item, visible)
                return found
    return None


def _with_action(finding: Finding, live: bool, action: dict) -> Finding:
    if live:
        finding.action = action
    return finding


def _field_column(source: _Source, name: str) -> int | None:
    return source.columns.get(name.strip().casefold())


def _rows_of(pivot: Pivot, book: _Book, source: _Source, rows: list[int] | None = None):
    """Je Datenzeile die Einträge der Zeilen- und Spaltenfelder, ohne die ausgeblendeten. None, wenn ein Feld nicht
    zu einer Spalte der Daten passt."""
    fields = list(pivot.rows) + list(pivot.columns)
    columns = [_field_column(source, f) for f in fields]
    skipped = []
    for name, items in pivot.skipped:
        col = _field_column(source, name)
        if col is None:
            return None
        skipped.append((col, {_item_key(i) for i in items}))
    if any(c is None for c in columns):
        return None
    out = []
    for r in rows if rows is not None else source.rows:
        keys = {}
        for col in {c for c in columns} | {c for c, _ in skipped}:
            keys[col] = _key(book.value(source.sheet, (r, col)))
        if any(keys[col] in items for col, items in skipped):
            continue
        out.append((r, tuple(keys[c] for c in columns)))
    # Ausgeblendete Einträge, die es in den Daten nicht gibt (anders geschrieben): lieber nicht nachrechnen.
    for col, items in skipped:
        present = {_key(book.value(source.sheet, (r, col))) for r in source.rows}
        if items - present:
            return None
    return out


def _place(line: _Line, col_label: str, col_total: bool, crossed: bool, caption: str) -> str:
    """Wie eine Zahl der Pivot-Tabelle heißt: „Nord“, „Nord“ / „Maus“, das Gesamtergebnis für „Maus“ …"""
    if line.total:
        return T(f"das Gesamtergebnis für „{col_label}“", f"the grand total for “{col_label}”") if crossed and not col_total \
            else T("das Gesamtergebnis", "the grand total")
    if crossed:
        return T(f"„{line.label}“ im Gesamtergebnis", f"“{line.label}” in the grand total") if col_total \
            else f"{_q(line.label)} / {_q(col_label)}"
    return _q(line.label) + (f" ({caption})" if caption else "")


def _stale(pivot: Pivot, home: str, book: _Book, source: _Source) -> Finding | None:
    """Zeigt die Pivot-Tabelle, was ihre Daten jetzt ergeben? Nur für einfache Pivot-Tabellen: ein Zeilenfeld, kein
    oder ein Spaltenfeld, bei einem Spaltenfeld ein Wertfeld."""
    if not pivot.plain or pivot.body is None or len(pivot.rows) != 1 or len(pivot.columns) > 1 or not pivot.value_fields:
        return None
    if pivot.columns and len(pivot.value_fields) != 1:
        return None
    top, left, bottom, right = pivot.body
    label_col, head_row = left - 1, top - 1
    if pivot.area is None or label_col < pivot.area[1] or head_row < pivot.area[0]:
        return None
    rows_of = _rows_of(pivot, book, source)
    if rows_of is None:
        return None
    all_rows = {_key(book.value(source.sheet, (r, _field_column(source, pivot.rows[0])))) for r in source.rows}

    # Zeilen der Pivot-Tabelle: Einträge, zuletzt das Gesamtergebnis.
    lines = []
    for r in range(top, bottom + 1):
        value = book.value(home, (r, label_col))
        total = pivot.grand[0] and r == bottom and bottom > top
        lines.append(_Line(r, None if total else _key(value), book.shown(home, (r, label_col)), total))
    items = [line for line in lines if not line.total]
    unknown = [line for line in items if line.key not in all_rows]
    if not items or len(unknown) * 2 > len(items) or len({line.key for line in items}) != len(items):
        return None

    # Spalten: die Wertfelder nebeneinander, oder die Einträge des Spaltenfelds und das Gesamtergebnis.
    columns: list[tuple[int, object, str, PivotValue, bool]] = []
    if not pivot.columns:
        if right - left + 1 != len(pivot.value_fields):
            return None
        for c, field in zip(range(left, right + 1), pivot.value_fields):
            head = book.value(home, (head_row, c))
            if _text(head).strip().casefold() != field.caption.strip().casefold():
                return None
            columns.append((c, None, "", field, False))
        col_unknown = []
    else:
        field = pivot.value_fields[0]
        all_cols = {_key(book.value(source.sheet, (r, _field_column(source, pivot.columns[0])))) for r in source.rows}
        for c in range(left, right + 1):
            head = book.value(home, (head_row, c))
            total = pivot.grand[1] and c == right and right > left
            columns.append((c, None if total else _key(head), book.shown(home, (head_row, c)), field, total))
        col_unknown = [col for col in columns if not col[4] and col[1] not in all_cols]
        if len(col_unknown) * 2 > len([col for col in columns if not col[4]]):
            return None
    usable = [col for col in columns if col[3].function and not col[3].shown_as and _field_column(source, col[3].field)]
    if not usable:
        return None

    shown_rows = {line.key for line in items}
    shown_cols = {col[1] for col in columns if not col[4]}
    # Blendet ein Filter Einträge aus (oder weiß man es nicht), zählen in den Gesamtergebnissen nur die gezeigten.
    sure = pivot.filtered is False

    def wanted(row_key, col_key, row_total, col_total) -> list[int]:
        out = []
        for r, keys in rows_of:
            if row_total:
                if not sure and keys[0] not in shown_rows:
                    continue
            elif keys[0] != row_key:
                continue
            if pivot.columns:
                if col_total:
                    if not sure and keys[1] not in shown_cols:
                        continue
                elif keys[1] != col_key:
                    continue
            out.append(r)
        return out

    wrong = []
    for line in lines:
        for c, col_key, col_label, field, col_total in usable:
            rows = wanted(line.key, col_key, line.total, col_total)
            data_col = _field_column(source, field.field)
            expected = _aggregate(field.function, [book.value(source.sheet, (r, data_col)) for r in rows]) if rows else None
            if line.total or col_total:
                if not rows:
                    expected = _ANY
            shown = book.value(home, (line.pos, c))
            if not _same(shown, expected):
                wrong.append(((line.pos, c), _place(line, col_label, col_total, bool(pivot.columns),
                                                    field.caption if len(usable) > 1 else ""), shown, expected))

    missing_rows = sorted({k[0] for _, k in rows_of} - shown_rows, key=str) if sure else []
    missing_cols = sorted({k[1] for _, k in rows_of} - shown_cols, key=str) if sure and pivot.columns else []
    extra = [line.label for line in unknown] + [col[2] for col in col_unknown]
    if not wrong and not missing_rows and not missing_cols and not extra:
        return None

    said = []
    for pos, name, shown, expected in wrong[:MAX_LISTED]:
        now = T("nichts mehr (keine Zeile der Daten gehört dazu)", "nothing any more (no row of the data belongs to it)") \
            if expected is None else \
            T("nichts mehr (dazu steht in den Daten keine Zahl mehr)", "nothing any more (the data no longer has a number "
              "for it)") if expected is _EMPTY else _text(expected)
        said.append(T(f"{address(pos)} {name} zeigt {_text(shown) or 'nichts'}, die Daten ergeben jetzt {now}",
                      f"{address(pos)} ({name}) shows {_text(shown) or 'nothing'} while the data now gives {now}"))
    if len(wrong) > MAX_LISTED:
        said.append(T(f"dazu {len(wrong) - MAX_LISTED} weitere Zahlen", f"plus {len(wrong) - MAX_LISTED} more numbers"))
    labels = {key: book.shown(source.sheet, (r, _field_column(source, pivot.rows[0]))) for r, (key, *_) in rows_of}
    if missing_rows:
        names = [labels.get(k) or _blank() for k in missing_rows]
        counts = [sum(1 for _, keys in rows_of if keys[0] == k) for k in missing_rows]
        n = counts[0] if len(names) == 1 else sum(counts)
        said.append(T(f"{_quoted(names)} {'fehlt' if len(names) == 1 else 'fehlen'} ganz, obwohl "
                      f"{n} Zeilen der Daten dazugehören",
                      f"{_quoted(names)} {'is' if len(names) == 1 else 'are'} missing entirely, although "
                      f"{n} {'row' if n == 1 else 'rows'} of the data belong there"))
    if missing_cols:
        col_labels = {keys[1]: book.shown(source.sheet, (r, _field_column(source, pivot.columns[0]))) for r, keys in rows_of}
        names = [col_labels.get(k) or _blank() for k in missing_cols]
        said.append(T(f"die Spalte {_quoted(names)} fehlt", f"the column {_quoted(names)} is missing") if len(names) == 1
                    else T(f"die Spalten {_quoted(names)} fehlen", f"the columns {_quoted(names)} are missing"))
    if extra:
        said.append(T(f"{_quoted(extra)} {'kommt' if len(extra) == 1 else 'kommen'} in den Daten nicht mehr vor",
                      f"{_quoted(extra)} no longer {'appears' if len(extra) == 1 else 'appear'} in the data"))
    where = _shown_ref(source.sheet, source.area, home) if not source.table else \
        T(f"der Tabelle „{source.table}“", f"the table “{source.table}”")
    text = T(f"Die Pivot-Tabelle „{pivot.name}“ zeigt nicht mehr, was ihre Daten ({where}) ergeben: "
             f"{'; '.join(said)}. Sie wurde nach einer Änderung der Daten nicht aktualisiert.",
             f"The pivot table “{pivot.name}” no longer shows what its data ({where}) gives: "
             f"{'; '.join(said)}. It was not refreshed after the data changed.")
    first = wrong[0][0] if wrong else (pivot.body[0], pivot.body[1])
    tip = T("Rechtsklick in die Pivot-Tabelle → Aktualisieren (oder hier „Aktualisieren“). Danach rechnet sie wieder mit "
            "den Daten, wie sie jetzt sind.",
            "Right-click in the pivot table → Refresh (or “Refresh” here). Then it calculates with the data as it "
            "is now again.")
    return Finding(f"pivot:{address(first)}", text, "fehler", tip)


def _row_kind(book: _Book, sheet: str, r: int, columns: range) -> str:
    """Eine Zeile unter der Quelle: Daten, Summenzeile oder leer."""
    values = [book.value(sheet, (r, c)) for c in columns]
    if all(v is None or v == "" for v in values):
        return "leer"
    if any(_is_total_label(v) for v in values):
        return "summe"
    for c in columns:
        cell = book.cell(sheet, (r, c))
        if cell is not None and cell.formula and re.match(r"\s*(SUM|SUBTOTAL|AGGREGATE)\(", cell.formula.upper()) \
                and _sums_column_above(cell.formula, r, c):
            return "summe"
    return "daten"


_AREA = re.compile(r"(?<![A-Za-z0-9_.!$])\$?([A-Z]{1,3})\$?(\d{1,7}):\$?([A-Z]{1,3})\$?(\d{1,7})(?![A-Za-z0-9_(])")


def _sums_column_above(formula: str, r: int, c: int) -> bool:
    """Summiert die Formel ihre eigene Spalte über sich (=SUMME(F2:F4001) in F4002)? Eine Laufsumme daneben
    (=SUMME($F$2:F4002) in I4002) ist eine Datenzeile wie die anderen."""
    for m in _AREA.finditer(re.sub(r'"[^"]*"', "", formula.upper())):
        if column_number(m[1]) == column_number(m[3]) == c and int(m[4]) < r:
            return True
    return False


def _short_source(pivot: Pivot, home: str, book: _Book, source: _Source) -> Finding | None:
    """Eine feste Quelle, unter der die Daten weitergehen: Diese Zeilen kommen auch beim Aktualisieren nicht dazu."""
    if not source.rows:
        return None
    top, left, bottom, right = source.area
    table = book.tables.get(source.table.casefold()) if source.table else None
    if source.table and (table is None or table[1].totals):
        return None
    if table is None:
        # Eine feste Quelle über genau einer Excel-Tabelle: Dann ist die Tabelle zu erweitern und als Quelle zu nehmen.
        table = next((t for t in book.tables.values() if t[0].casefold() == source.sheet.casefold()
                      and t[1].area == source.area and not t[1].totals), None)
    columns = range(left, right + 1)
    width = right - left + 1
    # Wie die Spalten aussehen: Zahlen, wo die Daten Zahlen haben.
    numeric = set()
    for c in columns:
        values = [book.value(source.sheet, (r, c)) for r in source.rows]
        filled = [v for v in values if v is not None and v != ""]
        if filled and sum(_number(v) for v in filled) >= 0.8 * len(filled):
            numeric.add(c)
    extra = []
    r = bottom + 1
    while len(extra) < MAX_ROWS:
        if _row_kind(book, source.sheet, r, columns) != "daten":
            break
        values = {c: book.value(source.sheet, (r, c)) for c in columns}
        filled = [c for c, v in values.items() if v is not None and v != ""]
        if len(filled) * 2 < width or any(c in numeric and not _number(values[c]) for c in filled):
            break
        extra.append(r)
        r += 1
    if not extra:
        return None
    rows_all = _rows_of(pivot, book, source, source.rows + extra)
    effect = ""
    if rows_all is not None and pivot.rows and pivot.value_fields and pivot.value_fields[0].function:
        field = pivot.value_fields[0]
        data_col = _field_column(source, field.field)
        if data_col is not None:
            before, after, labels = {}, {}, {}
            for r, keys in rows_all:
                labels.setdefault(keys[0], book.shown(source.sheet, (r, _field_column(source, pivot.rows[0]))))
                if r <= bottom:
                    before.setdefault(keys[0], []).append(r)
                after.setdefault(keys[0], []).append(r)
            changed, added = [], []
            for key, rows in after.items():
                old = _aggregate(field.function, [book.value(source.sheet, (r, data_col)) for r in before.get(key, [])])
                new = _aggregate(field.function, [book.value(source.sheet, (r, data_col)) for r in rows])
                label = labels[key] or _blank()
                if _number(new) and not before.get(key):
                    added.append(f"{_q(label)} ({_text(new)})")
                elif _number(new) and _number(old) and not _same(old, new):
                    changed.append(T(f"„{label}“ {_text(new)} statt {_text(old)}",
                                     f"“{label}” would be {_text(new)} instead of {_text(old)}"))
            more = max(0, len(changed) + len(added) - MAX_LISTED)
            changed = changed[:MAX_LISTED]
            added = added[:MAX_LISTED - len(changed)]
            if changed or added:
                what = T(f" bei „{field.caption}“", f" for “{field.caption}”") if len(pivot.value_fields) > 1 else ""
                said = []
                if changed:
                    listed = ', '.join(changed) if added or more else _listed(changed)
                    said.append(T(f"{'wäre' if len(changed) == 1 else 'wären'}{what} {listed}", listed))
                if added:
                    said.append(T(f"{'es ' if changed else ''}{'käme' if len(added) == 1 else 'kämen'}{what if not changed else ''} "
                                  f"{_listed(added)} dazu",
                                  f"{_listed(added)} would be added"))
                if more:
                    said.append(T(f"{more} weitere {'wäre' if more == 1 else 'wären'} anders",
                                  f"{more} more would be different"))
                joined = said[0] if len(said) == 1 else ', '.join(said[:-1]) + T(' und ', ' and ') + said[-1]
                effect = T(f" Mit ihnen {joined}.", f" With them{what}, {joined}.")
    if len(extra) == 1:
        missing = T(f"Zeile {extra[0]} fehlt in der Pivot-Tabelle, auch nach dem Aktualisieren.",
                    f"row {extra[0]} is missing from the pivot table, even after refreshing.")
    else:
        missing = T(f"die {len(extra)} Zeilen {extra[0]} bis {extra[-1]} fehlen in der Pivot-Tabelle, auch nach dem "
                    "Aktualisieren.",
                    f"the {len(extra)} rows {extra[0]} to {extra[-1]} are missing from the pivot table, even after "
                    "refreshing.")
    fixed = _shown_ref(source.sheet, (top, left, extra[-1], right), home)
    if source.table:
        name = table[1].name
        text = T(f"Die Pivot-Tabelle „{pivot.name}“ nimmt ihre Daten aus der Tabelle {name} "
                 f"({_shown_ref(source.sheet, source.area, home)}). Darunter gehen die Daten aber weiter, außerhalb der "
                 f"Tabelle: {missing}{effect}",
                 f"The pivot table “{pivot.name}” takes its data from the table {name} "
                 f"({_shown_ref(source.sheet, source.area, home)}). But the data continues below it, outside the "
                 f"table: {missing}{effect}")
        tip = T(f"Tabelle {name} auf {fixed} erweitern (in die Tabelle klicken, Tabellenentwurf → Tabellengröße "
                "ändern), dann die Pivot-Tabelle aktualisieren.",
                f"Extend the table {name} to {fixed} (click in the table, Table Design → Resize Table), then refresh "
                "the pivot table.")
    else:
        text = T(f"Die Pivot-Tabelle „{pivot.name}“ nimmt ihre Daten fest aus "
                 f"{_shown_ref(source.sheet, source.area, home)}. Darunter gehen die Daten aber weiter: {missing}{effect}",
                 f"The pivot table “{pivot.name}” takes its data from the fixed range "
                 f"{_shown_ref(source.sheet, source.area, home)}. But the data continues below it: {missing}{effect}")
        if table is not None:
            tip = T(f"Tabelle {table[1].name} auf {fixed} erweitern (Tabellenentwurf → Tabellengröße ändern), dann "
                    f"PivotTable-Analyse → Datenquelle ändern → {table[1].name}. Dann kommen neue Zeilen von selbst dazu.",
                    f"Extend the table {table[1].name} to {fixed} (Table Design → Resize Table), then "
                    f"PivotTable Analyze → Change Data Source → {table[1].name}. New rows are then added automatically.")
        else:
            tip = T(f"PivotTable-Analyse → Datenquelle ändern → {fixed}. Damit neue Zeilen von selbst dazukommen: in die "
                    "Daten klicken, Befehl-T (Als Tabelle formatieren) und die Tabelle als Datenquelle nehmen.",
                    f"PivotTable Analyze → Change Data Source → {fixed}. To have new rows added automatically: click in "
                    "the data, press Command-T (Format as Table) and use the table as the data source.")
    cell = (extra[0], left) if source.sheet.casefold() == home.casefold() else None
    finding = Finding(f"pivotsource:{address(cell or ((pivot.area or pivot.body)[0], (pivot.area or pivot.body)[1]))}",
                      text, "fehler", tip)
    if cell is None:
        finding.cell = ""
    finding._area = (top, left, extra[-1], right)
    finding._table = (table[1].name if table is not None else "", bool(source.table))
    return finding


def _counts(pivot: Pivot, home: str, book: _Book, source: _Source) -> list[tuple[Finding, str]]:
    """„Anzahl von Umsatz“ unter einer Spalte mit Zahlen: Excel zählt, weil eine Zelle leer ist oder Text hat."""
    out = []
    for i, field in enumerate(pivot.value_fields):
        match = _COUNT_CAPTION.match(field.caption.strip())
        if field.function != "count" or field.shown_as or not match or \
                match["field"].strip().casefold() != field.field.strip().casefold():
            continue
        col = _field_column(source, field.field)
        if col is None or set(_words(field.field)) & _COUNTED:
            continue
        values = [(r, book.value(source.sheet, (r, col))) for r in source.rows]
        numbers = [v for _, v in values if _number(v)]
        odd = [(r, v) for r, v in values if not _number(v)]
        if len(numbers) < 3 or not odd or len(numbers) < 0.8 * len(values):
            continue
        # Lauter verschiedene ganze Zahlen sind eher Nummern als Beträge.
        if all(v == int(v) for v in numbers) and len(set(numbers)) == len(numbers):
            continue
        r, v = odd[0]
        where = _shown_ref(source.sheet, (r, col, r, col), home)
        why = T(f"{where} ist leer", f"{where} is empty") if v is None or v == "" else \
            T(f"in {where} steht „{book.shown(source.sheet, (r, col))}“", f"{where} contains “{book.shown(source.sheet, (r, col))}”")
        total = math.fsum(numbers)
        filled = len([1 for _, v in values if v not in (None, '')])
        text = T(f"„{field.caption}“ in der Pivot-Tabelle „{pivot.name}“ zählt, wie viele Zeilen etwas unter „{field.field}“ "
                 f"haben ({filled}), statt die Werte zu summieren ({_text(total)}). "
                 f"Excel nimmt Anzahl, sobald in der Spalte beim Anlegen eine leere Zelle oder Text steht ({why}).",
                 f"“{field.caption}” in the pivot table “{pivot.name}” counts how many rows have something under "
                 f"“{field.field}” ({filled}) instead of adding up the values ({_text(total)}). "
                 f"Excel uses Count as soon as the column contains an empty cell or text when the pivot table is created "
                 f"({why}).")
        tip = T("Ist die Summe gemeint: Rechtsklick auf eine Zahl der Pivot-Tabelle → Werte zusammenfassen nach → Summe "
                "(oder hier „Summe nehmen“).",
                "If the sum is meant: right-click a number in the pivot table → Summarize Values By → Sum "
                "(or “Use sum” here).")
        body = pivot.body
        cell = (body[0], body[1] + (i if not pivot.columns else 0)) if body else (pivot.area[0], pivot.area[1])
        out.append((Finding(f"pivotcount:{address(cell)}", text, "pruefen", tip), field.caption))
    return out


# --- Diagramme ---


@dataclass
class _Series:
    index: int
    name: str
    sheet: str
    # Werte und Beschriftungen als Linie: fest die Spalte (senkrecht) oder die Zeile (waagrecht), von first bis last.
    vertical: bool
    fixed: int
    first: int
    last: int
    cat_fixed: int | None = None
    cat_first: int | None = None
    cat_last: int | None = None
    values_text: str = ""
    categories_text: str = ""
    name_ref: str = ""


def _pos(series: _Series, i: int, fixed: int | None = None) -> Pos:
    line = series.fixed if fixed is None else fixed
    return (i, line) if series.vertical else (line, i)


def _line_area(vertical: bool, fixed: int, first: int, last: int) -> Area:
    return (first, fixed, last, fixed) if vertical else (fixed, first, fixed, last)


def _series_of(chart: Chart, home: str) -> list[_Series]:
    out = []
    for i, series in enumerate(chart.series):
        values = _ref(series.values, home)
        if values is None:
            continue
        sheet, (top, left, bottom, right) = values
        categories = _ref(series.categories, home) if series.categories else None
        if categories is not None and categories[0].casefold() != sheet.casefold():
            categories = None
        if left == right and bottom > top:
            vertical = True
        elif top == bottom and right > left:
            vertical = False
        elif categories is not None and categories[1][1] == categories[1][3] and categories[1][2] > categories[1][0]:
            vertical = True
        elif categories is not None and categories[1][0] == categories[1][2] and categories[1][3] > categories[1][1]:
            vertical = False
        else:
            continue
        item = _Series(i, series.name, sheet, vertical, left if vertical else top, top if vertical else left,
                       bottom if vertical else right, values_text=series.values, categories_text=series.categories,
                       name_ref=series.name_ref)
        if categories is not None:
            ct, cl, cb, cr = categories[1]
            if vertical and cl == cr:
                item.cat_fixed, item.cat_first, item.cat_last = cl, ct, cb
            elif not vertical and ct == cb:
                item.cat_fixed, item.cat_first, item.cat_last = ct, cl, cr
        out.append(item)
    return out


def _label_text(book: _Book, series: _Series, i: int) -> str:
    """Die Beschriftung so, wie Excel sie zeigt."""
    if series.cat_fixed is not None:
        return book.shown(series.sheet, _pos(series, i, series.cat_fixed))
    for line in range(series.fixed - 1, 0, -1):
        value = book.value(series.sheet, _pos(series, i, line))
        if isinstance(value, str) and value.strip():
            return book.shown(series.sheet, _pos(series, i, line))
    return ""


def _label(book: _Book, series: _Series, i: int):
    """Die Beschriftung einer Stelle: aus den Beschriftungen der Reihe, sonst der erste Text links (oder oben)."""
    if series.cat_fixed is not None:
        return book.value(series.sheet, _pos(series, i, series.cat_fixed))
    for line in range(series.fixed - 1, 0, -1):
        value = book.value(series.sheet, _pos(series, i, line))
        if isinstance(value, str) and value.strip():
            return value
    return None


def _is_total(book: _Book, series: _Series, i: int) -> bool:
    if _is_total_label(_label(book, series, i)):
        return True
    cell = book.cell(series.sheet, _pos(series, i))
    if cell is None or not cell.formula:
        return False
    try:
        node = formel.parse(cell.formula)
    except formel.Unknown:
        return False
    if node.kind != "call" or node.value not in _SUM_FUNCTIONS:
        return False
    _, areas, _ = formel.references(node)
    for r1, c1, r2, c2 in areas:
        a = _line_area(series.vertical, series.fixed, min(series.first, i - 1), i - 1)
        inside = r1 >= a[0] and c1 >= a[1] and r2 <= a[2] and c2 <= a[3]
        if inside and (r2 - r1) + (c2 - c1) >= 1:
            return True
    return False


def _is_data(book: _Book, series: _Series, i: int) -> bool:
    if i < 1:
        return False
    if not _number(book.value(series.sheet, _pos(series, i))):
        return False
    if series.cat_fixed is not None:
        label = book.value(series.sheet, _pos(series, i, series.cat_fixed))
        if label is None or label == "":
            return False
    return not _is_total(book, series, i)


def _run(book: _Book, series: _Series) -> tuple[int, int] | None:
    """Die Zeilen (oder Spalten) der Tabelle, zu denen die Reihe gehört: alle Zahlen am Stück um ihren Bereich herum,
    ohne Summenzeile."""
    seeds = [i for i in range(series.first, series.last + 1) if _is_data(book, series, i)]
    if len(seeds) < 2:
        return None
    start, end = min(seeds), max(seeds)
    while end - start < MAX_ROWS and _is_data(book, series, end + 1):
        end += 1
    while start > 1 and _is_data(book, series, start - 1):
        start -= 1
    return start, end


def _span_words(book: _Book, series: _Series, positions: list[int]) -> str:
    labels = [_label_text(book, series, i) for i in positions]
    if all(labels) and len(positions) <= 6:
        return _quoted(labels) if len(labels) > 1 else _q(labels[0])
    if len(positions) == 1:
        where = positions[0] if series.vertical else column_letters(positions[0])
        return T(f"{'Zeile' if series.vertical else 'Spalte'} {where}", f"{'row' if series.vertical else 'column'} {where}")
    first, last = positions[0], positions[-1]
    if not series.vertical:
        return T(f"die Spalten {column_letters(first)} bis {column_letters(last)}",
                 f"the columns {column_letters(first)} to {column_letters(last)}")
    return T(f"die Zeilen {first} bis {last}", f"the rows {first} to {last}")


def _chart_findings(chart: Chart, home: str, book: _Book, live: bool) -> list[Finding]:
    title = set(_words(f"{chart.title} {chart.name}"))
    subset = bool(title & _PERIOD) or any(_YEAR.match(w) for w in title)
    said, changes, level, first_cell = [], [], "pruefen", None
    named = []
    for series in _series_of(chart, home):
        if not book.has(series.sheet):
            continue
        series.sheet = book.name(series.sheet)
        run = _run(book, series)
        if run is None:
            continue
        start, end = run
        name = _q(series.name) if series.name else f"{series.index + 1}"
        declared = _shown_ref(series.sheet, _line_area(series.vertical, series.fixed, series.first, series.last), home)
        problems = []
        total = series.last == end + 1 and _is_total(book, series, series.last)
        shifted = series.cat_first is not None and series.cat_first != series.first and \
            series.cat_last - series.cat_first == series.last - series.first
        if total and not shifted:
            problems.append(T(f"nimmt mit {address(_pos(series, series.last))} auch die Summenzeile mit: Sie steht als "
                              "eigener Punkt im Diagramm und lässt die übrigen klein aussehen",
                              f"also includes the total row with {address(_pos(series, series.last))}: it appears as a "
                              "point of its own in the chart and makes the others look small"))
            level = "fehler"
        if not subset and series.last < end:
            missing = list(range(series.last + 1, end + 1))
            problems.append(T(f"zeigt nur {declared}, die Tabelle geht aber bis {address(_pos(series, end))}: "
                              f"{_span_words(book, series, missing)} {'fehlt' if len(missing) == 1 else 'fehlen'} im Diagramm",
                              f"shows only {declared}, but the table goes down to {address(_pos(series, end))}: "
                              f"{_span_words(book, series, missing)} {'is' if len(missing) == 1 else 'are'} missing from "
                              "the chart"))
            level = "fehler"
        if not subset and series.first > start and not (series.cat_first is not None and series.cat_first != series.first):
            missing = list(range(start, series.first))
            problems.append(T(f"fängt erst bei {address(_pos(series, series.first))} an, die Tabelle schon bei "
                              f"{address(_pos(series, start))}: {_span_words(book, series, missing)} "
                              f"{'fehlt' if len(missing) == 1 else 'fehlen'} im Diagramm",
                              f"starts only at {address(_pos(series, series.first))}, but the table starts at "
                              f"{address(_pos(series, start))}: {_span_words(book, series, missing)} "
                              f"{'is' if len(missing) == 1 else 'are'} missing from the chart"))
        if shifted:
            cats = _shown_ref(series.sheet, _line_area(series.vertical, series.cat_fixed, series.cat_first, series.cat_last), home)
            k = series.first - series.cat_first
            here = series.cat_first if series.cat_first >= start else start
            label, other = _label_text(book, series, here), here + k
            steps = T({1: "eine", 2: "zwei", 3: "drei"}.get(abs(k), str(abs(k))),
                      {1: "one", 2: "two", 3: "three"}.get(abs(k), str(abs(k))))
            text = T(f"ist gegen ihre Beschriftungen {cats} um {steps} {'Zeile' if series.vertical else 'Spalte'}"
                     f"{'n' if abs(k) > 1 else ''} verschoben: Bei „{label}“ steht der Wert aus {address(_pos(series, other))}",
                     f"is shifted by {steps} {'row' if series.vertical else 'column'}{'s' if abs(k) > 1 else ''} against "
                     f"its labels {cats}: at “{label}” is the value from {address(_pos(series, other))}") \
                + (f" ({_q(_label_text(book, series, other))})" if _label_text(book, series, other) else "")
            if total:
                text += T(f", bei „{_label_text(book, series, series.cat_last)}“ sogar die Summenzeile "
                          f"{address(_pos(series, series.last))}",
                          f", at “{_label_text(book, series, series.cat_last)}” even the total row "
                          f"{address(_pos(series, series.last))}")
            problems.append(text)
            level = "fehler"
        # Bleibt die Reihe am Anfang oder Ende so, wie sie ist (gewollter Ausschnitt), setzt die Korrektur nur, was
        # falsch ist.
        new_first = start if (not subset or shifted) else series.first
        new_last = end if (not subset or total or shifted) else min(series.last, end)
        if not problems:
            if series.name_ref:
                odd = _name_from_elsewhere(book, series, start, home)
                if odd:
                    named.append(odd)
            continue
        said.append(T(f"Die Reihe {name} {'; '.join(problems)}", f"The series {name} {'; '.join(problems)}"))
        new = _line_area(series.vertical, series.fixed, new_first, new_last)
        # was: wie Excel die Bereiche jetzt nennt (daran sieht das Seitenfenster, ob sich seither etwas geändert hat),
        # undo: dieselben Bereiche zum Zurücksetzen.
        change = {"index": series.index, "sheet": series.sheet, "values": area_text(new),
                  "was": {"values": series.values_text, "categories": series.categories_text},
                  "undo": {"values": area_text(_line_area(series.vertical, series.fixed, series.first, series.last))}}
        if series.cat_fixed is not None:
            change["categories"] = area_text(_line_area(series.vertical, series.cat_fixed, new_first, new_last))
            change["undo"]["categories"] = area_text(_line_area(series.vertical, series.cat_fixed, series.cat_first,
                                                                series.cat_last))
        changes.append(change)
        if first_cell is None:
            first_cell = (series.sheet, _pos(series, series.first))
    found = []
    if said:
        sheet, pos = first_cell
        titled = f" ({_q(chart.title)})" if chart.title and chart.title != chart.name else ""
        text = T(f"Im Diagramm „{chart.name}“{titled}: ", f"In the chart “{chart.name}”{titled}: ") + ". ".join(said) + "."
        targets = ", ".join(f"{_shown_ref(c['sheet'], formel.area(c['values']), home)}"
                            + (T(" mit den Beschriftungen ", " with the labels ")
                               + _shown_ref(c['sheet'], formel.area(c['categories']), home) if c.get("categories") else "")
                            for c in changes)
        tip = T(f"Rechtsklick ins Diagramm → Daten auswählen, dort {'die Reihe' if len(changes) == 1 else 'die Reihen'} auf "
                f"{targets} setzen (oder hier „Einsetzen“).",
                f"Right-click the chart → Select Data, and set the series to {targets} there (or “Apply” here).")
        finding = Finding(f"chart:{address(pos)}", text, level, tip)
        if sheet.casefold() != home.casefold():
            finding.cell = ""
        found.append(_with_action(finding, live, {"kind": "series", "chart": chart.name, "series": changes}))
    found += named
    return found


def _name_from_elsewhere(book: _Book, series: _Series, start: int, home: str) -> Finding | None:
    """Die Reihe heißt nach einer anderen Spalte, als ihre Werte stehen („Kosten“ über den Zahlen unter „Umsatz“)."""
    ref = _ref(series.name_ref, home)
    if ref is None or ref[0].casefold() != series.sheet.casefold() or ref[1][0] != ref[1][2] or ref[1][1] != ref[1][3]:
        return None
    pos = (ref[1][0], ref[1][1])
    head = _pos(series, start - 1)
    if start < 2 or pos == head:
        return None
    same_line = pos[0] == head[0] if series.vertical else pos[1] == head[1]
    name, header = book.value(series.sheet, pos), book.value(series.sheet, head)
    if not same_line or not isinstance(name, str) or not isinstance(header, str) or not header.strip() or \
            name.strip().casefold() == header.strip().casefold():
        return None
    values = _shown_ref(series.sheet, _line_area(series.vertical, series.fixed, series.first, series.last), home)
    text = T(f"Die Reihe „{name}“ im Diagramm hat ihren Namen aus {address(pos)}, zeigt aber die Werte unter "
             f"„{header}“ ({values}). In der Legende steht darum der falsche Name, oder es sind die falschen Werte.",
             f"The series “{name}” in the chart takes its name from {address(pos)} but shows the values under "
             f"“{header}” ({values}). So either the legend shows the wrong name, or these are the wrong values.")
    tip = T("Rechtsklick ins Diagramm → Daten auswählen → Reihe bearbeiten: Name und Werte aus derselben Spalte nehmen.",
            "Right-click the chart → Select Data → Edit Series: take the name and the values from the same column.")
    finding = Finding(f"chartname:{address(pos)}", text, "pruefen", tip)
    if series.sheet.casefold() != home.casefold():
        finding.cell = ""
    return finding


# --- alles zusammen ---


def tables_of(sheets: list[Sheet]) -> dict[str, tuple[str, ExcelTable]]:
    """Die Excel-Tabellen der Mappe nach Namen (Quelle einer Pivot-Tabelle)."""
    return {table.name: (sheet.name, table) for sheet in sheets for table in sheet.tables}


def checks(home: str, sheet: Sheet, cells_of: dict[str, dict[Pos, Cell]],
           tables: dict[str, tuple[str, ExcelTable]] | None = None, live: bool = False) -> list[Finding]:
    """Hinweise zu den Pivot-Tabellen und Diagrammen auf dem Blatt home. cells_of: alle Blätter der Mappe (die Daten
    können auf einem anderen stehen). live: aus dem Seitenfenster in Excel, dann mit Knöpfen."""
    if not sheet.pivots and not sheet.charts:
        return []
    book = _Book(cells_of, tables or {})
    found = []
    for pivot in sheet.pivots:
        try:
            found += _pivot_findings(pivot, home, book, live)
        except (ValueError, KeyError, IndexError, TypeError, ZeroDivisionError, OverflowError):
            continue
    for chart in sheet.charts:
        try:
            found += _chart_findings(chart, home, book, live)
        except (ValueError, KeyError, IndexError, TypeError, ZeroDivisionError, OverflowError):
            continue
    # Zwei Diagramme (oder Pivot-Tabellen) über dieselben Daten: Jeder Hinweis behält seine Zelle, der Schlüssel wird
    # eindeutig.
    seen: dict[str, int] = {}
    for finding in found:
        count = seen[finding.key] = seen.get(finding.key, 0) + 1
        if count > 1:
            if finding.cell is None:
                finding.cell = finding.key.split(":", 1)[1]
            finding.key += f"~{count}"
    return found


def areas(sheet: Sheet) -> list[Area]:
    """Wo Pivot-Tabellen stehen: Ihre Zahlen sind keine Eingaben und keine abgetippten Werte."""
    return [p.area for p in sheet.pivots if p.area is not None]


