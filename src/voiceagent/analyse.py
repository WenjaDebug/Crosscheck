"""Das Blatt als Rechenwerk prüfen: Abhängigkeiten, Muster, Nachrechnen, Sensitivität und Daten.

Alles hier ist deterministisch und erklärbar, ohne KI:

- Rechenkette: welche Zelle wovon abhängt, Kreisbezüge, Bezüge auf leere Zellen, Summen, die
  Zwischensummen noch einmal zählen.
- Formelmuster: In einer Spalte rechnen Formeln meist gleich (gleiche Formel in Z1S1-Schreibweise).
  Fällt eine aus dem Muster, ist das oft ein Fehler beim Kopieren (wie ExceLint und die
  Excel-Warnung „inkonsistente Formel“, aber über die ganze Spalte gezählt).
- Nachrechnen: Jede Formel, die der eigene Rechner kann (formel.py), wird mit den Werten aus Excel
  nachgerechnet. Weicht Excel ab, steht die Berechnung vielleicht auf manuell.
- Sensitivität: Für die Ergebnisse (Formeln, auf die nichts mehr aufbaut) wird jede Eingabe und jede
  feste Zahl in den Formeln um 1 % erhöht und die ganze Kette neu gerechnet. Die Elastizität sagt,
  um wie viel Prozent sich das Ergebnis dann ändert, also wovon es wirklich abhängt.
- Daten: robuste Ausreißer (Median und MAD), doppelte Namen, Anteile, die nicht 100 % ergeben, und
  Anteile in gemischter Schreibweise (0,33 neben 28).
"""

from __future__ import annotations

import logging
import math
import re
import statistics
import threading
import time
import traceback
from bisect import bisect_left, bisect_right
from collections import Counter
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path

from . import formel, woerter
from .sprache import T
from .formel import Err, Node
from .workbook import FUNCTIONS, Layout, _as_number, _left_label, pattern, readable
from .xlsx import Cell, _date_text, column_letters, column_number, cut_of, number_text, shift

Pos = tuple[int, int]

# Ab so vielen gleichen Formeln in einer Spalte gilt das als Muster, und so groß muss die Mehrheit sein.
PATTERN_MIN = 3
PATTERN_SHARE = 0.7
# So viele Ergebnisse und Eingaben je Ergebnis prüft die Sensitivität höchstens, und so lange.
MAX_OUTPUTS = 4
MAX_INPUTS = 150
# Bei mehr Ergebnissen sucht sie die größten zuerst grob und verfolgt nur so viele Ketten ganz.
MAX_CHAINED = 200
TIME_BUDGET = 1.5
# Diese festen Zahlen sind meist Umrechnungen, keine Annahmen.
PLAIN_NUMBERS = {0.0, 1.0, -1.0, 100.0, 1000.0, 2.0, 12.0, 24.0, 60.0}
# Robuster Ausreißer: so viele MAD vom Median weg, ab so vielen Werten.
OUTLIER_Z = 6.0
OUTLIER_MIN = 8
_PERCENT = ("%", "anteil", "prozent", "quote", "share", "percent", "pourcentage", "percentuale")
# Veränderungen in Prozent dürfen negativ und über 100 sein.
_CHANGES = ("veränderung", "änderung", "wachstum", "delta", "diff", "abweichung", "change", "growth", "zunahme", "abnahme",
            *woerter.CHANGES)
_SUMS = ("SUM", "SUBTOTAL", "AGGREGATE")
_ERROR_CODES = {"#NV": "#N/A", "#WERT!": "#VALUE!", "#BEZUG!": "#REF!", "#ZAHL!": "#NUM!", "#NAME?": "#NAME?"}


@dataclass
class Finding:
    key: str
    text: str
    level: str  # fehler, pruefen, info
    tip: str = ""
    # Die Korrektur zum Einsetzen (einsetzen.fix): neue Formel je Zelle und was dort vorher stand.
    fix: dict | None = None
    # Was das Seitenfenster in Excel sonst tun kann (pivotdiagramm.py): eine Pivot-Tabelle aktualisieren, die Summe
    # statt der Anzahl nehmen, die Bereiche eines Diagramms setzen.
    action: dict | None = None
    # Die Zelle zum Markieren, wenn sie nicht aus dem Schlüssel hervorgeht („“: keine auf diesem Blatt).
    cell: str | None = None
    # Zwei Korrekturen zur Wahl, wenn die Daten nicht sagen, welche stimmt (kategorien.py): [{"label": …, "fix": …}].
    choices: list | None = None


@dataclass
class Driver:
    """Eine Eingabe, von der ein Ergebnis abhängt: 1 % mehr dort → elasticity % mehr im Ergebnis."""

    output: str
    output_value: str
    input: str
    input_value: str
    elasticity: float


@dataclass
class Report:
    findings: list[Finding] = field(default_factory=list)
    drivers: list[Driver] = field(default_factory=list)
    stats: dict = field(default_factory=dict)
    # Formeln, die unter die gelesenen Zeilen eines großen Blatts reichen (Sheet.partial), und bis wohin die Blätter
    # gelesen sind (None: dieses Blatt), wenn eines größer ist.
    partial: set[Pos] = field(default_factory=set)
    cuts: dict[str | None, int] = field(default_factory=dict)
    # Prüfungen, die an diesem Blatt gescheitert sind (_Guard), mit ihrem Namen für das Fenster.
    failed: list[str] = field(default_factory=list)


_log = logging.getLogger("voiceagent.pruefer")


def failure(err: BaseException) -> str:
    """Wo eine Prüfung gescheitert ist, für den Fehlerbericht: Art des Fehlers und die letzten Stellen im Code, nie
    die Meldung selbst. Darin könnte ein Wert aus der Mappe stehen (KeyError('Dr. Schmidt'))."""
    frames = traceback.extract_tb(err.__traceback__)[-3:]
    return f"{type(err).__name__} in " + " < ".join(f"{Path(f.filename).name}:{f.lineno} {f.name}" for f in reversed(frames))


class _Guard:
    """Führt eine Prüfung aus. Scheitert sie, merkt der Bericht ihren Namen, der Fehlerbericht wo, und die übrigen
    Prüfungen laufen weiter. Ein Anhalten (Stopped) gehört nicht dazu: Das reicht er weiter."""

    def __init__(self, report: Report) -> None:
        self.report = report

    def __call__(self, what: str, run, default):
        try:
            return run()
        except Stopped:
            raise
        except Exception as err:  # noqa: BLE001 - eine ungewöhnliche Mappe darf nicht die ganze Prüfung kosten
            _log.error("Prüfung „%s“ gescheitert: %s", what, failure(err))
            shown = T(what, _CHECK_NAMES.get(what, what))
            if shown not in self.report.failed:
                self.report.failed.append(shown)
            return default


# Die Namen der Prüfungen für das Fenster auf Englisch (_Guard; auch die aus pruefer.py).
_CHECK_NAMES = {
    "Kreisbezüge": "circular references", "Formelmuster": "formula patterns", "Blöcke einer Spalte": "blocks in a column",
    "Verdeckte Zahlen": "hidden numbers", "Als 0 gezeigt": "shown as 0", "Zwischensummen": "subtotals",
    "Finanzplan": "financial plan", "Nachrechnen": "recalculation", "leere Zellen in Formeln": "empty cells in formulas",
    "SVERWEIS und Nachschlagen": "VLOOKUP and lookups", "schwer prüfbare Formeln": "formulas that are hard to check",
    "Formeln als Text": "formulas as text", "Daten": "data", "Auswertung der Daten": "data analysis",
    "Folgen beim Ursprung": "consequences at the source", "Summen und ihre Teile": "totals and their parts",
    "Statistik": "statistics", "Statistiktests": "statistical tests", "Zeiträume": "time periods",
    "eigene Regeln": "your own rules", "Listen": "lists", "Spalte zum Namen": "column for the name",
    "Doppelte Schlüssel": "duplicate keys", "Einheiten je Zeile": "units per row", "Verrutschte Sätze": "shifted rates",
    "Geschäftsjahr und Monatsende": "fiscal year and month end", "Modell": "model", "Finanzen": "finance",
    "Einsetzen": "applying fixes", "Sensitivitätstabellen": "sensitivity tables", "Fehlerwerte": "error values",
    "Sensitivität": "sensitivity", "Rechenkette": "calculation chain",
    "schnelle Prüfungen": "quick checks", "weitere Tabellen": "other tables", "verbundene Zellen": "merged cells",
    "Verknüpfungen": "links", "ausgeblendete Zeilen": "hidden rows", "Pivot-Tabellen und Diagramme": "pivot tables and charts",
}


def value_of(cell: Cell | None):
    """Der Wert einer Zelle, wie Excel ihn rechnet: Zahl, Text, Wahrheitswert, Fehler oder leer."""
    if cell is None or cell.value == "":
        return None
    if cell.text is not None:
        # Ziffern als Text oder „A “ mit Leerzeichen: Excel rechnet und sucht damit als Text.
        return cell.text
    text = cell.value.strip()
    upper = text.upper()
    if upper.startswith("#") and (upper in _ERROR_CODES or upper.endswith(("!", "?", "A"))):
        return Err(_ERROR_CODES.get(upper, upper))
    if upper in ("WAHR", "TRUE"):
        return True
    if upper in ("FALSCH", "FALSE"):
        return False
    number = _as_number(text)
    if number is not None:
        return number / 100 if text.endswith("%") else number
    return text


def address(pos: Pos) -> str:
    return f"{column_letters(pos[1])}{pos[0]}"


def _fmt(value) -> str:
    if isinstance(value, float):
        if abs(value) >= 1e6 or value == int(value):
            return number_text(round(value, 6))
        # 28086,77 bleibt 28086,77 (nicht 28086,8), nur ein krummes Ergebnis wird kürzer.
        if abs(value) >= 1:
            return number_text(round(value, 4))
        return number_text(float(f"{value:.6g}"))
    if isinstance(value, Err):
        return _GERMAN_ERRORS.get(value.code, value.code)
    return str(value)


_GERMAN_ERRORS = {"#N/A": "#NV", "#VALUE!": "#WERT!", "#REF!": "#BEZUG!", "#NUM!": "#ZAHL!", "#SPILL!": "#ÜBERLAUF!",
                  "#CALC!": "#KALK!"}


def said(value) -> str:
    """Ein Ergebnis so, wie das deutsche Excel es zeigt, für einen Satz: 7602,5, #WERT!, WAHR. Text bleibt Text."""
    if isinstance(value, bool):
        return "WAHR" if value else "FALSCH"
    if isinstance(value, Err):
        return _GERMAN_ERRORS.get(value.code, value.code)
    if isinstance(value, float):
        return _fmt(value)
    return "" if value is None else str(value)


# So viele Formeln rechnet changed_by() höchstens nach.
MAX_CHANGED = 20_000


def more_formulas(count: int) -> str:
    """„ und 2 weitere Formeln“, „ und eine weitere Formel“."""
    return T(" und eine weitere Formel", " and one more formula") if count == 1 else \
        T(f" und {count} weitere Formeln", f" and {count} more formulas")


def changed_by(sheet: Sheet, direct, overrides: dict, near: bool = False) -> list[tuple[Pos, object, object]]:
    """Die Formeln, die mit overrides anders rechnen: die in direct und alles, was mit ihnen rechnet, je (Zelle,
    jetzt, mit overrides). Zuerst, was außerhalb der Tabelle zusammenfasst (die Auswertung daneben); near: zuerst die
    Formeln, die selbst lesen (der BMI in derselben Zeile)."""
    chain, todo = set(), list(direct)
    while todo and len(chain) <= MAX_CHANGED:
        pos = todo.pop()
        if pos in chain or pos in sheet.unknown or sheet.nodes.get(pos) is None:
            continue
        chain.add(pos)
        todo.extend(sheet.dependents.get(pos, ()))
    if not chain or len(chain) > MAX_CHANGED:
        return []
    _, now = sheet.evaluator(chain=chain)
    _, right = sheet.evaluator(overrides=overrides, chain=chain)
    changed = []
    order = (lambda p: (p not in direct, p)) if near else (lambda p: (sheet.plan.in_table(*p), p not in direct, p))
    for pos in sorted(chain, key=order):
        try:
            before, after = now(pos), right(pos)
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, IndexError, RecursionError):
            continue
        if isinstance(before, float) and isinstance(after, float):
            if math.isclose(before, after, rel_tol=1e-9, abs_tol=1e-9):
                continue
        elif before == after:
            continue
        changed.append((pos, before, after))
    return changed


def calc_value(value, cell: Cell | None = None):
    """Womit eine Formel rechnet: Datumszellen zeigt der Prüfer als „01.10.2024“, Excel rechnet mit der Zahl."""
    if isinstance(value, str) and (cell is None or cell.text is None):
        serial = formel.serial_of(value)
        if serial is not None:
            return serial
    return value


class Foreign:
    """Ein anderes Blatt, auf das Formeln zeigen: seine Werte erst beim Lesen umgerechnet, dazu gelesene Bereiche und
    Suchverzeichnisse. Excel rechnet sie dort, sie ändern sich beim Nachrechnen hier nie. Über Prüfungen hinweg
    gemerkt, solange das Blatt dasselbe ist: Eine kleine Auswertung über ein großes Datenblatt liest es nur einmal."""

    def __init__(self, cells: dict[Pos, Cell]) -> None:
        self.cells = cells
        self.values: dict[Pos, object] = {}
        self.grids: dict = {}
        self.found: dict = {}

    def get(self, pos: Pos):
        try:
            return self.values[pos]
        except KeyError:
            pass
        cell = self.cells.get(pos)
        value = None if cell is None else calc_value(value_of(cell), cell)
        self.values[pos] = value
        return value


# Je Blattname das zuletzt gelesene andere Blatt. Kommt das Blatt neu (andere Zellen), fängt es von vorn an.
_FOREIGN: dict[str, Foreign] = {}
_FOREIGN_LOCK = threading.Lock()
MAX_FOREIGN = 40


def foreign(name: str, cells: dict[Pos, Cell]) -> Foreign:
    with _FOREIGN_LOCK:
        found = _FOREIGN.pop(name, None)
        if found is None or found.cells is not cells:
            found = Foreign(cells)
        _FOREIGN[name] = found
        while len(_FOREIGN) > MAX_FOREIGN:
            del _FOREIGN[next(iter(_FOREIGN))]
        return found


# Größere Bereiche speichert das Blatt nicht Zelle für Zelle: Eine Duplikat-Spalte ZÄHLENWENN($A$2:$A$20001;A2) in
# jeder Zeile wären sonst 400 Millionen Paare (Formel, gelesene Zelle).
WIDE = 256


class Columns:
    """Die vorhandenen Zellen eines Blatts je Spalte, sortiert nach Zeile: Welche liegen in einem Bereich?"""

    def __init__(self, positions) -> None:
        by_col: dict[int, list[int]] = {}
        for r, c in positions:
            by_col.setdefault(c, []).append(r)
        self.rows = {c: sorted(rs) for c, rs in by_col.items()}
        self.cols = sorted(self.rows)

    def inside(self, area: tuple[int, int, int, int]):
        r1, c1, r2, c2 = area
        for c in self.cols[bisect_left(self.cols, c1):bisect_right(self.cols, c2)]:
            rows = self.rows[c]
            for r in rows[bisect_left(rows, r1):bisect_right(rows, r2)]:
                yield r, c

    def count(self, area: tuple[int, int, int, int]) -> int:
        r1, c1, r2, c2 = area
        return sum(bisect_right(rows, r2) - bisect_left(rows, r1)
                   for rows in (self.rows[c] for c in self.cols[bisect_left(self.cols, c1):bisect_right(self.cols, c2)]))


def _within(area: tuple[int, int, int, int], pos: Pos) -> bool:
    return area[0] <= pos[0] <= area[2] and area[1] <= pos[1] <= area[3]


class Refs:
    """Was eine Formel mit großen Bereichen liest, ohne jede Zelle einzeln zu speichern. Verhält sich wie die Menge
    der gelesenen Zellen: die einzelnen Bezüge (auch leere Zellen) und die vorhandenen Zellen in den Bereichen."""

    __slots__ = ("singles", "areas", "_columns")

    def __init__(self, singles, areas, columns: Columns) -> None:
        self.singles, self.areas, self._columns = tuple(dict.fromkeys(singles)), tuple(areas), columns

    def __iter__(self):
        present = self._columns
        for p in self.singles:
            if not any(_within(a, p) for a in self.areas) or not present.count((*p, *p)):
                yield p
        for i, area in enumerate(self.areas):
            before = self.areas[:i]
            for p in present.inside(area):
                if not any(_within(a, p) for a in before):
                    yield p

    def __contains__(self, pos) -> bool:
        return pos in self.singles or (any(_within(a, pos) for a in self.areas) and self._columns.count((*pos, *pos)) > 0)

    def __len__(self) -> int:
        return sum(1 for _ in self)

    def __bool__(self) -> bool:
        return next(iter(self), None) is not None


class Readers:
    """Welche Formeln eine Zelle lesen (wie ein dict: Zelle → Formeln), große Bereiche über ein Verzeichnis je Spalte
    statt Zelle für Zelle."""

    def __init__(self, single: dict[Pos, tuple[Pos, ...]], wide: dict[tuple, list[Pos]], columns: Columns | None) -> None:
        self._single, self._columns = single, columns
        self._by_col: dict[int, list[tuple[int, int, tuple[Pos, ...]]]] = {}
        self._broad: list[tuple[tuple[int, int, int, int], tuple[Pos, ...]]] = []
        for area, formulas in wide.items():
            r1, c1, r2, c2 = area
            if c2 - c1 > 64:
                self._broad.append((area, tuple(formulas)))
                continue
            for c in range(c1, c2 + 1):
                self._by_col.setdefault(c, []).append((r1, r2, tuple(formulas)))

    def get(self, pos: Pos, default=None):
        found = self._single.get(pos, ())
        if self._columns is not None and (self._by_col or self._broad) and self._columns.count((*pos, *pos)):
            r = pos[0]
            more = [f for r1, r2, formulas in self._by_col.get(pos[1], ()) if r1 <= r <= r2 for f in formulas]
            more += [f for area, formulas in self._broad if _within(area, pos) for f in formulas]
            if more:
                found = tuple(dict.fromkeys((*found, *more)))
        return found or default

    def __getitem__(self, pos: Pos):
        found = self.get(pos)
        if found is None:
            raise KeyError(pos)
        return found

    def __contains__(self, pos) -> bool:
        return self.get(pos) is not None


# Spaltenköpfe, die nur „die Zahl dazu“ heißen: In „Kennzahl | Wert“ heißt B12 nach seiner Zeile, nicht „Wert …“.
_PLAIN_HEADS = {"wert", "werte", "value", "values", "ergebnis", "ergebnisse", "result", "zahl"}


class Sheet:
    """Ein Blatt mit geparsten Formeln und Abhängigkeiten.

    name und book (die übrigen Blätter der Mappe: Name → Zellen) braucht es nur für Formeln wie
    „Daten!B3“ oder SVERWEIS auf ein anderes Blatt. Ohne sie gelten solche Formeln als nicht nachrechenbar."""

    def __init__(self, cells: dict[Pos, Cell], plan: Layout, name: str | None = None,
                 book: dict[str, dict[Pos, Cell]] | None = None) -> None:
        self.cells, self.plan, self.name = cells, plan, name
        # Ausgeblendete und weggefilterte Zeilen, wie sie mit den Zellen gelesen sind (xlsx.Cells), sonst None.
        self.hidden_rows: frozenset[int] | None = getattr(cells, "hidden_rows", None)
        self.filtered_rows: frozenset[int] | None = getattr(cells, "filtered_rows", None)
        # Was Prüfungen je Blatt einmal ausrechnen und mehrmals brauchen (Schlüssel: Modul und Zweck).
        self.memo: dict[str, dict] = {}
        self.shown = {pos: value_of(cell) for pos, cell in cells.items()}
        self.inputs = {pos: calc_value(v, cells[pos]) for pos, v in self.shown.items()}
        self.book = {k.casefold(): v for k, v in (book or {}).items() if name is None or k.casefold() != name.casefold()}
        self._other: dict[str, Foreign] = {}
        self._extents: dict[str | None, tuple[int, int]] = {}
        self.nodes: dict[Pos, Node | None] = {}
        self.unknown: set[Pos] = set()
        self.precedents: dict[Pos, set[Pos]] = {}
        self.ranges: dict[Pos, list[tuple[int, int, int, int]]] = {}
        for pos, cell in cells.items():
            if cell.formula is None:
                continue
            try:
                node = formel.parse(cell.formula)
            except formel.Unknown:
                self.nodes[pos] = None
                self.unknown.add(pos)
                continue
            self.nodes[pos] = node
            singles, areas, unknown = formel.references(node, name)
            if unknown or not formel.foreign(node, name) <= set(self.book):
                self.unknown.add(pos)
            self.ranges[pos] = areas
            if sum((r2 - r1 + 1) * (c2 - c1 + 1) for r1, c1, r2, c2 in areas) > WIDE:
                self.precedents[pos] = (singles, areas)  # wird unten zu Refs
                continue
            found = set(singles)
            for r1, c1, r2, c2 in areas:
                found.update(self._inside(r1, c1, r2, c2))
            self.precedents[pos] = found
        # Vorhandene Zellen und Formeln je Spalte, nur wenn eine Formel einen großen Bereich liest.
        wide = [pos for pos, before in self.precedents.items() if isinstance(before, tuple)]
        self.columns: Columns | None = Columns(cells) if wide else None
        self.formula_columns: Columns | None = Columns(self.nodes) if wide else None
        dependents: dict[Pos, list[Pos]] = {}
        spans: dict[tuple, list[Pos]] = {}
        for pos, before in self.precedents.items():
            if isinstance(before, tuple):
                singles, areas = before
                for p in dict.fromkeys(singles):
                    # Liegt die Zelle in einem der Bereiche, kommt die Formel dort über den Bereich dazu.
                    if not any(_within(a, p) for a in areas) or p not in cells:
                        dependents.setdefault(p, []).append(pos)
                for area in dict.fromkeys(areas):
                    spans.setdefault(area, []).append(pos)
                self.precedents[pos] = Refs(singles, areas, self.columns)
                continue
            for p in before:
                dependents.setdefault(p, []).append(pos)
        # Als Tupel: Bei großen Blättern liest fast jede Zelle genau eine Formel, und ein Tupel braucht ein Viertel
        # des Speichers einer Menge.
        self.dependents = Readers({p: tuple(after) for p, after in dependents.items()}, spans, self.columns)
        del dependents
        # Ein Blatt ist größer, als gelesen (xlsx.Cells.cut): Formeln, die dort weiter unten lesen, rechnen mit Zellen,
        # die der Prüfer nicht kennt. Sie rechnet er nicht nach, und Hinweise zu ihnen fallen weg (partial).
        names = {k.casefold(): k for k in (book or {})}
        cuts = {k: c for k, c in ((None, cut_of(cells)), *((k, cut_of(v)) for k, v in self.book.items())) if c is not None}
        # Für den Hinweis im Fenster: je Blatt (wie es heißt, None: dieses) die letzte gelesene Zeile.
        self.cuts: dict[str | None, int] = {names.get(k, k) if k is not None else None: c for k, c in cuts.items()}
        self.partial: set[Pos] = set()
        if cuts:
            for pos, node in self.nodes.items():
                # BEREICH.VERSCHIEBEN und INDIREKT können überall hin zeigen, auch unter das Gelesene.
                if node is not None and (formel.dynamic(node) or
                                         any(row > cuts[k] for k, row in formel.reach(node, name).items() if k in cuts)):
                    self.partial.add(pos)
            self.unknown |= self.partial

    def other(self, sheet: str, pos: Pos):
        """Ein Wert von einem anderen Blatt, so wie Excel ihn zeigt."""
        key = sheet.casefold()
        if self.name is not None and key == self.name.casefold():
            return self.inputs.get(pos)
        return self.foreign(sheet).get(pos)

    def foreign(self, sheet: str) -> Foreign:
        """Ein anderes Blatt der Mappe mit seinen Werten, wie Excel sie zeigt."""
        key = sheet.casefold()
        found = self._other.get(key)
        if found is None:
            if key not in self.book:
                raise formel.Unknown(f"Blatt {sheet}")
            found = self._other[key] = foreign(key, self.book[key])
        return found

    def extent(self, sheet: str | None) -> tuple[int, int]:
        """Bis wohin ein Blatt reicht, damit A:A nicht eine Million leere Zellen liest."""
        key = None if sheet is None or (self.name and sheet.casefold() == self.name.casefold()) else sheet.casefold()
        if key not in self._extents:
            cells = self.cells if key is None else self.book.get(key, {})
            self._extents[key] = max((r for r, _ in cells), default=1), max((c for _, c in cells), default=1)
        return self._extents[key]

    def formula_inputs(self, pos: Pos):
        """Die Formeln, die die Formel in pos liest (für Kreisbezüge und Rechenketten), ohne große Bereiche Zelle für
        Zelle durchzugehen."""
        before = self.precedents.get(pos, ())
        if not isinstance(before, Refs):
            return [p for p in before if p in self.nodes]
        found = dict.fromkeys(p for p in before.singles if p in self.nodes)
        for area in before.areas:
            found.update(dict.fromkeys(self.formula_columns.inside(area)))
        return list(found)

    def read_cells(self) -> set[Pos]:
        """Alle Zellen, die irgendeine Formel liest."""
        found: set[Pos] = set()
        spans: dict[int, list[tuple[int, int]]] = {}
        for before in self.precedents.values():
            if not isinstance(before, Refs):
                found |= before
                continue
            found.update(before.singles)
            for r1, c1, r2, c2 in before.areas:
                for c in self.columns.cols[bisect_left(self.columns.cols, c1):bisect_right(self.columns.cols, c2)]:
                    spans.setdefault(c, []).append((r1, r2))
        # Je Spalte die Bereiche zusammenlegen: Hundert gleiche Bereiche liest es nur einmal.
        for c, parts in spans.items():
            parts.sort()
            start, end = parts[0]
            for r1, r2 in parts[1:] + [(formel.MAX_ROW + 2, formel.MAX_ROW + 2)]:
                if r1 > end + 1:
                    found.update(self.columns.inside((start, c, end, c)))
                    start, end = r1, r2
                else:
                    end = max(end, r2)
        return found

    def _inside(self, r1, c1, r2, c2):
        if (r2 - r1 + 1) * (c2 - c1 + 1) <= len(self.cells):
            return [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1) if (r, c) in self.cells]
        return [p for p in self.cells if r1 <= p[0] <= r2 and c1 <= p[1] <= c2]

    def label(self, pos: Pos) -> str:
        """„D7 (Anzahl Sitze NFP)“ oder „D2 (Total Sitze)“."""
        r, c = pos
        plan = self.plan
        if plan.in_table(r, c):
            name = plan.names.get(r)
            word = plan.word(r)
            if plan.across and word and c != plan.name_column and self._across_line(r):
                return f"{address(pos)} ({word} {plan.headers[c]})"  # „E18 (Umsatz 2028E)“ in einem Finanzplan
            if plan.across and word and c != plan.name_column:
                return f"{address(pos)} ({word})"  # „B23 (Kapitalwert)“ unter den Jahren
            if name and plan.headers[c].strip().casefold() in _PLAIN_HEADS and c != plan.name_column:
                return f"{address(pos)} ({word})"  # „B12 (Mittelwert Kosten)“ in der Spalte „Wert“
            if name and word != name and c != plan.name_column:
                return f"{address(pos)} ({plan.headers[c]}, {word})"  # „I2400 (Laufsumme, Fall-Nr F242398)“
            return f"{address(pos)} ({plan.headers[c]}{f' {name}' if name else ''})"
        word = _left_label(self.cells, r, c)
        return f"{address(pos)} ({word})" if word else address(pos)

    def _across_line(self, r: int) -> bool:
        """Geht die Zeile eines Finanzplans über mindestens zwei Jahre (oder steht nur eine Kennzahl darin)?"""
        lines = self.memo.setdefault("analyse.lines", {})
        if r not in lines:
            plan = self.plan
            lines[r] = sum(1 for c in plan.headers if c != plan.name_column
                           and ((r, c) in self.nodes or _is_number(self.inputs.get((r, c))))) >= 2
        return lines[r]

    def figure_label(self, pos: Pos) -> str:
        """„L2 (Mittelwert Alter)“: Eine Kennzahl neben der Tabelle, die über andere Zeilen rechnet, heißt nach
        dem Wort links von ihr, nicht nach Spalte und Zeile der Tabelle."""
        r, c = pos
        left = self.cells.get((r, c - 1))
        before = self.precedents.get(pos, ())
        if isinstance(before, Refs):
            # Ohne jede Zelle der großen Bereiche durchzugehen: Liest sie eine vorhandene Zelle einer anderen Zeile?
            figure = any(p[0] != r for p in before.singles) or any(
                self.columns.count(a) > (self.columns.count((r, a[1], r, a[3])) if a[0] <= r <= a[2] else 0)
                for a in before.areas)
        else:
            figure = pos in self.nodes and any(p[0] != r for p in before)
        if left is not None and left.formula is None and isinstance(self.shown.get((r, c - 1)), str) and figure:
            return f"{address(pos)} ({left.value})"
        return self.label(pos)

    def readable(self, pos: Pos, formula: str | None = None) -> str:
        return readable(formula if formula is not None else self.cells[pos].formula, pos, self.plan, self._label)

    def _label(self, pos: Pos) -> str:
        """Der Name einer Zahl außerhalb der Tabelle, die eine Formel liest: das Wort links daneben oder darüber."""
        if pos in self.nodes or not _is_number(self.inputs.get(pos)):
            return ""
        word = _left_label(self.cells, *pos)
        return word if 0 < len(word) <= 40 else ""

    # --- Rechnen ---

    def evaluator(self, overrides: dict[Pos, object] | None = None, constants: dict[float, float] | None = None,
                  chain: set[Pos] | None = None):
        """Rechnet die Formeln in chain neu (mit Merken), alles andere kommt aus Excel."""
        overrides = overrides or {}
        memo: dict[Pos, object] = {}
        busy: set[Pos] = set()

        def value(pos: Pos):
            if pos in overrides:
                return overrides[pos]
            node = self.nodes.get(pos)
            if chain is None or pos not in chain or node is None or pos in self.unknown:
                return self.inputs.get(pos)
            if pos in memo:
                return memo[pos]
            if pos in busy:
                raise formel.Unknown("Kreisbezug")
            busy.add(pos)
            result = calc.run(node)
            busy.discard(pos)
            memo[pos] = result
            return result

        def other(sheet: str, pos: Pos):
            if self.name is not None and sheet.casefold() == self.name.casefold():
                return value(pos)
            return self.other(sheet, pos)

        def shared(sheet: str):
            if self.name is not None and sheet.casefold() == self.name.casefold():
                return None
            return self.foreign(sheet)

        calc = formel.Evaluator(value, constants, other=other, extent=self.extent, static=chain is None and not overrides and not constants,
                                shared=shared)
        calc.nested = self.nested
        calc.hidden, calc.filtered = self.hidden_rows, self.filtered_rows
        return calc, value

    def nested(self, pos: Pos) -> bool:
        """Steht in der Zelle ein TEILERGEBNIS oder AGGREGAT? Die zählt ein TEILERGEBNIS darüber nicht mit."""
        found = self.memo.get("analyse.nested")
        if found is None:
            found = self.memo["analyse.nested"] = {
                p for p, node in self.nodes.items()
                if node is not None and any(call.value in ("SUBTOTAL", "AGGREGATE") for call in formel.calls(node))}
        return pos in found


class Stopped(Exception):
    """Eine Prüfung im Hintergrund hört auf, weil etwas Dringenderes kommt (stoppable)."""


_stop = threading.local()


@contextmanager
def stoppable(stop):
    """Solange das läuft, hört analyse() zwischen zwei Prüfschritten auf (Stopped), sobald stop() wahr ist."""
    _stop.check = stop
    try:
        yield
    finally:
        _stop.check = None


def step() -> None:
    check = getattr(_stop, "check", None)
    if check is not None and check():
        raise Stopped()


def analyse(cells: dict[Pos, Cell], plan: Layout, budget: float = TIME_BUDGET, exact: bool = True,
            name: str | None = None, book: dict[str, dict[Pos, Cell]] | None = None,
            filtered: list[tuple[int, int, int, int]] | None = None, rules: list[dict] | None = None) -> Report:
    """exact: Die Werte kommen genau aus Excel. Aus der gespeicherten Datei sind sie gerundet wie angezeigt,
    dann wird nicht nachgerechnet und die Sensitivität verträgt Rundung.

    name und book: das Blatt heißt so, und das sind die übrigen Blätter der Mappe (für Bezüge dorthin).
    filtered: Bereiche, in denen gerade ein Filter Zeilen ausblendet. rules: die eigenen Regeln je Spalte."""
    with formel.remembered():
        return _analyse(cells, plan, budget, exact, name, book, filtered, rules)


def _analyse(cells: dict[Pos, Cell], plan: Layout, budget: float, exact: bool, name: str | None,
             book: dict[str, dict[Pos, Cell]] | None, filtered: list[tuple[int, int, int, int]] | None,
             rules: list[dict] | None) -> Report:
    from . import (auswertung, bloecke, einsetzen, fehlerwerte, finanz, folgen, formeltext, modell, nachschlagen,
                   sensitivitaet, spaltenregeln, statistik, statistiktests, summen, zeilenmodell, zeitraeume, zwischensummen)

    # step(): Hier darf eine Prüfung im Hintergrund aufhören (stoppable), zwischen zwei Schritten.
    sheet = Sheet(cells, plan, name, book)
    step()
    report = Report()
    # Jede Prüfung für sich (guard): Scheitert eine an einer ungewöhnlichen Mappe, laufen die übrigen weiter, und das
    # Fenster sagt, welcher Teil fehlt.
    guard = _Guard(report)
    report.findings += guard("Kreisbezüge", lambda: _circles(sheet), [])
    report.findings += guard("Formelmuster", lambda: _patterns(sheet), [])
    # Quartale aus den Monaten: Ein Block, der eine Zeile auslässt oder doppelt nimmt, sieht nicht anders aus als die
    # übrigen Formeln, nur seine Zeilen passen nicht.
    report.findings += guard("Blöcke einer Spalte", lambda: bloecke.checks(sheet), [])
    report.findings += guard("Verdeckte Zahlen", lambda: _clipped(sheet), [])
    report.findings += guard("Als 0 gezeigt", lambda: _shown_as_zero(sheet), [])
    report.findings += guard("Zwischensummen", lambda: _double_counts(sheet), [])
    # Finanzpläne und Budgets (Zeilen = Größen, Spalten = Jahre): feste Zahlen, mit $ festgehaltene Annahmen und
    # Summen, die in jedem Jahr eine Zeile auslassen.
    more, covered = guard("Finanzplan", lambda: zeilenmodell.checks(sheet, report.findings), ([], set()))
    report.findings = [f for f in report.findings if f.key not in covered] + more
    step()
    if exact:
        report.findings += guard("Nachrechnen", lambda: _recalc(sheet), [])
    report.findings += guard("leere Zellen in Formeln", lambda: _empty_refs(sheet), [])
    step()
    # Die falsche Spalte, ein zu kurzer Bereich oder ein Leerzeichen sagt genauer, warum ein Schlüssel nicht gefunden
    # wird: Diese Zellen fehlen dann in den allgemeinen Hinweisen zum Nachschlagen.
    looked, hidden, explained = guard("SVERWEIS und Nachschlagen", lambda: nachschlagen.checks(sheet), ([], set(), set()))
    report.findings += guard("SVERWEIS und Nachschlagen", lambda: _lookups(sheet, explained), [])
    report.findings = [f for f in report.findings if f.key not in hidden] + looked
    report.findings += guard("schwer prüfbare Formeln", lambda: _complex(sheet), [])
    report.findings += guard("Formeln als Text", lambda: formeltext.checks(sheet), [])
    report.findings += guard("Daten", lambda: _data(sheet), [])
    step()
    more, covered = guard("Auswertung der Daten", lambda: auswertung.checks(sheet, filtered), ([], set()))
    # Eine Laufsumme, die eine Zeile zu weit reicht, soll in ihrer Zeile enden, nicht bis ans Ende der Spalte.
    moved = sheet.memo.get("moved", {})
    more = [f for f in more if not (f.key.startswith("range:") and f.key.partition(":")[2] in moved)]
    report.findings = [f for f in report.findings if f.key not in covered] + more
    # Ein unmöglicher BMI, weil die Größe in Metern dasteht: ein Fehler beim Ursprung, nicht zwei.
    report.findings = guard("Folgen beim Ursprung", lambda: folgen.fold(sheet, report.findings), report.findings)
    step()
    # Summen, die nicht zu ihren Teilen passen: Welcher Teil fehlt, sagt mehr als der allgemeine Hinweis zur Zelle.
    footed, explained = guard("Summen und ihre Teile", lambda: summen.checks(sheet, *auswertung.table(sheet)), ([], set()))
    report.findings = [f for f in report.findings if not (f.key.partition(":")[2] in explained and
                                                          f.key.partition(":")[0] in _FOOTED)] + footed
    # Berichte mit Zwischensummen: eine Gesamtsumme, die eine Gruppe auslässt oder alles doppelt zählt.
    grouped, covered = guard("Zwischensummen", lambda: zwischensummen.checks(sheet), ([], set()))
    report.findings = [f for f in report.findings if f.key not in covered] + grouped
    step()
    # Statistik über Codes, Nullen und leere Zellen, die falsche Streuung, Anteile von der falschen Summe.
    stats, covered = guard("Statistik", lambda: statistik.checks(sheet), ([], set()))
    report.findings = [f for f in report.findings if f.key not in covered] + stats
    step()
    # T.TEST, Konfidenzintervalle und Prozentpunkte, die falsch aufgesetzt sind.
    report.findings += guard("Statistiktests", lambda: statistiktests.checks(sheet), [])
    step()
    # Alter und Zeiträume: /365 statt DATEDIF, HEUTE() in Fällen, Tage mit oder ohne den ersten Tag.
    report.findings += guard("Zeiträume", lambda: zeitraeume.checks(sheet), [])
    step()
    # Die eigenen Regeln je Spalte: Was sie verlangen, sagt mehr als die allgemeinen Grenzen für dieselbe Spalte.
    ruled, covered = guard("eigene Regeln", lambda: spaltenregeln.checks(sheet, rules or [], report.findings), ([], set()))
    report.findings = [f for f in report.findings if f.key not in covered] + ruled
    step()
    # Anteile, die 100 % ergeben müssen, Abzinsen mit der falschen Annahme, die eigene Prüfung des Modells.
    from . import doppelt, geschaeftsjahr, listen, massstab, spaltenwahl

    report.findings += guard("Listen", lambda: listen.checks(sheet), [])
    report.findings += guard("Spalte zum Namen", lambda: spaltenwahl.checks(sheet, report.findings), [])
    report.findings += guard("Doppelte Schlüssel", lambda: doppelt.checks(sheet), [])
    report.findings += guard("Einheiten je Zeile", lambda: massstab.mixed_rows(sheet), [])
    report.findings += guard("Verrutschte Sätze", lambda: massstab.shifted_rates(sheet), [])
    report.findings += guard("Geschäftsjahr und Monatsende", lambda: geschaeftsjahr.checks(sheet), [])
    report.findings += guard("Modell", lambda: modell.checks(sheet, report.findings), [])
    # Controlling und Buchhaltung: Was diese Prüfungen zu einer Zelle sagen, nennt die Ursache genauer.
    money = guard("Finanzen", lambda: finanz.checks(sheet, auswertung.table(sheet)[0] or []), [])
    at = {f.key.partition(":")[2] for f in money}
    # Was die Korrektur ohnehin umschreibt (C6 über 100 %, weil der Nenner mitrutscht), ist eine Folge, kein eigener Fehler.
    at |= {cell for f in money if f.level == "fehler" and f.fix for cell in (f.fix.get("cells") or {})}
    report.findings = [f for f in report.findings if f.key.partition(":")[2] not in at] + money
    step()
    guard("Einsetzen", lambda: einsetzen.attach(sheet, report.findings), None)
    # Sensitivitätstabellen als Gegenprobe: Mit den Annahmen des Modells muss die Tabelle das Ergebnis zeigen.
    report.findings += guard("Sensitivitätstabellen", lambda: sensitivitaet.checks(sheet, report.findings), [])
    _one_per_cell(report)
    _impossible_from_pattern(sheet, report)
    step()
    # Fehlerwerte (#DIV/0!, #NAME? …): wo sie entstehen und warum. Behebt eine Korrektur oben sie schon, sagt deren
    # Hinweis sie mit.
    errors, covered = guard("Fehlerwerte", lambda: fehlerwerte.checks(sheet, report.findings, hidden), ([], set()))
    report.findings = [f for f in report.findings if f.key not in covered] + errors
    report.drivers = guard("Sensitivität", lambda: _sensitivity(sheet, time.monotonic() + budget, 1e-6 if exact else 1e-2), [])
    formulas = [p for p in sheet.nodes]
    inputs = {p for p in sheet.read_cells() if p not in sheet.nodes}
    outputs = [p for p in formulas if not sheet.dependents.get(p)]
    report.partial, report.cuts = sheet.partial, sheet.cuts
    report.stats = {
        "formeln": len(formulas),
        "eingaben": len(inputs),
        "ergebnisse": len(outputs),
        "tiefe": guard("Rechenkette", lambda: _longest_chain(sheet), 0),
        "nicht_nachrechenbar": len(sheet.unknown),
    }
    return report


def _impossible_from_pattern(sheet: Sheet, report: Report) -> None:
    """G10 rechnet ohne REST und zeigt darum -16 Stunden: Der unmögliche Wert ist die Folge der falschen Formel in
    derselben Zelle. Ein Hinweis mit beidem, die Korrektur der Formel behebt auch den Wert."""
    bad = sheet.memo.get("auswertung.bad", {})
    patterns = {f.key.partition(":")[2]: f for f in report.findings
                if f.key.startswith("pattern:") and f.level == "fehler" and f.fix}
    drop = set()
    for item in report.findings:
        at = item.key.partition(":")[2]
        if not item.key.startswith("impossible:") or at not in patterns or item.key not in bad:
            continue
        col, _, values, rows, why = bad[item.key]
        if len(rows) != 1:
            continue
        patterns[at].text += T(f" Darum zeigt {at} {_fmt(values[rows[0]])}, und {why[0].lower()}{why[1:]}.",
                               f" That is why {at} shows {_fmt(values[rows[0]])} and {why[0].lower()}{why[1:]}.")
        drop.add(item.key)
    report.findings = [f for f in report.findings if f.key not in drop]


def _one_per_cell(report: Report) -> None:
    """Sagt ein genauerer Hinweis zur selben Zelle, was falsch ist, reicht er: „Rechnet anders als die übrigen“ stünde
    sonst noch einmal da (SVERWEIS mit der falschen Spalte, ein zu kurzer Bereich). Bringt er keine eigene Korrektur
    mit, bekommt er die nachgerechnete Formel des Musters, wenn er sagt, dass der Bereich nicht stimmt."""
    patterns = {f.key.partition(":")[2]: f for f in report.findings if f.key.startswith("pattern:")}
    merged = set()
    for item in report.findings:
        at = item.key.partition(":")[2]
        if item.level != "fehler" or item.key.startswith("pattern:") or at not in patterns:
            continue
        if item.fix and at in (item.fix.get("cells") or {}):
            merged.add(at)
        elif not item.fix and item.key.partition(":")[0] in ("partial", "range") and patterns[at].fix:
            item.fix = patterns[at].fix
            merged.add(at)
    report.findings = [f for f in report.findings if not (f.key.startswith("pattern:") and f.key.partition(":")[2] in merged)]


# --- Rechenkette ---


# Hinweise zu einer Summe, die summen.py genauer sagt.
_FOOTED = {"label", "count", "range", "pattern", "double", "total", "group"}


def _circles(sheet: Sheet) -> list[Finding]:
    """Kreisbezüge (Tarjan): Zellen, die über Umwege von sich selbst abhängen."""
    index: dict[Pos, int] = {}
    low: dict[Pos, int] = {}
    stack: list[Pos] = []
    on: set[Pos] = set()
    groups: list[list[Pos]] = []
    counter = 0
    for start in sheet.precedents:
        if start in index:
            continue
        work = [(start, iter(sorted(sheet.formula_inputs(start))))]
        index[start] = low[start] = counter
        counter += 1
        stack.append(start)
        on.add(start)
        while work:
            pos, children = work[-1]
            advanced = False
            for child in children:
                if child not in sheet.precedents:
                    continue
                if child not in index:
                    index[child] = low[child] = counter
                    counter += 1
                    stack.append(child)
                    on.add(child)
                    work.append((child, iter(sorted(sheet.formula_inputs(child)))))
                    advanced = True
                    break
                if child in on:
                    low[pos] = min(low[pos], index[child])
            if advanced:
                continue
            work.pop()
            if work:
                parent = work[-1][0]
                low[parent] = min(low[parent], low[pos])
            if low[pos] == index[pos]:
                group = []
                while True:
                    item = stack.pop()
                    on.discard(item)
                    group.append(item)
                    if item == pos:
                        break
                if len(group) > 1 or pos in sheet.precedents.get(pos, ()):
                    groups.append(sorted(group))
    return [
        Finding(f"circle:{address(g[0])}",
                T(f"Kreisbezug: {', '.join(address(p) for p in g[:6])} {'hängt' if len(g) == 1 else 'hängen'} über Umwege von "
                  f"{'sich' if len(g) == 1 else 'sich gegenseitig'} selbst ab.",
                  f"Circular reference: {', '.join(address(p) for p in g[:6])} "
                  f"{'depends on itself' if len(g) == 1 else 'depend on each other'} in a roundabout way."),
                "fehler", T("Excel rechnet so nur mit Iteration und oft falsch. Prüfe, welcher Bezug im Kreis nicht dorthin gehört.",
                            "Excel can only calculate this with iteration, and often gets it wrong. Check which reference "
                            "in the circle does not belong there."))
        for g in groups[:3]
    ]


def _runs(rows: list[int]) -> list[list[int]]:
    runs: list[list[int]] = []
    for r in rows:
        if runs and r - runs[-1][-1] <= 1:
            runs[-1].append(r)
        else:
            runs.append([r])
    return runs


def _patterns(sheet: Sheet) -> list[Finding]:
    """Formeln, die aus dem Muster ihrer Spalte (oder Zeile) fallen."""
    found: list[Finding] = []
    by_col: dict[int, list[int]] = {}
    by_row: dict[int, list[int]] = {}
    for (r, c), cell in sheet.cells.items():
        if cell.formula is not None:
            by_col.setdefault(c, []).append(r)
            by_row.setdefault(r, []).append(c)
    flagged: set[Pos] = set()
    columns = [[(r, c) for r in run] for c, rows in sorted(by_col.items()) for run in _runs(sorted(rows))]
    lines = [[(r, c) for c in run] for r, cols in sorted(by_row.items()) for run in _runs(sorted(cols))]
    if sheet.plan.across:
        # In einem Finanzplan rechnen die Jahre gleich, eine Spalte „Summe 2026-2030“ oder „CAGR“ daneben anders.
        from .zeilenmodell import kind

        heads = sheet.plan.headers
        kinds = {c: kind(heads[c]) if c in heads else "" for c in {p[1] for line in lines for p in line}}
        lines = [part for line in lines for part in _split(line, lambda a, b: kinds[a[1]] != kinds[b[1]])]
    # Eine Formel, die genau wie ihre Spalte rechnet (der Anteil in der Summenzeile), fällt nicht aus ihrer Zeile.
    along_columns, along_rows = _fitting(columns, sheet), _fitting(lines, sheet)
    for group in columns:
        found += _odd_ones(group, sheet, flagged, "Spalte", along_rows)
    for group in lines:
        found += _odd_ones(group, sheet, flagged, "Zeile", along_columns)
    return found[:6]


def _split(items: list, apart) -> list[list]:
    """items in Stücke geteilt, wo apart(vorher, nachher) gilt."""
    parts: list[list] = []
    for item in items:
        if parts and not apart(parts[-1][-1], item):
            parts[-1].append(item)
        else:
            parts.append([item])
    return parts


def _common(group: list[Pos], sheet: Sheet):
    """Die Formeln, die zum Muster zählen, ihr Muster und wie oft es vorkommt, oder None."""
    if len(group) < PATTERN_MIN + 1:
        return None
    along = "Spalte" if group[0][1] == group[-1][1] else "Zeile"
    shapes = {p: pattern(sheet.cells[p].formula, p, along) for p in group}
    # Eine Summe am Ende einer Spalte ist kein Kopierfehler und zählt nicht zum Muster, außer sie rechnet genau wie
    # die Zeilen darüber (die letzte Zeile einer Laufsumme).
    usual = Counter(shapes[p] for p in group[:-1]).most_common(1)[0][0]
    if _is_total(sheet, group[-1]) and shapes[group[-1]] != usual \
            and not (along == "Zeile" and sheet.plan.across) and not (along == "Zeile" and _moved_total(sheet, group, shapes, usual)):
        del shapes[group[-1]]
        group = group[:-1]
        if len(group) < PATTERN_MIN + 1:
            return None
    common, count = Counter(shapes.values()).most_common(1)[0]
    # Der Startwert vorne (Umsatz im ersten Jahr =$B$6, dann =B22*(1+$B$7)) kann nicht wie die übrigen rechnen: Die
    # Formel des Musters zeigte dort auf den Namen oder die Kopfzeile. Er zählt nicht gegen das Muster, sonst ginge
    # ein Fehler in einem der vier Folgejahre im Startwert unter (3 von 5 statt 3 von 4).
    if count / len(group) < PATTERN_SHARE and len(group) > PATTERN_MIN + 1 and shapes[group[0]] != common:
        example = next(p for p in group if shapes[p] == common)
        start = shift(sheet.cells[example].formula, group[0][0] - example[0], group[0][1] - example[1])
        if list(shapes.values()).count(shapes[group[0]]) == 1 and not _sensible(sheet, start):
            del shapes[group[0]]
            group = group[1:]
    if count >= PATTERN_MIN and count / len(group) < PATTERN_SHARE and along == "Zeile" and not sheet.plan.across \
            and len(group) - count <= MAX_ALIKE:
        # In einer Zeile unter einer Liste (Mittelwert je Spalte) rechnet MITTELWERT(Alter) so viel wie
        # MITTELWERTWENN(Alter;"<>9"), wo keine 9 steht: Die Formel folgt dem Muster im Ergebnis.
        example = next(p for p in group if shapes[p] == common)
        count += sum(1 for p in group if shapes[p] != common and _alike(sheet, p, example))
    if count < PATTERN_MIN or count / len(group) < PATTERN_SHARE:
        return None
    return group, shapes, common, count


MAX_ALIKE = 4


def _control_row(sheet: Sheet, pos: Pos) -> bool:
    from .summen import _CONTROL, _row_text

    return bool(_CONTROL.search(_row_text(sheet, pos)))


def _moved_total(sheet: Sheet, group: list[Pos], shapes: dict, usual) -> bool:
    """Die Summe am Ende der Zeile rechnet wie ihre Nachbarn (=SUMME(I48:I51) neben lauter =SUMME(B48:B52)), nur
    endet ihr Bereich woanders: ein Kopierfehler, keine andere Rechnung."""
    example = next((p for p in group[:-1] if shapes[p] == usual), None)
    if example is None:
        return False
    last = group[-1]
    expected = shift(sheet.cells[example].formula, last[0] - example[0], last[1] - example[1])
    return _moved_end(sheet.cells[last].formula or "", expected) is not None


def _quote_column(sheet: Sheet, pos: Pos) -> bool:
    """Steht pos unter einer Spalte von Quoten (Marge, Anteil, %): Die Zellen darüber teilen zwei Zahlen."""
    above = [sheet.cells.get((r, pos[1])) for r in range(max(1, pos[0] - 12), pos[0])]
    shares = [c for c in above if c is not None and c.formula]
    if len(shares) < 3:
        return False
    def divides(formula: str) -> bool:
        inner = _unguarded(re.sub(r"\s+", "", formula.lstrip("=")).upper())
        return "/" in inner and not re.search(r"\b(SUM|SUMME|AVERAGE)\(", inner)
    return sum(1 for c in shares if divides(c.formula)) >= 0.8 * len(shares)


def _alike(sheet: Sheet, pos: Pos, example: Pos) -> bool:
    """Ergibt die Formel des Musters in pos dasselbe wie die Formel dort?"""
    expected = shift(sheet.cells[example].formula, pos[0] - example[0], pos[1] - example[1])
    calc, _ = sheet.evaluator()
    try:
        want, have = calc.run(formel.parse(expected)), calc.run(formel.parse(sheet.cells[pos].formula))
    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError, IndexError, KeyError):
        return False
    if isinstance(want, bool) or isinstance(have, bool) or not isinstance(want, float) or not isinstance(have, float):
        return want == have and want is not None
    return math.isfinite(want) and abs(want - have) <= 1e-9 * max(1.0, abs(want))


def _fitting(groups: list[list[Pos]], sheet: Sheet) -> set[Pos]:
    out: set[Pos] = set()
    for group in groups:
        found = _common(group, sheet)
        if found is not None:
            out |= {p for p, shape in found[1].items() if shape == found[2]}
    return out


def _odd_ones(group: list[Pos], sheet: Sheet, flagged: set[Pos], along: str, fits: frozenset | set = frozenset()) -> list[Finding]:
    found = _common(group, sheet)
    if found is None:
        return []
    group, shapes, common, count = found
    example = next(p for p in group if shapes[p] == common)
    result = []
    for pos in group:
        if shapes[pos] == common or pos in flagged or pos in fits:
            continue
        expected = shift(sheet.cells[example].formula, pos[0] - example[0], pos[1] - example[1])
        moved = _moved_end(sheet.cells[pos].formula, expected) if _is_total(sheet, pos) else None
        # Zwischensummen in einer Spalte rechnen natürlich anders. In einem Finanzplan ist die SUMME einer Zeile
        # (EBITDA je Jahr) keine Zwischensumme ihrer Nachbarn: Jedes Jahr rechnet gleich.
        if _is_total(sheet, pos) and moved is None and not _same_cells(sheet.cells[pos].formula, expected) \
                and not (along == "Zeile" and sheet.plan.across):
            continue
        if not _sensible(sheet, expected):
            continue  # Die erste Zeile einer Laufsumme kann nicht wie die übrigen rechnen.
        if along == "Zeile" and _quote_column(sheet, pos) and re.match(r"=?\s*SUM\(", expected.upper()):
            continue  # Eine Marge summiert man nicht: Die Summenzeile teilt die Summen (=K30/H30), das ist richtig.
        if along == "Zeile" and not sheet.plan.across and _alike(sheet, pos, example):
            continue  # Unter einer Liste je Spalte etwas anderes, das hier dasselbe ergibt: kein Fehler.
        if along == "Spalte" and _control_row(sheet, pos):
            continue  # „Kontrolle: … − Gesamt (0 = OK)“ in der Summenspalte prüft, statt zu summieren.
        if _same_meaning(sheet, sheet.cells[pos].formula, expected):
            continue  # Nur die $ anders (=C27*C10 statt =C27*$C$10 in Spalte C) oder C5 statt C27, wo C27 =C5 ist.
        flagged.add(pos)
        span = f"{address(group[0])}:{address(group[-1])}"
        text = T(f"{sheet.label(pos)} rechnet anders als die übrigen {count} Formeln in {span}: "
                 f"{sheet.readable(pos)} statt {sheet.readable(pos, expected)}.",
                 f"{sheet.label(pos)} calculates differently from the other {count} formulas in {span}: "
                 f"{sheet.readable(pos)} instead of {sheet.readable(pos, expected)}.")
        tip = T("Oft ein Kopierfehler oder ein vergessenes $. Ist es Absicht, schreib eine Notiz in die Zelle.",
                "Often a copy error or a missing $. If it is intentional, add a note to the cell.")
        if moved is not None:
            text = _moved_text(sheet, text, *moved) + _shows(sheet, pos, expected, along)
            sheet.memo.setdefault("moved", {})[address(pos)] = True
        elif sheet.plan.across and along == "Zeile":
            # Im Finanzplan auch, was das Jahr dadurch zeigt: Daran sieht man den Fehler in Excel sofort.
            text += _shows(sheet, pos, expected, along)
        fixed = _checked(sheet, expected, sheet.cells[pos].value)
        if fixed:
            # Die Formel des Musters zum Einsetzen, nachgerechnet. Eine verschobene Summe ist sicher keine Absicht.
            tip = T(f"Korrigiert: {fixed}.", f"Fixed: {fixed}.") + (
                "" if moved is not None else T(" Ist es Absicht, schreib eine Notiz in die Zelle.",
                                               " If it is intentional, add a note to the cell."))
        result.append(Finding(f"pattern:{address(pos)}", text, "fehler", tip))
    return result


_CELL_TEXT = re.compile(r"(?<![A-Za-z0-9_.!$\\])\$?([A-Z]{1,3})\$?(\d{1,7})(?![0-9A-Za-z_(!:])")


def _same_meaning(sheet: Sheet, formula: str, expected: str) -> bool:
    """Rechnet die Formel genau wie die des Musters, mit denselben Zellen? Die $ zählen nicht, und ein Bezug auf eine
    Zelle, die nur =C5 enthält, ist ein Bezug auf C5. Dann ist der Unterschied in dieser Zelle kein Fehler."""
    if "!" in formula or "!" in expected or ":" in formula + expected:
        return False

    def origin(match: re.Match) -> str:
        pos = (int(match.group(2)), column_number(match.group(1)))
        for _ in range(8):
            cell = sheet.cells.get(pos)
            text = (cell.formula or "").strip().lstrip("=").strip() if cell is not None else ""
            inner = _CELL_TEXT.fullmatch(text)
            if inner is None:
                break
            pos = (int(inner.group(2)), column_number(inner.group(1)))
        return address(pos)

    def plain(text: str) -> str:
        return _unguarded(re.sub(r"\s+", "", _CELL_TEXT.sub(origin, text.lstrip("="))).upper())

    return plain(formula) == plain(expected)


# Ein Schutz vor dem Teilen durch 0 ändert nichts an der Rechnung: =WENN(H18=0;0;K18/H18) rechnet wie =K18/H18, solange
# H18 nicht 0 ist. Die Summenzeile darunter (=K30/H30) rechnet darum wie die Monate.
_GUARDS = (
    (re.compile(r'IF\((\$?[A-Z]+\$?\d+)=0,(?:0|""),(.+)\)'), 2),
    (re.compile(r'IF\((\$?[A-Z]+\$?\d+)<>0,(.+),(?:0|"")\)'), 2),
    (re.compile(r'IFERROR\((.+),(?:0|"")\)'), 1),
)


def _unguarded(text: str) -> str:
    for rx, group in _GUARDS:
        found = rx.fullmatch(text)
        if found is None:
            continue
        inner = found.group(group)
        if group == 2 and not re.search(rf"/\(?{re.escape(found.group(1))}(?![0-9])", inner):
            continue  # Geteilt wird durch etwas anderes: Die Bedingung ist dann mehr als ein Schutz.
        try:
            formel.parse(inner)
        except formel.Unknown:
            continue
        return inner
    return text


def _same_cells(formula: str, expected: str) -> bool:
    """Liest die Formel dieselben Zellen wie die des Musters, nur anders verrechnet? Dann ist sie keine Zwischensumme
    (die zählt andere Zellen), sondern ein Kopierfehler: =SUMME(BEREICH.VERSCHIEBEN(B10;1-$F$1;0;$F$1-1;1)) unter
    lauter Fenstern über $F$1 Monate."""
    def reads(node) -> list:
        out = [(node.kind, node.value)] if node.kind in ("ref", "range", "xref", "xrange", "unknown") else []
        for child in node.args:
            out += reads(child)
        return out

    try:
        have, want = formel.parse(formula), formel.parse(expected)
    except formel.Unknown:
        return False
    found = reads(have)
    return bool(found) and sorted(map(repr, found)) == sorted(map(repr, reads(want)))


def _checked(sheet: Sheet, formula: str, like: str = "") -> str | None:
    """„=SUMME($F$2:F2400), nachgerechnet 22112387,14“ für den Tipp, oder None, wenn die Formel hier nicht ohne
    Fehlerwert rechnet. like: wie die Zelle jetzt aussieht; ein Datum („01.04.2026“) zeigt das Ergebnis als Datum."""
    from .bereiche import _german_formula, _num
    from .einsetzen import _DATE_TEXT, _date_text

    calc, _ = sheet.evaluator()
    try:
        value = calc.run(formel.parse(formula))
    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError, IndexError, KeyError):
        return None
    if isinstance(value, bool):
        shown = "WAHR" if value else "FALSCH"
    elif isinstance(value, float) and math.isfinite(value):
        shown = _date_text(value, False) if _DATE_TEXT.fullmatch(like.strip()) else _num(value)
    elif isinstance(value, str) and value:
        shown = T(f"„{value}“", f"“{value}”")
    else:
        return None
    return T(f"={_german_formula(formula)}, nachgerechnet {shown}",
             f"={_german_formula(formula)}, recalculated this gives {shown}")


MAX_LISTED_MOVED = 3
_AREA = re.compile(r"(?<![A-Za-z0-9_.$\\])\$?([A-Z]{1,3})\$?(\d{1,7}):\$?([A-Z]{1,3})\$?(\d{1,7})(?![A-Za-z0-9_(])")


def _moved_end(formula: str, expected: str) -> tuple[list[Pos], int, list[Pos], int] | None:
    """Dieselbe Summe wie im Muster, nur endet ein Bereich woanders. Das ist keine Zwischensumme, sondern ein
    Kopierfehler: =SUMME($E$2:E2399) in Zeile 2400 einer Laufsumme, =SUMME(B7:C7) unter lauter =SUMME(B:D).
    Gibt die ersten Zellen, die so fehlen, und wie viele es sind, dann dasselbe für die zu viel gezählten, sonst None."""
    try:
        have, want = formel.parse(formula), formel.parse(expected)
    except formel.Unknown:
        return None
    if have.kind != "call" or want.kind != "call" or have.value != want.value or len(have.args) != len(want.args):
        return None
    pairs = []
    for got, wanted in zip(have.args, want.args):
        if got.kind == wanted.kind == "range" and got.value[:2] == wanted.value[:2]:
            if got.value != wanted.value:
                pairs.append((got.value, wanted.value))
        elif got != wanted:
            return None
    if not pairs:
        return None
    missing: list[Pos] = []
    extra: list[Pos] = []
    lacking = surplus = 0
    for got, wanted in pairs:
        for outer, inner, cells in ((wanted, got, missing), (got, wanted, extra)):
            r1, c1, r2, c2 = outer
            both = (min(r2, inner[2]) - r1 + 1) * (min(c2, inner[3]) - c1 + 1)
            size = (r2 - r1 + 1) * (c2 - c1 + 1) - both
            if cells is missing:
                lacking += size
            else:
                surplus += size
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if len(cells) >= MAX_LISTED_MOVED:
                        break
                    if not (r <= inner[2] and c <= inner[3]):
                        cells.append((r, c))
                if len(cells) >= MAX_LISTED_MOVED:
                    break
    return missing, lacking, extra, surplus


def _moved_text(sheet: Sheet, text: str, missing: list[Pos], lacking: int, extra: list[Pos], surplus: int) -> str:
    """Sagt, welche Zellen die verschobene Summe auslässt oder zu viel zählt."""

    def listed(cells: list[Pos], total: int) -> str:
        if total == 1:
            value = sheet.inputs.get(cells[0])
            # Mit der Zahl, die fehlt: „So fehlt F2400 (Kosten CHF, Fall-Nr F242398) mit 7512,78.“
            if isinstance(value, float) and not isinstance(value, bool) and math.isfinite(value):
                from .bereiche import _num

                return T(f"{sheet.label(cells[0])} mit {_num(value)}", f"{sheet.label(cells[0])} with {_num(value)}")
            return sheet.label(cells[0])
        shown = ", ".join(address(p) for p in cells)
        return shown + (T(f" und {total - len(cells)} weitere", f" and {total - len(cells)} more") if total > len(cells) else "")

    if lacking:
        text += T(f" So {'fehlt' if lacking == 1 else 'fehlen'} {listed(missing, lacking)}.",
                  f" So {listed(missing, lacking)} {'is' if lacking == 1 else 'are'} missing.")
    if surplus:
        text += T(f" So zählt sie {listed(extra, surplus)} zu viel.", f" So it also counts {listed(extra, surplus)}, which does not belong.")
    return text


def _shows(sheet: Sheet, pos: Pos, expected: str, along: str = "Spalte") -> str:
    """„ I2400 zeigt darum 22104874,36 statt 22112387,14.“: was die Zelle zeigt und was die Formel des Musters dort
    ergibt, beide Zahlen. along="Zeile" in einem Finanzplan: Der Nachbar ist das Vorjahr links, nicht die Zeile
    darüber („E28 zeigt darum -573,89 (so viel wie D28) statt -638,21“)."""
    from .bereiche import _num

    now = sheet.inputs.get(pos)
    calc, _ = sheet.evaluator()
    try:
        want = calc.run(formel.parse(expected))
    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError, IndexError, KeyError):
        return ""
    if not all(isinstance(v, float) and not isinstance(v, bool) and math.isfinite(v) for v in (now, want)) or now == want:
        return ""
    if _num(now) == _num(want):
        return ""  # Eine fest getippte Zahl, die (noch) gleich ist wie die Annahme: Sie zeigt dasselbe.
    # Eine Laufsumme ohne die eigene Zeile zeigt dasselbe wie die Zelle darüber: Das sieht man in Excel sofort.
    above = (pos[0], pos[1] - 1) if along == "Zeile" and sheet.plan.across else (pos[0] - 1, pos[1])
    same = T(f" (so viel wie {address(above)})", f" (the same as {address(above)})") \
        if sheet.inputs.get(above) == now and (above in sheet.nodes) else ""
    return T(f" {address(pos)} zeigt darum {_num(now)}{same} statt {_num(want)}.",
             f" So {address(pos)} shows {_num(now)}{same} instead of {_num(want)}.")


_WINDOWED = {"AVERAGE", "SUM", "MEDIAN", "MIN", "MAX", "STDEV", "STDEV.S", "STDEV.P", "COUNT", "PRODUCT"}


def _sensible(sheet: Sheet, formula: str) -> bool:
    """Würde die Formel des Musters hier überhaupt gehen? Nicht, wenn sie auf Kopfzeile oder Text zeigt."""
    if "#BEZUG!" in formula:
        return False
    # Ein Bereich, der hier rückwärts liefe, passt nicht: =SUMME($B$6:B2) über dem Neubeginn einer Laufsumme in B6.
    for area in _AREA.finditer(re.sub(r'"(?:[^"]|"")*"', '""', formula)):
        c1, r1, c2, r2 = area.groups()
        if int(r1) > int(r2) or column_number(c1) > column_number(c2):
            return False
    try:
        singles, _, unknown = formel.references(formel.parse(formula))
    except formel.Unknown:
        return False
    header = sheet.plan.header_row or 0
    # Auf Leeres oder über die Kopfzeile hinaus: Dann ist die andere Formel hier wohl richtig. In einem Finanzplan
    # mit durchgezählten Jahren (Jahr | 0 | 1 | 2) rechnet der Diskontfaktor mit der Zahl in der Kopfzeile. Eine
    # Zahl neben der Tabelle auf Höhe der Kopfzeile ist eine Annahme („MwSt | 0,081“ in H1:I1), kein Kopf.
    if any(sheet.shown.get(p) is None or (header - 1 <= p[0] <= header
                                          and not (sheet.plan.across and p[0] == header and _is_number(sheet.inputs.get(p)))
                                          and not (p[1] not in sheet.plan.headers and _is_number(sheet.inputs.get(p))))
           for p in singles):
        return False
    # Ein gleitendes Fenster (MITTELWERT(B3:B5)) in der ersten Zeile: Nach oben verschoben reicht es über die
    # Kopfzeile hinaus (B1:B3). Dort fehlen die Monate davor, das Fenster der übrigen Zeilen geht hier nicht.
    if header and any(call.value in _WINDOWED and arg.kind == "range" and arg.value[0] <= header < arg.value[2]
                      and arg.value[2] < formel.MAX_ROW
                      for call in formel.calls(formel.parse(formula)) for arg in call.args):
        return False
    # Ein Datum oder eine Uhrzeit zeigt der Prüfer als „08:30“, Excel rechnet mit der Zahl: Das ist kein Text.
    texts = {p for p in singles if isinstance(sheet.shown.get(p), str) and not _is_number(sheet.inputs.get(p))}
    if not unknown and not texts:
        return True
    # Text als Suchwert (SVERWEIS(D8;…)), als Bedingung (SUMMEWENN(…;A5;…)) oder ein anderes Blatt: Geht, wenn jede
    # Suche hier findet, was sie sucht (nicht erst WENNFEHLER rettet sie), jede Bedingung mit Text aus einer Zelle
    # mindestens eine Zeile trifft und die Formel ohne Fehler rechnet. Trifft sie keine (eine Zeile „Durchschnitt“
    # unter einer Übersicht mit SUMMEWENNS) oder rechnet sie mit dem Text, gehört hier eine andere Formel hin.
    from .bereiche import _CONDITIONAL, _ranges_of

    node = formel.parse(formula)
    lookups, counts, keys = [], [], set()
    for call in formel.calls(node):
        if call.value in _LOOKUP_KEYS and call.args:
            lookups.append(call)
            if call.args[0].kind == "ref":
                keys.add(call.args[0].value)
        elif call.value in _CONDITIONAL and (parts := _ranges_of(call)) is not None:
            tests = [t for t in parts[1] if t[1].kind == "ref" and t[1].value in texts]
            if tests:
                counts.append(formel.Node("call", "COUNTIFS", tuple(x for t in parts[1] for x in t)))
                keys |= {t[1].value for t in tests}
    if texts - keys:
        return False
    calc, _ = sheet.evaluator()
    try:
        found = [calc.run(c) for c in lookups + [node]]
        hits = [calc.run(c) for c in counts]
    except (formel.Unknown, RecursionError):
        return False
    return all(v is not None and not isinstance(v, formel.Err) for v in found) and \
        all(isinstance(v, float) and v > 0 for v in hits)


_LOOKUP_KEYS = {"VLOOKUP", "HLOOKUP", "MATCH", "XLOOKUP", "XMATCH", "LOOKUP"}


def _is_number(value) -> bool:
    return isinstance(value, float) and not isinstance(value, bool) and math.isfinite(value)


def _is_total(sheet: Sheet, pos: Pos) -> bool:
    node = sheet.nodes.get(pos)
    return node is not None and node.kind == "call" and node.value in _SUMS


def _double_counts(sheet: Sheet) -> list[Finding]:
    """Eine Summe, die eine Zwischensumme und deren Zeilen zusammen zählt."""
    found = []
    # Nur die Summen selbst durchgehen: Eine Laufsumme in jeder Zeile hätte sonst je Zeile ihren ganzen Bereich.
    totals = Columns([p for p in sheet.nodes if _is_total(sheet, p)])
    order = {p: i for i, p in enumerate(sheet.cells) if p in sheet.nodes and _is_total(sheet, p)}
    for pos, node in sheet.nodes.items():
        if node is None or not _is_total(sheet, pos):
            continue
        for area in sheet.ranges.get(pos, ()):
            r1, c1, r2, c2 = area
            # In der Reihenfolge von _inside: Zeile für Zeile, bei großen Bereichen wie die Zellen im Blatt stehen.
            small = (r2 - r1 + 1) * (c2 - c1 + 1) <= len(sheet.cells)
            for inner in sorted(totals.inside(area), key=lambda p: p if small else order[p]):
                # TEILERGEBNIS über TEILERGEBNIS: Excel lässt die inneren weg, so baut Daten → Teilergebnis die Summen.
                if inner == pos or node.value in ("SUBTOTAL", "AGGREGATE") and sheet.nested(inner):
                    continue
                for sub in sheet.ranges.get(inner, ()):
                    if r1 <= sub[0] and sub[2] <= r2 and c1 <= sub[1] and sub[3] <= c2:
                        found.append(Finding(
                            f"double:{address(pos)}",
                            T(f"Die Summe in {sheet.label(pos)} zählt {address(inner)} mit, und {address(inner)} ist selbst "
                              f"schon die Summe von {address(sub[:2])}:{address(sub[2:])}. Diese Zeilen zählen doppelt.",
                              f"The total in {sheet.label(pos)} includes {address(inner)}, and {address(inner)} is itself "
                              f"already the total of {address(sub[:2])}:{address(sub[2:])}. These rows count twice."),
                            "fehler", T("Nimm die Zwischensumme aus dem Bereich, oder summiere nur die Zwischensummen.",
                                        "Take the subtotal out of the range, or add up only the subtotals."),
                        ))
                        break
                else:
                    continue
                break
    return found[:3]


def _close(a: float, b: float, tolerance: float = 1e-6) -> bool:
    return abs(a - b) <= tolerance * max(1.0, abs(a), abs(b))


def _recalc(sheet: Sheet) -> list[Finding]:
    """Jede Formel mit den Werten aus Excel nachrechnen. Nur Zahl gegen Zahl zählt als Widerspruch.
    Viele Widersprüche in einer Spalte sind ein Hinweis (meist steht die Berechnung auf manuell)."""
    by_column: dict[int, list[Finding]] = {}
    for item in _recalc_cells(sheet):
        pos = _key_pos(item.key)
        by_column.setdefault(pos[1], []).append(item)
    found = []
    for col, items in by_column.items():
        if len(items) < 3:
            found += items
            continue
        first, last = _key_pos(items[0].key), _key_pos(items[-1].key)
        header = sheet.plan.headers.get(col)
        found.append(Finding(
            items[0].key,
            T(f"In {len(items)} Zellen{f' von {header}' if header else ''} ({address(first)} bis {address(last)}) zeigt Excel "
              f"etwas anderes, als die Formel ergibt. Zum Beispiel {items[0].text}",
              f"In {len(items)} cells{f' of {header}' if header else ''} ({address(first)} to {address(last)}), Excel shows "
              f"something other than what the formula gives. For example, {items[0].text}"),
            "fehler", items[0].tip,
        ))
    return found[:4]


def _key_pos(key: str) -> Pos:
    """„recalc:C7“ → (7, 3)."""
    match = re.fullmatch(r"[a-z]+:([A-Z]+)(\d+)", key)
    return int(match.group(2)), column_number(match.group(1))


def _recalc_cells(sheet: Sheet) -> list[Finding]:
    found = []
    calc, _ = sheet.evaluator()
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown:
            continue
        shown = sheet.inputs.get(pos)
        if not isinstance(shown, float) or isinstance(shown, bool):
            continue
        try:
            ours = calc.run(node)
        except formel.Unknown:
            continue
        if isinstance(ours, bool) or not isinstance(ours, float) or _close(ours, shown):
            continue
        text = sheet.shown.get(pos)
        if isinstance(text, str) and abs(ours - shown) < _precision(text):
            continue  # 1/11 als Uhrzeit zeigt 02:10: Die Sekunden sieht man nur nicht.
        if _fmt(ours) == _fmt(shown):
            continue  # Ein Rest unter dem, was man sieht (Uhrzeiten auf die Sekunde gelesen): „zeigt 7,5, ergibt 7,5“.
        if sheet.nested(pos):
            continue  # TEILERGEBNIS lässt weggefilterte und ausgeblendete Zeilen weg: Die kennt der Prüfer nicht.
        found.append(Finding(
            f"recalc:{address(pos)}",
            T(f"{sheet.label(pos)} zeigt {_shown(sheet, pos)}, nachgerechnet ergibt {sheet.readable(pos)} aber "
              f"{_date_text(ours, False) if isinstance(sheet.shown.get(pos), str) else _fmt(ours)}.",
              f"{sheet.label(pos)} shows {_shown(sheet, pos)}. Recalculated, {sheet.readable(pos)} gives "
              f"{_date_text(ours, False) if isinstance(sheet.shown.get(pos), str) else _fmt(ours)}."),
            "fehler",
            T("Steht die Berechnung auf manuell (Formeln → Berechnungsoptionen → Automatisch)? Oder steht eine Zahl als Text da?",
              "Is calculation set to manual (Formulas → Calculation Options → Automatic)? Or is a number stored as text?"),
        ))
    return found


def _precision(text: str) -> float:
    """Wie genau ein als Datum oder Uhrzeit gezeigter Wert noch zu sehen ist, in Tagen."""
    if re.search(r"\d:\d\d:\d\d", text):
        return 1 / 86400
    if re.search(r"\d:\d\d", text):
        return 1 / 1440
    return 1.0


def _shown(sheet: Sheet, pos: Pos) -> str:
    value = sheet.shown.get(pos)
    return value if isinstance(value, str) else _fmt(value)


def _empty_refs(sheet: Sheet) -> list[Finding]:
    """Eine Formel, die eine leere Zelle liest: Excel rechnet still mit 0."""
    found = []
    last = sheet.plan.last_row if sheet.plan.header_row is not None else None
    # Ein Bezug ohne $, der beim Ziehen mitrutscht, sagt zeilenmodell.py genauer, mit der Korrektur.
    drift = sheet.memo.get("zeilenmodell", {}).get("drift", {})
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or _asks_for_empty(node) or pos in drift:
            continue
        if last is not None and pos[0] > last and sheet.plan.in_table(sheet.plan.header_row + 1, pos[1]):
            continue  # Vorbereitete Zeilen unter der Tabelle: Da fehlen die Werte noch.
        empty = sorted(p for p in _counted(node) if sheet.shown.get(p) is None and p != pos)
        if empty:
            found.append(Finding(
                f"emptyref:{address(pos)}",
                T(f"{sheet.label(pos)} rechnet mit {', '.join(address(p) for p in empty[:3])}, "
                  f"{'die ist' if len(empty) == 1 else 'die sind'} aber leer (Excel nimmt dann 0).",
                  f"{sheet.label(pos)} calculates with {', '.join(address(p) for p in empty[:3])}, but "
                  f"{'it is' if len(empty) == 1 else 'they are'} empty (Excel then uses 0)."),
                "pruefen", T("Fehlt dort noch ein Wert, oder zeigt der Bezug auf die falsche Zelle?",
                             "Is a value still missing there, or does the reference point to the wrong cell?"),
            ))
    return found[:4]


_EMPTY_AWARE = {"IF", "IFS", "IFERROR", "IFNA", "ISBLANK", "COUNTA", "COUNTBLANK", "N", "LEN", "ISNUMBER"}
# Hier wird eine leere Zelle zu leerem Text, nicht zu 0: „Excel nimmt dann 0“ stimmt dort nicht.
_TEXTUAL = {"CONCAT", "CONCATENATE", "TEXTJOIN", "TEXT", "TRIM", "UPPER", "LOWER", "LEFT", "RIGHT", "MID", "SUBSTITUTE",
            "EXACT", "FIND", "SEARCH", "REPT"}


def _counted(node: Node):
    """Die einzelnen Zellen, mit denen node als Zahl rechnet (nicht die, die nur in Text eingehen)."""
    if node.kind == "ref":
        yield node.value
        return
    if (node.kind == "op" and node.value == "&") or (node.kind == "call" and node.value in _TEXTUAL):
        return
    for child in node.args:
        yield from _counted(child)


def _asks_for_empty(node: Node) -> bool:
    """WENN(C7="";"";…): Die Formel rechnet bewusst mit leeren Zellen."""
    if node.kind == "call" and (node.value in _EMPTY_AWARE or node.value.startswith("IS")):
        return True
    return any(_asks_for_empty(a) for a in node.args)


def _complex(sheet: Sheet) -> list[Finding]:
    """Formeln, die man kaum noch prüfen kann."""
    found = []
    for pos, node in sorted(sheet.nodes.items()):
        if node is None:
            continue
        deep = formel.depth(node, "IF")
        length = len(sheet.cells[pos].formula)
        if deep >= 4 or length > 250:
            what = T(f"{deep} WENN ineinander", f"{deep} nested IFs") if deep >= 4 else \
                T(f"{length} Zeichen lang", f"{length} characters long")
            found.append(Finding(
                f"complex:{address(pos)}", T(f"Die Formel in {sheet.label(pos)} ist schwer zu prüfen ({what}).",
                                             f"The formula in {sheet.label(pos)} is hard to check ({what})."), "pruefen",
                T("Teil sie in Hilfsspalten auf, oder nimm eine Nachschlagetabelle statt vieler WENN.",
                  "Split it into helper columns, or use a lookup table instead of many IFs."),
            ))
    return found[:3]


# --- Nachschlagen ---

_LOOKUP_CALLS = ("VLOOKUP", "HLOOKUP", "MATCH", "XLOOKUP", "XMATCH")


def _german(name: str) -> str:
    return FUNCTIONS.get(name, name)


def _called(name: str) -> str:
    """Der Name einer Funktion im Satz: SVERWEIS auf Deutsch, VLOOKUP auf Englisch (name: wie _german ihn gibt)."""
    return T(name, _ENGLISH_NAMES.get(name, name))


_ENGLISH_NAMES = {german: english for english, german in FUNCTIONS.items()}


def _area_text(evaluator, node: Node) -> str:
    sheet, (r1, c1, r2, c2) = evaluator.area_of(node)
    where = f"{address((r1, c1))}:{address((r2, c2))}" if (r1, c1) != (r2, c2) else address((r1, c1))
    return f"{sheet}!{where}" if sheet else where


def _keys_of(evaluator, call: Node) -> list | None:
    """Die Spalte (oder Zeile), in der gesucht wird."""
    try:
        if call.value == "VLOOKUP":
            return [row[0] for row in evaluator.grid(call.args[1]) if row]
        if call.value == "HLOOKUP":
            return list(evaluator.grid(call.args[1])[0])
        return evaluator._line(call.args[1])  # noqa: SLF001 - dieselbe Regel wie beim Rechnen
    except (formel.Unknown, formel._Fail):  # noqa: SLF001
        return None


def _approximate(evaluator, call: Node) -> tuple[bool, bool]:
    """(sucht ungefähr, ist es ausdrücklich so geschrieben)."""
    index = {"VLOOKUP": 3, "HLOOKUP": 3, "MATCH": 2}.get(call.value)
    if index is None:
        return False, False
    if index >= len(call.args) or call.args[index].kind == "empty":
        return True, False
    try:
        flag = evaluator.eval(call.args[index])
    except (formel.Unknown, formel._Fail):  # noqa: SLF001
        return False, True
    if call.value == "MATCH":
        return isinstance(flag, float) and flag != 0, True
    return bool(flag) if isinstance(flag, (bool, float)) else False, True


def _sorted(keys: list, kind: str, descending: bool = False) -> bool:
    same = [k.casefold() if isinstance(k, str) else k for k in keys if formel._kind(k) == kind]  # noqa: SLF001
    # Eine Kopfzeile über sortierten Schlüsseln stört Excel kaum.
    return same == sorted(same, reverse=descending) or same[1:] == sorted(same[1:], reverse=descending)


def _match_type(evaluator, call: Node) -> float:
    try:
        return float(evaluator.arg(call.args, 2, 1.0))
    except (formel.Unknown, formel._Fail, TypeError, ValueError):  # noqa: SLF001
        return 1.0


def _lookups(sheet: Sheet, explained: frozenset[Pos] | set[Pos] = frozenset()) -> list[Finding]:  # noqa: C901 - eine Prüfung je Abschnitt
    """Typische Fehler beim Nachschlagen: ungefähr statt genau, mitrutschende Tabelle, versteckte Lücken, doppelte Schlüssel.
    explained: Zellen, deren Ursache nachschlagen.py schon genau nennt."""
    calc, _ = sheet.evaluator()
    seen: dict[tuple, tuple] = {}
    approx: dict[tuple, list[Pos]] = {}
    nearest: dict[tuple, list[tuple[Pos, object, object]]] = {}
    hidden: dict[tuple, list[tuple[Pos, object, str, str]]] = {}
    twice: dict[tuple, tuple[Pos, object, list[int]]] = {}
    areas: dict[tuple, tuple | None] = {}
    current: dict = {}
    tables: dict[tuple, list[tuple[Pos, tuple]]] = {}
    # Welche Funktion auf welcher Tabelle: daraus baut die Korrektur die genaue Suche.
    looked: dict[tuple, tuple[str, tuple]] = {}
    # Alle Formeln, die dort ungefähr suchen, ohne es so hinzuschreiben: Die Korrektur ändert sie alle gleich.
    loose: dict[tuple, list[Pos]] = {}
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown:
            continue
        guarded = {id(inner): c.value for c in formel.calls(node) if c.value in ("IFERROR", "IFNA") and c.args
                   for inner in formel.calls(c.args[0])}
        for call in formel.calls(node):
            if call.value not in _LOOKUP_CALLS or len(call.args) < 2:
                continue
            try:
                table = calc.area_of(call.args[1])
                where = _area_text(calc, call.args[1])
            except formel.Unknown:
                continue
            name = _german(call.value)
            looked.setdefault((name, where), (call.value, table))
            if pos not in explained:
                tables.setdefault((pos[1], name), []).append((pos, table))
            cached = seen.get((call.value, table))
            if cached is None:
                keys = _keys_of(calc, call)
                index: dict = {}
                for i, k in enumerate(keys or []):
                    index.setdefault(formel._norm(k), []).append(i)  # noqa: SLF001
                cached = seen[(call.value, table)] = (keys, index, {})
            keys, index, order = cached
            if keys is None:
                continue
            try:
                key = calc.eval(call.args[0])
            except (formel.Unknown, formel._Fail):  # noqa: SLF001
                continue
            key = calc_value(key)
            kind = formel._kind(key)  # noqa: SLF001
            rough, written = _approximate(calc, call)
            if rough and not written and pos not in loose.setdefault((name, where), []):
                loose[(name, where)].append(pos)
            if pos in explained and not (not rough and len(index.get(formel._norm(key), [])) > 1):  # noqa: SLF001
                continue
            descending = call.value == "MATCH" and rough and _match_type(calc, call) < 0
            if rough and kind and (kind, descending) not in order:
                order[(kind, descending)] = _sorted(keys, kind, descending)
            if rough and kind and not order[(kind, descending)]:
                approx.setdefault((name, where), []).append(pos)
                continue
            exact = index.get(formel._norm(key), []) if kind else []  # noqa: SLF001
            if rough and not written and kind == "s" and not exact:
                below = [k for k in keys if isinstance(k, str) and k.casefold() <= key.casefold()]
                if below:
                    nearest.setdefault((name, where), []).append((pos, key, below[-1]))
                continue
            if not rough and not exact and id(call) in guarded and kind:
                hidden.setdefault((name, where), []).append((pos, key, _why_missing(key, keys), guarded[id(call)]))
            if not rough and len(exact) > 1 and current.get((name, where)) is None:
                r1 = table[1][0] if call.value != "HLOOKUP" else table[1][1]
                marked = _current_row(sheet, calc, call, table[1], [r1 + i for i in exact]) \
                    if call.value == "VLOOKUP" else None
                # Eine zweite Spalte derselben Suche, in der der gültige Eintrag etwas anderes liefert, sagt mehr.
                if (name, where) not in twice or marked is not None:
                    twice[(name, where)] = (pos, key, [r1 + i for i in exact])
                    # Auf diesem Blatt: Dann kann Einsetzen die Suche von unten schreiben.
                    areas[(name, where)] = table[1] if call.args[1].kind == "range" else None
                    current[(name, where)] = marked
    from . import einsetzen

    found: list[Finding] = []
    for (name, where), places in list(approx.items())[:2]:
        many = T(f" (und {len(places) - 1} weitere)", f" (and {len(places) - 1} more)") if len(places) > 1 else ""
        found.append(Finding(
            f"approx:{address(places[0])}",
            T(f"{name} in {sheet.label(places[0])}{many} sucht ungefähr, aber {where} ist nicht aufsteigend sortiert. "
              f"Dann liefert Excel leicht einen falschen Wert, ohne Fehlermeldung.",
              f"{_called(name)} in {sheet.label(places[0])}{many} uses an approximate match, but {where} is not sorted in "
              f"ascending order. Excel can then easily return a wrong value without any error."),
            "fehler", T(f"Für eine genaue Suche schreib FALSCH (bei VERGLEICH 0) als letztes Argument, oder sortiere {where} aufsteigend.",
                        f"For an exact match, use FALSE (0 for MATCH) as the last argument, or sort {where} in ascending order."),
            # Genau suchen nur, wenn dann jeder Schlüssel gefunden wird; sonst sind es Stufen (Tarife), und sortieren hilft.
            fix=einsetzen.exact(sheet, places, *looked[(name, where)], errors=False),
        ))
    for (name, where), places in list(nearest.items())[:2]:
        pos, key, taken = places[0]
        many = T(f" Das passiert in {len(places)} Formeln.", f" This happens in {len(places)} formulas.") if len(places) > 1 else ""
        found.append(Finding(
            f"nearest:{address(pos)}",
            T(f"{name} in {sheet.label(pos)} findet „{_fmt(key)}“ nicht in {where} und nimmt still den Eintrag „{_fmt(taken)}“, "
              f"weil das letzte Argument fehlt (ungefähre Suche).{many}",
              f"{_called(name)} in {sheet.label(pos)} does not find “{_fmt(key)}” in {where} and silently takes the entry "
              f"“{_fmt(taken)}”, because the last argument is missing (approximate match).{many}"),
            "fehler", T("Schreib FALSCH als letztes Argument, dann zeigt Excel #NV, wenn etwas fehlt.",
                        "Use FALSE as the last argument; then Excel shows #N/A when something is missing."),
            fix=einsetzen.exact(sheet, loose.get((name, where)) or [p for p, _, _ in places], *looked[(name, where)],
                                errors=True),
        ))
    for (name, where), places in list(hidden.items())[:2]:
        pos, key, why, guard = places[0]
        guard = _german(guard)
        if len(places) == 1:
            text = T(f"In {sheet.label(pos)} findet {name} „{_fmt(key)}“ nicht in {where}{why}, und {guard} setzt still einen Ersatzwert ein.",
                     f"In {sheet.label(pos)}, {_called(name)} does not find “{_fmt(key)}” in {where}{why}, and {_called(guard)} "
                     f"silently puts in a substitute value.")
        else:
            listed = [T(f"„{_fmt(k)}“ in {address(p)}{w}", f"“{_fmt(k)}” in {address(p)}{w}") for p, k, w, _ in places[:4]]
            more = T(f" und {len(places) - 4} weitere", f" and {len(places) - 4} more") if len(places) > 4 else ""
            text = T(f"In {len(places)} Formeln findet {name} den gesuchten Wert nicht in {where}, und {guard} setzt still "
                     f"einen Ersatzwert ein: {', '.join(listed)}{more}.",
                     f"In {len(places)} formulas, {_called(name)} does not find the value it looks for in {where}, and "
                     f"{_called(guard)} silently puts in a substitute value: {', '.join(listed)}{more}.")
        tip = T("Prüfe die Schreibweise der Schlüssel.", "Check how the keys are spelled.")
        if guard == _german("IFERROR"):
            tip += T(f" Mit {_german('IFNA')} statt {guard} bleiben andere Fehler sichtbar.",
                     " With IFNA instead of IFERROR, other errors stay visible.")
        found.append(Finding(f"hidden:{address(pos)}", text, "pruefen", tip))
    for (name, where), (pos, key, rows) in list(twice.items())[:2]:
        # In einer Tabelle mit gültigem Eintrag nennt die Zelle links (Status „Gebucht“) die Formel nicht.
        shown = sheet.label(pos) if current.get((name, where)) is None else address(pos)
        text = T(f"„{_fmt(key)}“ steht {len(rows)}-mal in {where} (Zeilen {', '.join(map(str, rows[:4]))}). {name} in "
                 f"{shown} nimmt immer den ersten Eintrag.",
                 f"“{_fmt(key)}” appears {len(rows)} times in {where} (rows {', '.join(map(str, rows[:4]))}). "
                 f"{_called(name)} in {shown} always takes the first entry.")
        level, tip = "pruefen", T("Sind es wirklich zwei verschiedene? Sonst einen Eintrag löschen oder eindeutige Schlüssel nehmen.",
                                  "Are they really two different ones? If not, delete one entry or use unique keys.")
        fix = None
        marked = current.get((name, where))
        if marked is not None:
            head, row, shown, old, new, mark_col = marked
            text += T(f" Gültig ist aber Zeile {row} ({head} „{shown}“), und dort steht {_fmt(new)} statt {_fmt(old)}. "
                      f"So zeigt {address(pos)} den alten Stand.",
                      f" But row {row} is the valid one ({head} “{shown}”), and it has {_fmt(new)} instead of {_fmt(old)}. "
                      f"So {address(pos)} shows the old value.")
            level = "fehler"
            fix = _from_below(sheet, pos, areas[(name, where)], mark_col) if areas.get((name, where)) else None
            if fix is not None:
                tip = T(f"Bei jedem doppelten Eintrag steht der gültige ({head} „{shown}“) unten: XVERWEIS von unten "
                        f"(Suchmodus -1) findet ihn. Korrigiert: {fix['german']}",
                        f"For every duplicate entry, the valid one ({head} “{shown}”) is the lower one: XLOOKUP from the "
                        f"bottom (search mode -1) finds it. Fixed: {fix['german']}")
            else:
                tip = T(f"Such den gültigen Eintrag: mit XVERWEIS von unten (Suchmodus -1), wenn der neuere unten steht, "
                        f"oder mit einer Bedingung auf {head}.",
                        f"Look up the valid entry: with XLOOKUP from the bottom (search mode -1) if the newer one is "
                        f"lower down, or with a condition on {head}.")
        found.append(Finding(f"twice:{address(pos)}", text, level, tip, fix=fix))
    found += _sliding_tables(sheet, tables)
    return found


_CURRENT = re.compile(r"^(aktuell|aktiv|gültig|gueltig|current|active|valid|latest|neuester?)\b", re.I)
_YES = {"ja", "yes", "x", "wahr", "true", "1", "aktiv", "aktuell", "j", "y"}


def _current_row(sheet: Sheet, calc, call, area, rows: list[int]):
    """Eine Spalte neben der Suchtabelle sagt, welcher der doppelten Einträge gilt (Aktuell = Ja). Liefert (Kopf,
    gültige Zeile, ihr Wert dort, was SVERWEIS liefert, was dort stünde), wenn der gültige nicht der erste ist und
    etwas anderes liefert. Sonst None."""
    r1, c1, _, c2 = area
    head_row = r1 - 1
    for c in range(c1, c2 + 12):
        cell = sheet.cells.get((head_row, c))
        if cell is None or cell.formula or not isinstance(cell.value, str) or not _CURRENT.search(cell.value.strip()):
            continue
        marks = [r for r in rows if str(sheet.inputs.get((r, c), sheet.cells.get((r, c)) and sheet.cells[(r, c)].value)
                                      ).strip().casefold() in _YES]
        if not marks:
            continue  # „Gültig ab“ mit Datum ist keine Markierung
        if len(marks) != 1 or marks[0] == rows[0]:
            return None
        try:
            index = int(calc_value(calc.eval(call.args[2])))
        except (formel.Unknown, formel._Fail, TypeError, ValueError, IndexError):  # noqa: SLF001
            return None
        col = c1 + index - 1
        old, new = sheet.inputs.get((rows[0], col)), sheet.inputs.get((marks[0], col))
        if old == new or new is None:
            return None
        shown = sheet.inputs.get((marks[0], c))
        return cell.value.strip(), marks[0], _fmt(shown), old, new, c
    return None


def _from_below(sheet: Sheet, pos: Pos, area, mark_col: int) -> dict | None:
    """Die Korrektur zu SVERWEIS, der bei doppelten Schlüsseln den alten Eintrag nimmt: XVERWEIS von unten
    (Suchmodus -1), in der ganzen heruntergezogenen Spalte. Nur wenn bei jedem doppelten Schlüssel der gültige
    Eintrag (Aktuell = Ja) der unterste ist; sonst fände auch XVERWEIS von unten den falschen."""
    from . import einsetzen
    from .auswertung import _block, _filled_down
    from .bereiche import _kept, _text

    r1, c1, r2, c2 = area
    groups: dict = {}
    for r in range(r1, r2 + 1):
        key = sheet.inputs.get((r, c1))
        if key is not None and str(key).strip():
            groups.setdefault(formel._norm(key), []).append(r)  # noqa: SLF001
    for rows in groups.values():
        if len(rows) < 2:
            continue
        marks = [r for r in rows if str(sheet.inputs.get((r, mark_col))).strip().casefold() in _YES]
        if marks != [rows[-1]]:
            return None
    places = _block(sheet, pos) if _filled_down(sheet, pos) else [pos]
    changes = {}
    for place in places:
        formula = (sheet.cells[place].formula or "") if place in sheet.cells else ""
        try:
            node = formel.parse(formula)
        except (ValueError, IndexError, KeyError, formel.Unknown):
            return None
        kept = _kept(formula)
        hit = []

        def walk(item):
            if not isinstance(item, formel.Node):
                return item
            item = formel.Node(item.kind, item.value, tuple(walk(a) for a in item.args))
            if item.kind != "call" or item.value != "VLOOKUP" or len(item.args) < 3:
                return item
            table, index = item.args[1], item.args[2]
            if table.kind != "range" or table.value != area or index.kind != "num":
                return item
            col = c1 + int(index.value) - 1
            if not c1 <= col <= c2:
                return item
            shown = kept.get(("range", area), "")
            m = re.fullmatch(r"(\$?)[A-Z]+(\$?)(\d+):(\$?)[A-Z]+(\$?)(\d+)", shown)
            for c in (c1, col):
                new_area = (r1, c, r2, c)
                if m:
                    letters = column_letters(c)
                    kept[("range", new_area)] = f"{m[1]}{letters}{m[2]}{m[3]}:{m[4]}{letters}{m[5]}{m[6]}"
            hit.append(True)
            return formel.Node("call", "XLOOKUP", (
                item.args[0], formel.Node("range", (r1, c1, r2, c1)), formel.Node("range", (r1, col, r2, col)),
                formel.Node("empty", None), formel.Node("num", 0.0), formel.Node("num", -1.0)))

        text = _text(walk(node), kept)
        if not hit or text is None:
            return None
        changes[place] = text
    return einsetzen.formulas(sheet, changes)


def _why_missing(key, keys: list) -> str:
    """Warum ein Schlüssel fehlt, wenn es fast passt: Leerzeichen oder Zahl als Text."""
    if isinstance(key, str):
        plain = key.strip().casefold()
        if any(isinstance(k, str) and k.strip().casefold() == plain for k in keys):
            return T(" (dort steht es mit Leerzeichen davor oder dahinter)", " (there it has a space before or after it)")
        number = _as_number(key)
        if number is not None and number in keys:
            return T(" (dort steht es als Zahl, hier als Text)", " (there it is a number, here it is text)")
    if isinstance(key, float) and any(isinstance(k, str) and _as_number(k) == key for k in keys):
        return T(" (dort steht es als Text, hier als Zahl)", " (there it is text, here it is a number)")
    return ""


def _sliding_tables(sheet: Sheet, tables: dict[tuple, list[tuple[Pos, tuple]]]) -> list[Finding]:
    """Eine Nachschlagetabelle ohne $, die beim Herunterkopieren mitrutscht."""
    found = []
    for (_, name), uses in tables.items():
        if len(uses) < 3:
            continue
        (first, (s1, a1)), (last, (s2, a2)) = uses[0], uses[-1]
        rows = last[0] - first[0]
        if s1 != s2 or rows <= 0 or a2[0] - a1[0] != rows or a2[2] - a1[2] != rows:
            continue
        found.append(Finding(
            f"sliding:{address(first)}",
            T(f"Die Tabelle, in der {name} sucht, rutscht beim Kopieren mit: {address(first)} sucht in "
              f"{address(a1[:2])}:{address(a1[2:])}, {address(last)} schon in {address(a2[:2])}:{address(a2[2:])}. "
              f"Weiter unten fehlen so die oberen Einträge.",
              f"The table that {_called(name)} searches moves along when copied: {address(first)} searches "
              f"{address(a1[:2])}:{address(a1[2:])}, {address(last)} already {address(a2[:2])}:{address(a2[2:])}. "
              f"So further down, the top entries are missing."),
            "fehler", T("Setz $ vor Zeilen und Spalten der Tabelle (Taste F4), zum Beispiel $A$2:$C$50.",
                        "Put $ before the rows and columns of the table (F4 key), for example $A$2:$C$50."),
        ))
        if len(found) >= 2:
            break
    return found


# --- Daten ---


def _data(sheet: Sheet) -> list[Finding]:
    from .auswertung import is_id, is_naming

    plan = sheet.plan
    found: list[Finding] = []
    if plan.header_row is None:
        return found
    rows = sorted(r for r in plan.names if plan.header_row < r <= plan.last_row) or list(range(plan.header_row + 1, plan.last_row + 1))
    for col, header in sorted(plan.headers.items()):
        # In einem Zeilenmodell (Finanzplan) steht in einer Spalte jede Größe einmal: Umsatz, Marge in %, Tilgung. Ein
        # Ausreißer gegen die übrigen Zeilen sagt dort nichts.
        if col == plan.name_column or plan.across:
            continue
        typed = {r: sheet.shown.get((r, col)) for r in rows if (r, col) in sheet.cells and sheet.cells[(r, col)].formula is None}
        numbers = {r: v for r, v in typed.items() if isinstance(v, float)}
        if not is_naming(header):
            found += _outliers(sheet, col, header, numbers)
        if any(word in header.lower() for word in _PERCENT) and not any(w in header.lower() for w in _CHANGES):
            found += _shares(sheet, col, header, numbers)
        # Steht dort noch einmal der Spaltenname, beginnt darunter eine zweite Tabelle.
        texts = [r for r, v in typed.items() if isinstance(v, str) and v.strip().lower() != header.strip().lower()
                 and not _formula_text(sheet, v)]
        if len(numbers) >= 3 and texts and len(texts) <= max(2, len(numbers) // 5):
            found.append(Finding(
                f"text:{header}",
                T(f"In {header} stehen Zahlen, aber bei {', '.join(address((r, col)) for r in texts[:3])} steht Text "
                  f"({', '.join(repr(sheet.shown[(r, col)]) for r in texts[:3])}). SUMME und MITTELWERT lassen das still weg.",
                  f"{header} contains numbers, but {', '.join(address((r, col)) for r in texts[:3])} "
                  f"{'contains' if len(texts[:3]) == 1 else 'contain'} text "
                  f"({', '.join(repr(sheet.shown[(r, col)]) for r in texts[:3])}). SUM and AVERAGE silently skip it."),
                "pruefen", T("Steht dort eine Zahl als Text, ein Tippfehler oder eine Notiz?",
                             "Is that a number stored as text, a typo or a note?"),
            ))
    # Bei Fall- und Patientennummern sagt auswertung.py es genauer.
    if plan.name_column is not None and not is_id(plan.headers.get(plan.name_column, "")):
        listed = set(rows)
        names = Counter(v.strip().lower() for r, v in plan.names.items() if r in listed and v and v.strip())
        # Ein Datum steht in einer Liste von Buchungen oft mehrmals: Das ist kein doppelter Name.
        if sum(1 for n in names.elements() if formel.serial_of(n) is not None) * 2 > sum(names.values()):
            names = Counter()
        # Wiederholt sich fast jeder Name, ist es eine Spalte mit Gruppen (Station, Code), keine Liste von Namen.
        unique = len(names) >= 0.75 * sum(names.values())
        twice = [n for n, k in names.items() if k > 1] if unique else []
        for name in twice[:2]:
            where = [r for r, v in plan.names.items() if v.strip().lower() == name]
            # „Jahresüberschuss“ noch einmal im Cashflow, mit =C29: übernommen, nicht doppelt.
            if all(links(sheet, r, where[0]) for r in where[1:]):
                continue
            if not _same_entry(sheet, where):
                continue
            found.append(Finding(
                f"dupe:{address((where[1], plan.name_column))}",
                T(f"„{plan.names[where[0]]}“ steht {len(where)}-mal in {plan.headers.get(plan.name_column, 'der Namensspalte')} "
                  f"(Zeilen {', '.join(str(r) for r in where[:5])}).",
                  f"“{plan.names[where[0]]}” appears {len(where)} times in "
                  f"{plan.headers.get(plan.name_column, 'the name column')} (rows {', '.join(str(r) for r in where[:5])})."),
                "pruefen", T("Wird es doppelt gezählt, oder sind es wirklich zwei verschiedene?",
                             "Is it counted twice, or are they really two different ones?"),
            ))
    return found


def _same_entry(sheet: Sheet, rows: list[int]) -> bool:
    """Sieht der doppelte Name nach demselben Eintrag zweimal aus? Nicht, wenn er eine Zahl ist (700 in „Opening“),
    wenn eine andere Textspalte die Zeilen unterscheidet („Food | City Bank“ und „Food | J.P Morgan“: eine Kategorie
    mit zwei Konten), oder wenn die spätere Zeile die Tabelle darüber nachschlägt („Pork“ unter der Tabelle, daneben
    =SVERWEIS(D7;D2:E5;2;FALSCH)): Dann ist sie eine Abfrage, kein zweiter Eintrag."""
    plan = sheet.plan
    if _as_number(plan.names[rows[0]].strip()) is not None:
        return False
    for col in plan.headers:
        if col == plan.name_column:
            continue
        texts = {r: sheet.shown.get((r, col)) for r in rows if (r, col) not in sheet.nodes}
        texts = {r: v.strip().casefold() for r, v in texts.items() if isinstance(v, str) and v.strip()}
        if len(set(texts.values())) > 1:
            return False
    first = rows[0]
    for later in rows[1:]:
        for col in plan.headers:
            for area in sheet.ranges.get((later, col), ()):
                if area[0] <= first <= area[2] and area[1] <= plan.name_column <= area[3]:
                    return False
    return True


def links(sheet: Sheet, later: int, first: int) -> bool:
    """Übernimmt Zeile later nur Zeile first (=C29 in jeder Spalte)? Dann ist sie mit Absicht gleich. Eine feste
    Zahl dazwischen, die dasselbe zeigt (ein Jahr als Werte eingefügt), ändert daran nichts: Das sagt zeilenmodell.py."""
    plan = sheet.plan
    found = 0
    for col in plan.headers:
        cell = sheet.cells.get((later, col))
        if col == plan.name_column or cell is None or (not cell.formula and not cell.value):
            continue
        before = sheet.precedents.get((later, col))
        if not cell.formula and sheet.inputs.get((later, col)) == sheet.inputs.get((first, col)):
            continue
        # Auch über Zwischenzeilen: =E93 in der Kapitalflussrechnung, =E33 in E93, die Annahme in E33.
        hops = 0
        while cell is not None and cell.formula and isinstance(before, set) and len(before) == 1 \
                and next(iter(before)) != (first, col) and next(iter(before))[1] == col and hops < 3:
            middle = next(iter(before))
            cell, before, hops = sheet.cells.get(middle), sheet.precedents.get(middle), hops + 1
            if cell is not None and not cell.formula:
                before = None
        if cell is None or not cell.formula or not isinstance(before, set) or before != {(first, col)}:
            return False
        found += 1
    return found >= 2 or (found == 1 and all(
        sheet.cells.get((later, c)) is None or sheet.cells[(later, c)].formula or not sheet.cells[(later, c)].value
        for c in plan.headers if c != plan.name_column))


_PLAIN = re.compile(r"-?\d+(?:,\d+)?")


# Zeilen, Spalten und Prüfungen, in denen jeder Rest zählt: Eine Abweichung von 0,40 darf nicht als 0 dastehen.
_DIFFERENCE = re.compile(r"abweich|differenz|\bdiff\b|delta|\brest\b|saldo|kontroll|\bcheck|prüf|fehlbetrag|unterschied|"
                         r"difference|variance|mismatch|\bgap\b|écart|differenza|scarto", re.IGNORECASE)


def _shown_as_zero(sheet: Sheet) -> list[Finding]:
    """0,4 in einer Zelle mit dem Format #,##0: Excel zeigt 0. Wo es um eine Abweichung geht (Bank − Modell, eine
    Kontrolle), sieht es dann aus, als stimme alles. Einsetzen gibt der Zelle dasselbe Format mit zwei
    Nachkommastellen."""
    from . import einsetzen

    found = []
    for pos, number in sorted(getattr(sheet.cells, "rounded", {}).items()):
        cell = sheet.cells.get(pos)
        if cell is None or len(found) >= 3:
            continue
        # Der Name vorne in der Zeile („Abweichung Bank − Modell“), auch wenn leere Spalten dazwischen liegen.
        left = next((c.value.strip() for c in (sheet.cells.get((pos[0], k)) for k in range(pos[1] - 1, 0, -1))
                     if c is not None and not c.formula and isinstance(c.value, str) and re.search(r"[^\W\d_]", c.value)), "")
        name = f"{left} {sheet.label(pos)}"
        readers = [p for p in sheet.dependents.get(pos, ()) if sheet.nodes.get(p) is not None
                   and any(c.value in ("IF", "IFS", "ROUND") for c in formel.calls(sheet.nodes[p]))]
        if not _DIFFERENCE.search(name) and not readers:
            continue
        fix = (einsetzen.formulas(sheet, {pos: cell.formula}) if cell.formula
               else einsetzen.values(sheet, {pos: number}, None, _fmt(number)))
        if fix is not None:
            fix["decimals"] = [address(pos)]
            fix["german"] = T("Zahlenformat mit zwei Nachkommastellen", "Number format with two decimal places")
        found.append(Finding(
            f"zeroshown:{address(pos)}",
            T(f"{f'{address(pos)} ({left})' if left else sheet.label(pos)} ist {_fmt(number)}, zeigt aber 0: Das "
              "Zahlenformat hat keine Nachkommastellen. So "
              "sieht es aus, als gäbe es keine Abweichung.",
              f"{f'{address(pos)} ({left})' if left else sheet.label(pos)} is {_fmt(number)} but shows 0: the number "
              "format has no decimal places. So it looks as if there were no difference."),
            "fehler", T(f"Gib {address(pos)} ein Zahlenformat mit Nachkommastellen (#.##0,00), dann steht "
                        f"{_fmt(number)} da.",
                        f"Give {address(pos)} a number format with decimal places (#,##0.00); then it shows "
                        f"{_fmt(number)}."), fix=fix))
    return found


def _clipped(sheet: Sheet) -> list[Finding]:
    """8 Stunden in einer Zelle mit dem Format hh:mm: Excel zeigt 00:00, als Datum den 07.01.1900. Gerechnet wird mit
    8, nur sieht man es nicht. Gemeldet, wenn die übrigen Zellen der Spalte die Zahl zeigen; Einsetzen gibt der Zelle
    deren Format."""
    from . import einsetzen

    found = []
    clipped = getattr(sheet.cells, "clipped", {})
    for pos, number in sorted(clipped.items())[:5]:
        others = [p for p in sorted(sheet.cells) if p[1] == pos[1] and p != pos and p not in clipped
                  and sheet.plan.header_row < p[0] <= sheet.plan.last_row and _PLAIN.fullmatch(sheet.cells[p].value.strip())]
        if len(others) < 2:
            continue
        like = min(others, key=lambda p: abs(p[0] - pos[0]))
        cell = sheet.cells[pos]
        fix = (einsetzen.formulas(sheet, {pos: cell.formula}) if cell.formula
               else einsetzen.values(sheet, {pos: number}, like, _fmt(number)))
        if fix is not None:
            fix["format"] = {address(pos): address(like)}
            fix["german"] = T(f"Zahlenformat wie {address(like)}", f"Number format like {address(like)}")
        found.append(Finding(
            f"clipped:{address(pos)}",
            T(f"{sheet.label(pos)} enthält {_fmt(number)}, ist aber als Uhrzeit oder Datum formatiert. Excel zeigt darum "
              f"{'00:00' if number == int(number) else 'eine Uhrzeit'} oder ein Datum im Januar 1900 statt {_fmt(number)}; "
              f"{address(like)} daneben zeigt die Zahl. Gerechnet wird trotzdem mit {_fmt(number)}.",
              f"{sheet.label(pos)} contains {_fmt(number)} but is formatted as a time or date. So Excel shows "
              f"{'00:00' if number == int(number) else 'a time'} or a date in January 1900 instead of {_fmt(number)}; "
              f"{address(like)} next to it shows the number. Calculations still use {_fmt(number)}."),
            "fehler",
            T(f"Korrigiert: Zahlenformat wie in {address(like)}, dann steht {_fmt(number)} da.",
              f"Fixed: number format like in {address(like)}; then it shows {_fmt(number)}."),
            fix=fix,
        ))
    return found


def _formula_text(sheet: Sheet, value: str) -> bool:
    """=SUMME(B2:B4) als Text: Das sagt formeltext.py genauer."""
    from .formeltext import _formula

    return value.lstrip().startswith("=") and _formula(sheet, value.strip()[1:]) is not None


def _outliers(sheet: Sheet, col: int, header: str, numbers: dict[int, float], where: str = "") -> list[Finding]:
    """where: mit welchen Werten verglichen wird („in Wert, wo Parameter „Kalium“ ist“, gruppen.py); sonst die Spalte."""
    where = where or f"in {header}"
    if len(numbers) < OUTLIER_MIN:
        return []
    from .auswertung import codes_in

    # Codes wie 999 für „fehlt“ meldet auswertung.py eigens.
    codes = codes_in(list(numbers.values()))
    numbers = {r: v for r, v in numbers.items() if v not in codes}
    values = list(numbers.values())
    if len(values) < OUTLIER_MIN:
        return []
    median = statistics.median(values)
    # Positive Messwerte auf der Log-Skala: Ein Tippfehler verschiebt um einen Faktor (720 oder 7,2 statt 72), und das
    # lange obere Ende von Kosten oder Laborwerten ist dort kein „weit weg“. Nicht bei Noten und Punkten (1 bis 5).
    # Punkte sind ganze Zahlen auf einer kurzen Skala; 10, 25, 50 und 200 Stück sind Mengen, keine Punkte.
    ordered = sorted(values)
    scores = all(v == int(v) for v in values) and len(set(values)) <= 10 and \
        ordered[-1 - len(ordered) // 10] - ordered[len(ordered) // 10] <= 10
    scale = {r: math.log10(v) for r, v in numbers.items()} if min(values) > 0 and not scores else numbers
    center = statistics.median(scale.values())
    mad = statistics.median(abs(v - center) for v in scale.values())
    if mad == 0:
        return []
    hits = _far(scale, center, mad)
    from .auswertung import CODES

    if scale is not numbers:
        # Auch auf der gewöhnlichen Skala nicht nah: Ein kleiner echter Wert nahe 0 ist auf der Log-Skala weit weg.
        raw = statistics.median(abs(v - median) for v in values)
        hits = [(z, r) for z, r in hits if raw == 0 or abs(numbers[r] - median) / (1.4826 * raw) > OUTLIER_Z / 2]
        # Ein Code wie 999 ist kein Messwert: Unter Werten von 4 bis 122 liegt er auf der Log-Skala nicht weit weg
        # (nur gut 8-mal das Größte), auf der gewöhnlichen schon. Für ihn gilt darum die gewöhnliche Skala.
        if raw:
            taken = {r for _, r in hits}
            hits += [(z, r) for z, r in _far(numbers, median, raw)
                     if r not in taken and numbers[r] in CODES and abs(numbers[r]) >= 99]
            hits.sort(reverse=True)
    found = []

    for z, r in hits[:3]:
        # Ein einzelnes 999 oder -99 ist meist ein Code für „fehlt“, kein Tippfehler: Das sagt der Hinweis gleich.
        code = numbers[r] in CODES and abs(numbers[r]) >= 99
        found.append(_scaled(sheet, col, where, numbers, r) or Finding(
            f"outlier:{address((r, col))}",
            T(f"{sheet.label((r, col))} = {_fmt(numbers[r])} liegt weit weg von den übrigen Werten {where} "
              f"(Median {_fmt(median)}, {z:.0f}-mal die übliche Streuung)."
              + (f" {_fmt(numbers[r])} sieht nach einem Code für „fehlt“ oder „unbekannt“ aus, aber MITTELWERT und SUMME "
                 "rechnen damit." if code else ""),
              f"{sheet.label((r, col))} = {_fmt(numbers[r])} is far away from the other values {where} "
              f"(median {_fmt(median)}; {z:.0f} times the usual spread)."
              + (f" {_fmt(numbers[r])} looks like a code for “missing” or “unknown”, but AVERAGE and SUM "
                 "calculate with it." if code else "")),
            "pruefen",
            T("Ist es ein Code, lass die Zelle leer oder nimm den Wert vor dem Rechnen heraus "
              f"(z. B. MITTELWERTWENN(…; \"<>{_fmt(numbers[r])}\")). Sonst: ein Tippfehler?",
              "If it is a code, leave the cell empty or take the value out before calculating "
              f"(e.g. AVERAGEIF(…, \"<>{_fmt(numbers[r])}\")). Otherwise: a typo?") if code else
            T("Ein Tippfehler (Komma, eine Null zu viel) oder ein echter Sonderfall?",
              "A typo (decimal point, one zero too many) or a genuine special case?"),
        ))
    return found


def _far(scale: dict[int, float], center: float, mad: float) -> list[tuple[float, int]]:
    """Die Werte hinter einer Lücke, die weiter als OUTLIER_Z-mal die übliche Streuung weg liegen, die weitesten zuerst."""
    apart = _apart(scale, center)
    far = sorted(((abs(v - center) / (1.4826 * mad), r) for r, v in scale.items() if r in apart), reverse=True)
    hits = [(z, r) for z, r in far if z > OUTLIER_Z]
    # Liegt hinter der Lücke ein Wert, der nicht weit weg ist (91 600 nach 40 500 und 46 600 unter Budgets bis
    # 243 600), trennt die Lücke keine Ausreißer ab: Sie liegt mitten in der Streuung.
    near = {r for z, r in far if z <= OUTLIER_Z}
    for side in (1, -1):
        group = {r for r in apart if (scale[r] - center) * side > 0}
        if group & near:
            hits = [(z, r) for z, r in hits if r not in group]
    return hits


# Um so viel verschiebt ein verrutschtes Komma oder eine andere Einheit (m statt cm, g statt kg).
_FACTORS = (10.0, 100.0, 1000.0, 0.1, 0.01, 0.001)


def _scaled(sheet: Sheet, col: int, where: str, numbers: dict[int, float], r: int) -> Finding | None:
    """1,74 in Größe (cm), wo die übrigen Werte zwischen 160 und 192 liegen: Mal 100 passt es dazu. Nur, wenn genau
    ein Faktor 10, 100 oder 1000 den Wert mitten unter die übrigen bringt (die mittleren 80 %), sonst bleibt es beim
    Ausreißer. Mit der Zahl zum Einsetzen und dem, was die Formeln damit anders rechnen."""
    from . import einsetzen

    value = numbers[r]
    others = sorted(v for k, v in numbers.items() if k != r)
    if value == 0 or len(others) < OUTLIER_MIN:
        return None
    low, high = others[len(others) // 10], others[-1 - len(others) // 10]
    # Eine weite Spalte (Beträge von 50 bis 5000): Welcher Faktor gemeint ist, ist nicht zu sagen.
    if low <= 0 or high >= 10 * low:
        return None
    fits = [f for f in _FACTORS if low <= value * f <= high]
    if len(fits) != 1:
        return None
    factor, pos = fits[0], (r, col)
    right = float(f"{value * factor:.12g}")
    how = T(f"Mal {_fmt(factor)}", f"Times {_fmt(factor)}") if factor > 1 else \
        T(f"Durch {_fmt(1 / factor)} geteilt", f"Divided by {_fmt(1 / factor)}")
    effect = ""
    changed = changed_by(sheet, sheet.dependents.get(pos, ()), {pos: right}, near=True)
    if changed:
        shown = [T(f"{sheet.label(p)} {said(b)} statt {said(a)}", f"{sheet.label(p)} gives {said(b)} instead of {said(a)}")
                 for p, b, a in changed[:3]]
        more = more_formulas(len(changed) - 3) if len(changed) > 3 else ""
        listed = shown[0] if len(shown) == 1 else T(", ", "; ").join(shown[:-1]) + T(f" und {shown[-1]}", f" and {shown[-1]}")
        effect = T(f" Darum ergibt {', '.join(shown)}{more}.", f" So {'; '.join(shown)}{more}.") if more else \
            T(f" Darum ergibt {listed}.", f" So {listed}.")
    like = next(((k, col) for k in numbers if k != r), None)
    # Allein sagt die Spalte nicht, ob es ein Tippfehler ist (Kreatinin 640 gibt es auch). Zeigt eine Formel damit
    # etwas Unmögliches (folgen.py), ist es sicher: Dann gilt der zweite Schluss.
    unsure = T(f" {how} passt es genau dazu: Ist es ein verrutschtes Komma oder eine andere Einheit, wäre es {_fmt(right)} "
               f"statt {_fmt(value)}.",
               f" {how}, it fits right in: if it is a misplaced decimal point or a different unit, it would be {_fmt(right)} "
               f"instead of {_fmt(value)}.")
    sure = T(f" {how} passt es genau dazu, wohl ein verrutschtes Komma oder eine andere Einheit. Richtig wäre {_fmt(right)} "
             f"statt {_fmt(value)}.",
             f" {how}, it fits right in, probably a misplaced decimal point or a different unit. The correct value would "
             f"be {_fmt(right)} instead of {_fmt(value)}.")
    sheet.memo.setdefault("analyse.scaled", {})[f"scale:{address(pos)}"] = (unsure, sure)
    return Finding(
        f"scale:{address(pos)}",
        T(f"{sheet.label(pos)} = {_fmt(value)} liegt weit weg von den übrigen Werten {where} (sonst {_fmt(low)} bis "
          f"{_fmt(high)}).{effect}{unsure}",
          f"{sheet.label(pos)} = {_fmt(value)} is far away from the other values {where} (otherwise {_fmt(low)} to "
          f"{_fmt(high)}).{effect}{unsure}"),
        "pruefen", T(f"Korrigiert in {address(pos)}: {_fmt(right)}. Stand es in einer anderen Einheit, rechne mit deren "
                     "genauem Faktor um. Ist es ein echter Sonderfall, lass es so.",
                     f"Fixed in {address(pos)}: {_fmt(right)}. If it was in a different unit, convert it with that unit's "
                     "exact factor. If it is a genuine special case, leave it as it is."),
        fix=einsetzen.values(sheet, {pos: right}, like, _fmt(right)),
    )


def _apart(scale: dict[int, float], center: float) -> set[int]:
    """Die Zeilen der höchstens drei äußersten Werte je Seite, die durch eine Lücke vom Rest getrennt sind.
    Ein Tippfehler (720 statt 72) steht allein; das obere Ende von Beträgen oder Kosten läuft dagegen ohne Lücke aus,
    dort ist der größte Wert kein Ausreißer, nur der größte."""
    found: set[int] = set()
    for side in (1, -1):
        order = sorted((r for r, v in scale.items() if (v - center) * side > 0), key=lambda r: -scale[r] * side)
        for i in range(min(3, len(order) - 1)):
            inner, outer = scale[order[i + 1]], scale[order[i]]
            if abs(outer - inner) >= 0.5 * abs(outer - center):
                found |= set(order[:i + 1])
                break
        if len(order) <= 3 and order and not found & set(order):
            # Kaum Werte auf dieser Seite: Der Abstand zur Mitte ist dann schon die Lücke.
            found |= set(order)
    return found


def _shares(sheet: Sheet, col: int, header: str, numbers: dict[int, float]) -> list[Finding]:
    if len(numbers) < 3:
        return []
    values = list(numbers.values())
    small = [v for v in values if 0 < abs(v) <= 1]
    big = [v for v in values if abs(v) > 1]
    # Gemischt nur, wenn es sonst keine kleinen Prozentwerte gibt: 0,5 % neben 1,2 % ist kein Bruch.
    if len(values) >= 3 and small and big and len(small) <= len(values) // 2 and max(map(abs, small)) < 1 and min(map(abs, big)) > 2:
        rows = [r for r, v in numbers.items() if 0 < abs(v) <= 1]
        adds_up = abs(math.fsum(big) + 100 * math.fsum(small) - 100) < 1.5
        if len(rows) == 1 and not adds_up:
            # Eine einzelne kleine Zahl kann auch ein kleiner Anteil sein (0,9 %): dann nur nachfragen.
            # Ergibt sie mal 100 mit den übrigen genau 100 %, ist es sicher ein Bruch.
            r = rows[0]
            return [Finding(
                f"mixed:{header}",
                T(f"In {header} stehen Anteile in Prozent, bei {address((r, col))} steht {_fmt(numbers[r])}. "
                  f"Sind das {_fmt(numbers[r] * 100)} % als Bruch geschrieben?",
                  f"{header} contains shares in percent, but {address((r, col))} has {_fmt(numbers[r])}. "
                  f"Is that {_fmt(numbers[r] * 100)} % written as a fraction?"),
                "pruefen", T("Schreib alle Anteile gleich, sonst stimmen Summen und Formeln nicht.",
                             "Write all shares the same way, otherwise totals and formulas will be wrong."),
            )]
        return [Finding(
            f"mixed:{header}",
            T(f"In {header} stehen die meisten Anteile in Prozent, bei {', '.join(address((r, col)) for r in rows[:3])} "
              f"aber als Bruch ({', '.join(_fmt(numbers[r]) for r in rows[:3])}).",
              f"In {header}, most shares are in percent, but in {', '.join(address((r, col)) for r in rows[:3])} "
              f"they are fractions ({'; '.join(_fmt(numbers[r]) for r in rows[:3])})."),
            "fehler", T("Schreib alle Anteile gleich, sonst stimmen Summen und Formeln nicht.",
                        "Write all shares the same way, otherwise totals and formulas will be wrong."),
        )]
    total = math.fsum(values)
    for target in (100.0, 1.0):
        if 0.9 * target < total < 1.1 * target and not _close(total, target) and abs(total - target) > 1e-9 * target:
            return [Finding(
                f"share:{header}",
                T(f"Die Anteile in {header} ergeben zusammen {_fmt(total)}, nicht {_fmt(target)}.",
                  f"The shares in {header} add up to {_fmt(total)}, not {_fmt(target)}."),
                "pruefen", T("Fehlt eine Zeile (z. B. „Übrige“), oder ist eine Zahl gerundet oder vertippt?",
                             "Is a row missing (e.g. “Other”), or is a number rounded or mistyped?"),
            )]
    return []


# --- Sensitivität ---


def _chain(sheet: Sheet, output: Pos) -> set[Pos]:
    seen, done, todo = set(), set(), [output]
    while todo:
        pos = todo.pop()
        if pos in done:
            continue
        done.add(pos)
        seen.add(pos)
        before = sheet.precedents.get(pos, ())
        if isinstance(before, Refs):
            # Große Bereiche auf einmal; weiter geht es nur über die Formeln darin.
            seen.update(before.singles)
            for area in before.areas:
                seen.update(sheet.columns.inside(area))
            todo += sheet.formula_inputs(pos)
        else:
            todo += before
    return seen


def _reads(sheet: Sheet, pos: Pos) -> int:
    """Wie viele Zellen eine Formel liest, mit denen ihrer Formeln eine Stufe weiter (ohne die ganze Kette)."""
    return sum(_size(sheet, p) for p in [pos, *sheet.formula_inputs(pos)])


def _size(sheet: Sheet, pos: Pos) -> int:
    before = sheet.precedents.get(pos, ())
    if not isinstance(before, Refs):
        return len(before)
    return len(before.singles) + sum(sheet.columns.count(a) for a in before.areas)


def _longest_chain(sheet: Sheet) -> int:
    """Die längste Kette von Formeln, die aufeinander aufbauen (ohne Rekursion, auch bei Laufsummen)."""
    formulas = set(sheet.nodes)
    waiting = {p: len([q for q in sheet.formula_inputs(p) if q != p]) for p in formulas}
    ready = [p for p, n in waiting.items() if n == 0]
    length = dict.fromkeys(ready, 1)
    while ready:
        pos = ready.pop()
        for after in sheet.dependents.get(pos, ()):
            if after not in waiting or after == pos:
                continue
            length[after] = max(length.get(after, 1), length[pos] + 1)
            waiting[after] -= 1
            if waiting[after] == 0:
                ready.append(after)
    return max(length.values(), default=0)


def _sensitivity(sheet: Sheet, deadline: float, tolerance: float = 1e-6) -> list[Driver]:
    outputs = [p for p in sheet.nodes if not sheet.dependents.get(p) and isinstance(sheet.shown.get(p), float)]
    if len(outputs) > MAX_CHAINED and sheet.columns is not None:
        # Tausende Ergebnisse über große Bereiche (eine Duplikat-Spalte, die nichts weiterverwendet): die ganze Kette
        # nur für die, die selbst und über ihre Formeln am meisten Zellen lesen.
        reads = {p: _reads(sheet, p) for p in outputs}
        outputs = sorted(outputs, key=lambda p: (-reads[p], p))[:MAX_CHAINED]
    # Zuerst die Ergebnisse, die auf den meisten Eingaben beruhen. Die Ketten selbst nur für die ersten: Über große
    # Bereiche hat jede Zehntausende Zellen.
    sizes = {p: len(_chain(sheet, p)) for p in outputs}
    outputs.sort(key=lambda p: (-sizes[p], p))
    chains = {p: _chain(sheet, p) for p in outputs[:MAX_OUTPUTS]}
    drivers: list[Driver] = []
    codes: dict[int, set[float]] = {}
    for output in outputs[:MAX_OUTPUTS]:
        chain = chains[output]
        if any(p in sheet.unknown for p in chain if p in sheet.nodes) or len(chain) < 2:
            continue
        base = _run(sheet, output, chain)
        shown = sheet.shown[output]
        if not isinstance(base, float) or not _close(base, shown, tolerance) or base == 0:
            continue
        candidates: list[tuple[str, str, object]] = []
        for pos in sorted(chain):
            if pos not in sheet.nodes and isinstance(sheet.shown.get(pos), float) and sheet.shown[pos] != 0:
                # 999 für „fehlt“ ist keine Eingabe, von der ein Ergebnis abhängen sollte (auswertung.py meldet es).
                if sheet.shown[pos] not in _column_codes(sheet, pos[1], codes):
                    candidates.append(("cell", sheet.label(pos), pos))
        constants = Counter(
            v for p in chain if sheet.nodes.get(p) is not None for v in set(formel.literals(sheet.nodes[p]))
            if v not in PLAIN_NUMBERS
        )
        for value, count in constants.most_common(5):
            places = T("Formel", "1 formula") if count == 1 else T(f"{count} Formeln", f"{count} formulas")
            candidates.append(("constant", T(f"die feste Zahl {_fmt(value)} in {places}",
                                             f"the fixed number {_fmt(value)} in {places}"), value))
        found = []
        for kind, name, what in candidates[:MAX_INPUTS]:
            if time.monotonic() > deadline:
                break
            if kind == "cell":
                x = sheet.shown[what]
                y = _run(sheet, output, chain, overrides={what: x * 1.01})
            else:
                x = what
                y = _run(sheet, output, chain, constants={what: what * 1.01})
            if not isinstance(y, float):
                continue
            elasticity = (y - base) / abs(base) / 0.01
            if abs(elasticity) >= 0.01:
                found.append(Driver(sheet.figure_label(output), _fmt(base), name, _fmt(x), round(elasticity, 2)))
        found.sort(key=lambda d: -abs(d.elasticity))
        drivers += found[:3]
    return drivers


def _column_codes(sheet: Sheet, col: int, cache: dict[int, set[float]]) -> set[float]:
    if col not in cache:
        from .auswertung import codes_in

        plan = sheet.plan
        if plan.header_row is None or col not in plan.headers:
            cache[col] = set()
        else:
            cache[col] = codes_in([v for (r, c), v in sheet.shown.items()
                                   if c == col and r > plan.header_row and isinstance(v, float) and (r, c) not in sheet.nodes])
    return cache[col]


def _run(sheet: Sheet, output: Pos, chain: set[Pos], overrides=None, constants=None):
    _, value = sheet.evaluator(overrides, constants, chain)
    try:
        return value(output)
    except (formel.Unknown, RecursionError):
        # Sehr lange Ketten (Laufsummen über Tausende Zeilen) rechnet er nicht nach.
        return None
