"""Zwei Versionen einer Mappe vergleichen: was sich geändert hat, und warum ein Ergebnis jetzt anders ist.

Vorher ist eine gespeicherte Datei, jetzt die offene Mappe. Zeilen und Spalten werden einander zugeordnet
(über Zeilennamen, Spaltenköpfe und Inhalt), damit eine eingefügte oder umsortierte Zeile nicht jede Formel
darunter als geändert zeigt. Die Formeln von vorher werden dafür auf die neue Lage umgeschrieben, so wie Excel
Bezüge beim Einfügen mitnimmt. Bei jedem Ergebnis mit neuem Wert steht, welche Änderung davor es verursacht.
Ohne KI, nur im Arbeitsspeicher.
"""

from __future__ import annotations

import re
from bisect import bisect_left
from collections import Counter
from dataclasses import dataclass, field
from difflib import SequenceMatcher

from . import herleitung
from .analyse import Sheet as Calc
from .analyse import address, calc_value, value_of
from .sprache import T
from .workbook import Layout, layout, pattern, readable
from .xlsx import Cell, Sheet, column_letters, column_number, outside_literals

Pos = tuple[int, int]

# So viele Zeilen werden einander zugeordnet; darunter zählt nur noch die Lage.
MAX_ROWS = 20000
MAX_ITEMS = 40
MAX_RESULTS = 8
MAX_EXAMPLES = 4
MAX_CHAIN = 50000
# Ab so viel Ähnlichkeit ist eine Zeile mit anderem Namen dieselbe Zeile, nur mit geändertem Namen.
SIMILAR = 0.5
VOLATILE = re.compile(r"\b(TODAY|NOW|RAND|RANDBETWEEN|RANDARRAY|OFFSET|INDIRECT)\(", re.IGNORECASE)
# Ein Bezug in einer Formel, auch auf ein anderes Blatt und als Bereich: B3, $B$3, Tarife!A2:B4, 'Daten 2024'!C1.
_REF = re.compile(
    r"(?<![A-Za-z0-9_.$\\])(?:('(?:[^']|'')+'|[^\W\d][\w.]*)!)?"
    r"(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Z]{1,3})(\$?)(\d{1,7}))?(?![A-Za-z0-9_(!])"
)
ORDER = {"ergebnis": 0, "folge": 1, "ersetzt": 2, "formel": 3, "eingabe": 4, "zeilen": 5, "spalten": 6}
LEVEL = {"ergebnis": "pruefen", "ersetzt": "fehler", "formel": "pruefen"}


@dataclass
class Axis:
    """Welche Zeile (oder Spalte) von vorher welcher von jetzt entspricht."""

    pairs: dict[int, int] = field(default_factory=dict)
    added: list[int] = field(default_factory=list)
    removed: list[int] = field(default_factory=list)
    # Umsortiert: dieselbe Zeile steht jetzt außer der Reihe anderswo.
    moved: dict[int, int] = field(default_factory=dict)
    _old: list[int] = field(default_factory=list)
    _new: list[int] = field(default_factory=list)

    def done(self) -> Axis:
        steps = sorted((o, n) for o, n in self.pairs.items() if o not in self.moved)
        self._old, self._new = [o for o, _ in steps], [n for _, n in steps]
        return self

    def back(self) -> dict[int, int]:
        return {n: o for o, n in self.pairs.items()}

    def new(self, old: int, end: bool = False) -> int:
        """Wohin ein Bezug auf diese Zeile von vorher jetzt zeigt, so wie Excel ihn beim Einfügen mitnimmt.
        Eine gelöschte Zeile: Ein Bereich beginnt dann bei der nächsten und endet bei der vorigen."""
        if old in self.pairs and old not in self.moved:
            return self.pairs[old]
        if not self._old:
            return old
        i = bisect_left(self._old, old)
        if end and i > 0:
            return self._new[i - 1] + (0 if i < len(self._old) else old - self._old[i - 1])
        if i < len(self._old):
            return max(1, self._new[i] - (self._old[i] - old) if i == 0 else self._new[i])
        return self._new[-1] + old - self._old[-1]


def align(old: list, new: list, similar=None) -> Axis:
    """Zwei Folgen von Schlüsseln (je Zeile oder Spalte) einander zuordnen. similar(i, j): ob die Zeile i von
    vorher und j von jetzt trotz anderem Schlüssel dieselbe sind (sonst gilt die eine als gelöscht, die andere als neu)."""
    axis = Axis()
    for tag, i1, i2, j1, j2 in SequenceMatcher(None, old, new, autojunk=False).get_opcodes():
        if tag == "equal":
            axis.pairs.update((i1 + k + 1, j1 + k + 1) for k in range(i2 - i1))
            continue
        k = 0
        if tag == "replace":
            while k < min(i2 - i1, j2 - j1) and (similar is None or similar(i1 + k + 1, j1 + k + 1)):
                axis.pairs[i1 + k + 1] = j1 + k + 1
                k += 1
        axis.removed += range(i1 + k + 1, i2 + 1)
        axis.added += range(j1 + k + 1, j2 + 1)
    # Umsortiert: gleicher, eindeutiger Schlüssel unter den gelöschten und den neuen Zeilen.
    gone = Counter(old[i - 1] for i in axis.removed)
    came = Counter(new[j - 1] for j in axis.added)
    where = {new[j - 1]: j for j in axis.added}
    for i in list(axis.removed):
        key = old[i - 1]
        if key[0] != "" and gone[key] == 1 and came[key] == 1:
            j = where[key]
            axis.pairs[i], axis.moved[i] = j, j
            axis.removed.remove(i)
            axis.added.remove(j)
    return axis.done()


def _norm(formula: str) -> str:
    """Zum Vergleichen: ohne Leerzeichen, „_xlfn.“ (so speichert die Datei neue Funktionen) und „@“ von Excel."""
    return outside_literals(formula.lstrip("="), lambda part: re.sub(r"_XL(?:FN|WS)\.|@|\s", "", part.upper()))


def _unquote(name: str) -> str:
    return name[1:-1].replace("''", "'") if name.startswith("'") else name


def _quote(name: str) -> str:
    return name if re.fullmatch(r"[^\W\d][\w.]*", name) else "'" + name.replace("'", "''") + "'"


def translate(formula: str, own: str, axes: dict[str, tuple[Axis, Axis, str]], move: bool = True) -> str:
    """Eine Formel von vorher auf die Lage von jetzt umschreiben. axes: Blatt von vorher (klein) → (Zeilen,
    Spalten, Name jetzt). move=False: nur umbenannte Blätter umschreiben."""

    def one(dollar_col, letters, dollar_row, digits, rows: Axis, cols: Axis, end: bool) -> str:
        if not move:
            return f"{dollar_col}{letters}{dollar_row}{digits}"
        return f"{dollar_col}{column_letters(cols.new(column_number(letters), end))}{dollar_row}{rows.new(int(digits), end)}"

    def sub(m: re.Match) -> str:
        sheet = m.group(1)
        found = axes.get((_unquote(sheet) if sheet else own).casefold())
        if found is None:
            return m.group(0)
        rows, cols, now = found
        text = f"{_quote(now)}!" if sheet else ""
        text += one(*m.group(2, 3, 4, 5), rows, cols, end=False)
        if m.group(7):
            text += ":" + one(*m.group(6, 7, 8, 9), rows, cols, end=True)
        return text

    return outside_literals(formula, lambda part: _REF.sub(sub, part))


def _value(cell: Cell | None):
    value = value_of(cell)
    return None if value is None else calc_value(value, cell)


def same(a: Cell | None, b: Cell | None) -> bool:
    x, y = _value(a), _value(b)
    if isinstance(x, float) and isinstance(y, float) and not isinstance(x, bool) and not isinstance(y, bool):
        return abs(x - y) <= 1e-9 * max(1.0, abs(x), abs(y))
    return type(x) is type(y) and x == y


def _change(a: Cell, b: Cell) -> str:
    """„1.200 → 1.350 (+12,5 %)“."""
    x, y = _value(a), _value(b)
    text = f"{a.value} → {b.value}"
    if isinstance(x, float) and isinstance(y, float) and x and not isinstance(x, bool) and not isinstance(y, bool):
        rel = (y - x) / abs(x) * 100
        if abs(rel) >= 0.05 and not a.value.endswith("%") and not b.value.endswith("%"):
            text += f" ({'+' if rel > 0 else '−'}{abs(rel):.1f} %)".replace(".", ",")
    return text


def _by_row(cells: dict[Pos, Cell]) -> dict[int, list[tuple[int, Cell]]]:
    rows: dict[int, list[tuple[int, Cell]]] = {}
    for (r, c), cell in sorted(cells.items()):
        rows.setdefault(r, []).append((c, cell))
    return rows


def _row_keys(cells: dict[Pos, Cell], plan: Layout, last: int) -> list:
    rows = _by_row(cells)
    keys = []
    for r in range(1, last + 1):
        name = plan.names.get(r) if plan.name_column else None
        if name:
            keys.append(("n", name))
            continue
        line = rows.get(r, [])
        values = tuple(c.value for _, c in line if c.formula is None and c.value != "")
        if values:
            keys.append(("v", values))
        else:
            formulas = tuple(pattern(c.formula, (r, col)) for col, c in line if c.formula)
            keys.append(("f", formulas) if formulas else ("",))
    return keys


def _col_keys(cells: dict[Pos, Cell], plan: Layout, last: int) -> list:
    columns: dict[int, list[str]] = {}
    for (_, c), cell in sorted(cells.items()):
        if cell.formula is None and cell.value and len(columns.setdefault(c, [])) < 5:
            columns[c].append(cell.value)
    keys = []
    for c in range(1, last + 1):
        if plan.header_row and plan.headers.get(c):
            keys.append(("h", plan.headers[c]))
        else:
            keys.append(("c", tuple(columns[c])) if c in columns else ("",))
    return keys


@dataclass
class Side:
    """Ein Blatt in einer der beiden Versionen."""

    name: str
    cells: dict[Pos, Cell]
    plan: Layout = None  # type: ignore[assignment]
    last: tuple[int, int] = (0, 0)
    height: int = 0

    def __post_init__(self) -> None:
        self.plan = layout(Sheet(self.name, self.cells))
        self.height = max((r for r, _ in self.cells), default=0)
        self.last = (min(self.height, MAX_ROWS), max((c for _, c in self.cells), default=0))
        self.rows = {r for r, _ in self.cells}
        self.cols = {c for _, c in self.cells}


def _sheet_axes(old: Side, new: Side) -> tuple[Axis, Axis]:
    cols = align(_col_keys(old.cells, old.plan, old.last[1]), _col_keys(new.cells, new.plan, new.last[1]))
    old_rows, new_rows = _by_row(old.cells), _by_row(new.cells)

    def similar(i: int, j: int) -> bool:
        a = {cols.pairs.get(c): cell.value for c, cell in old_rows.get(i, []) if cell.formula is None}
        b = {c: cell.value for c, cell in new_rows.get(j, []) if cell.formula is None}
        if not a and not b:
            return True
        hits = sum(1 for c, v in b.items() if a.get(c) == v)
        return hits >= SIMILAR * max(len(a), len(b))

    rows = align(_row_keys(old.cells, old.plan, old.last[0]), _row_keys(new.cells, new.plan, new.last[0]), similar)
    return rows, cols


def _match_sheets(old: dict[str, dict], new: dict[str, dict]) -> list[tuple[str | None, str | None]]:
    """Blätter nach Namen zuordnen; ein umbenanntes Blatt am Inhalt erkennen."""
    names = {n.casefold(): n for n in new}
    pairs: list[tuple[str | None, str | None]] = []
    used = set()
    for o in old:
        n = names.get(o.casefold())
        if n is not None:
            pairs.append((o, n))
            used.add(n)
    gone = [o for o in old if not any(o == p[0] for p in pairs)]
    came = [n for n in new if n not in used]
    for o in gone:
        best, score = None, 0.0
        for n in came:
            a, b = old[o], new[n]
            if not a or not b:
                continue
            hits = sum(1 for p, cell in a.items() if p in b and b[p].value == cell.value)
            overlap = hits / max(len(a), len(b))
            if overlap > score:
                best, score = n, overlap
        if best is not None and score >= 0.6:
            pairs.append((o, best))
            came.remove(best)
        else:
            pairs.append((o, None))
    pairs += [(None, n) for n in came]
    order = {n: i for i, n in enumerate(new)}
    return sorted(pairs, key=lambda p: (order.get(p[1], len(order)), p[0] or ""))


def compare(old: dict[str, dict[Pos, Cell]], new: dict[str, dict[Pos, Cell]], before: str = "", now: str = "") -> dict:
    """old, new: je Blatt (Name → Zellen). Gibt je Blatt die Unterschiede zurück, Ergebnisse zuerst."""
    pairs = _match_sheets(old, new)
    sides = {o: (Side(o, old[o]), Side(n, new[n])) for o, n in pairs if o is not None and n is not None}
    axes = {}
    for o, (a, b) in sides.items():
        rows, cols = _sheet_axes(a, b)
        axes[o.casefold()] = (rows, cols, b.name)
    found = {o: _sheet_diff(a, b, axes, new) for o, (a, b) in sides.items()}
    # Ein Ergebnis kann sich wegen eines anderen Blattes ändern: das bei den Gründen nennen.
    touched = {sides[o][1].name.casefold(): d.touched() for o, d in found.items()}
    sheets = []
    for o, n in pairs:
        if o is None:
            sheets.append({"name": n, "old": None, "state": "neu", "items": [], "more": 0})
        elif n is None:
            sheets.append({"name": o, "old": o, "state": "weg", "items": [], "more": 0})
        else:
            diff = found[o]
            items = diff.items({k: v for k, v in touched.items() if v and k != n.casefold()})
            sheets.append({"name": n, "old": o if o != n else None, "state": "geaendert" if items else "gleich",
                           "items": items[:MAX_ITEMS], "more": max(0, len(items) - MAX_ITEMS), "cut": diff.cut})
    counts: Counter = Counter()
    for s in sheets:
        for i in s["items"]:
            counts[i["kind"] + (f"_{i['what']}" if i.get("what") else "")] += i.get("count", 1)
    counts["blatt"] = sum(1 for s in sheets if s["state"] in ("neu", "weg") or s["old"])
    # Kein Blatt gleich: Dann ist es wohl eine ganz andere Mappe, keine frühere Version.
    related = bool(sides)
    return {"state": "ok", "before": before, "now": now, "counts": dict(counts), "text": _summary(counts, sheets),
            "sheets": sheets, "related": related}


_SUMMARY = (("ergebnis", "1 Ergebnis ist anders", "{n} Ergebnisse sind anders"),
            ("ersetzt", "1 Formel ist jetzt eine feste Zahl", "{n} Formeln sind jetzt feste Zahlen"),
            ("formel", "1 Formel ist geändert", "{n} Formeln sind geändert"),
            ("eingabe", "1 Eingabe ist geändert", "{n} Eingaben sind geändert"),
            ("zeilen_neu", "1 Zeile ist neu", "{n} Zeilen sind neu"),
            ("zeilen_weg", "1 Zeile ist weg", "{n} Zeilen sind weg"),
            ("spalten_neu", "1 Spalte ist neu", "{n} Spalten sind neu"),
            ("spalten_weg", "1 Spalte ist weg", "{n} Spalten sind weg"),
            ("blatt", "1 Blatt ist neu, weg oder umbenannt", "{n} Blätter sind neu, weg oder umbenannt"))
_SUMMARY_EN = {"ergebnis": ("1 result is different", "{n} results are different"),
               "ersetzt": ("1 formula is now a fixed number", "{n} formulas are now fixed numbers"),
               "formel": ("1 formula has changed", "{n} formulas have changed"),
               "eingabe": ("1 input has changed", "{n} inputs have changed"),
               "zeilen_neu": ("1 row is new", "{n} rows are new"),
               "zeilen_weg": ("1 row is gone", "{n} rows are gone"),
               "spalten_neu": ("1 column is new", "{n} columns are new"),
               "spalten_weg": ("1 column is gone", "{n} columns are gone"),
               "blatt": ("1 sheet is new, gone or renamed", "{n} sheets are new, gone or renamed")}


def _summary(counts: Counter, sheets: list[dict]) -> str:
    parts = []
    for kind, one_de, many_de in _SUMMARY:
        one, many = T(one_de, _SUMMARY_EN[kind][0]), T(many_de, _SUMMARY_EN[kind][1])
        n = counts.get(kind, 0)
        if n:
            parts.append(one if n == 1 else many.format(n=n))
    if not parts:
        return T("In Zellen und Formeln hat sich nichts geändert. (Formatierungen vergleiche ich nicht.)",
                 "Nothing has changed in cells and formulas. (I do not compare formatting.)")
    return ", ".join(parts[:-1]) + (T(" und ", " and ") if len(parts) > 1 else "") + parts[-1] + "."


class _Diff:
    """Die Unterschiede eines Blatts, in den Zellen von jetzt."""

    def __init__(self, old: Side, new: Side, rows: Axis, cols: Axis, book: dict[str, dict]) -> None:
        self.old, self.new, self.rows, self.cols = old, new, rows, cols
        self.calc = Calc(new.cells, new.plan, new.name, book)
        self.inputs: list[tuple[Pos, Cell | None, Cell | None]] = []
        # Geänderte Formeln, auch Zahl → Formel (dann ist die Zelle von vorher eine Eingabe oder fehlt).
        self.formulas: list[tuple[Pos, Pos, Cell | None, Cell]] = []
        self.replaced: list[tuple[Pos, Cell, Cell | None]] = []
        self.results: list[tuple[Pos, Cell, Cell]] = []
        self.causes: dict[Pos, str] = {}
        self.names = {n.casefold(): n for n in book}
        self.book = book
        self._recalc: herleitung.Book | None = None
        self.cut = max(old.height, new.height) > MAX_ROWS

    def label(self, pos: Pos) -> str:
        return self.calc.figure_label(pos)

    def touched(self) -> set[Pos]:
        """Was sich auf diesem Blatt geändert hat (in Zellen von jetzt), für Gründe auf anderen Blättern."""
        added = set(self.rows.added) & self.new.rows
        return set(self.causes) | {p for p in self.new.cells if p[0] in added}

    def items(self, other_sheets: dict[str, set[Pos]]) -> list[dict]:
        found = self._results(other_sheets) + self._replaced() + self._formulas() + self._inputs() + self._lines()
        return sorted(found, key=lambda i: ORDER[i["kind"]])

    # --- Ergebnisse und warum ---

    def _results(self, other_sheets: dict[str, set[Pos]]) -> list[dict]:
        changed = {p for p, *_ in self.results}
        dependents = self.calc.dependents
        # Zuerst die Ergebnisse am Ende der Rechnung: Dazwischen ändert sich ohnehin alles mit.
        top = [r for r in self.results if not (set(dependents.get(r[0], ())) & changed)]
        items = []
        for pos, before, after in top[:MAX_RESULTS]:
            why, cells = self._why(pos, other_sheets)
            items.append({"kind": "ergebnis", "level": LEVEL["ergebnis"], "cell": address(pos), "cells": cells,
                          "text": f"{self.label(pos)}: {_change(before, after)}", "why": why})
        rest = len(self.results) - len(items)
        if rest > 0:
            text = (T("Dazu 1 weitere Formel mit neuem Wert, die auf dem Weg dorthin liegt.",
                      "Plus 1 more formula with a new value along the way there.") if rest == 1
                    else T(f"Dazu {rest} weitere Formeln mit neuem Wert, die auf dem Weg dorthin liegen.",
                           f"Plus {rest} more formulas with new values along the way there."))
            items.append({"kind": "folge", "level": "info", "cell": None, "cells": [], "text": text, "why": ""})
        return items

    def _why(self, pos: Pos, other_sheets: dict[str, set[Pos]]) -> tuple[str, list[str]]:
        calc, seen, stack = self.calc, set(), [pos]
        inputs, formulas, volatile = [], [], False
        sheets: dict[str, set[Pos]] = {}
        added = set(self.rows.added)
        gaps = {self.rows.new(r) for r in self.rows.removed}
        grew = shrank = False
        while stack and len(seen) < MAX_CHAIN:
            p = stack.pop()
            if p in seen:
                continue
            seen.add(p)
            if p != pos and p in self.causes:
                (formulas if self.causes[p] != "eingabe" else inputs).append(p)
            node = calc.nodes.get(p)
            if node is not None:
                if other_sheets:
                    _foreign_hits(node, other_sheets, sheets)
                for r1, c1, r2, c2 in calc.ranges.get(p, []):
                    grew = grew or any(r1 <= r <= r2 for r in added)
                    shrank = shrank or any(r1 < g <= r2 for g in gaps)
            cell = calc.cells.get(p)
            if cell is not None and cell.formula and VOLATILE.search(cell.formula):
                volatile = True
            stack.extend(calc.precedents.get(p, ()))
        parts = self._wrong(seen)
        if inputs:
            inputs.sort()
            if len(inputs) == 1:
                p = inputs[0]
                parts.append(f"{self.label(p)} {self._input_change(p)}")
            else:
                shown = ", ".join(address(p) for p in inputs[:MAX_EXAMPLES]) + (" …" if len(inputs) > MAX_EXAMPLES else "")
                parts.append(T(f"sich {len(inputs)} Eingaben davor geändert haben ({shown})",
                               f"{len(inputs)} inputs before it have changed ({shown})"))
        if formulas:
            formulas.sort()
            shown = ", ".join(address(p) for p in formulas[:MAX_EXAMPLES]) + (" …" if len(formulas) > MAX_EXAMPLES else "")
            parts.append(T(f"die Formel in {shown} anders ist", f"the formula in {shown} is different") if len(formulas) == 1
                         else T(f"die Formeln in {shown} anders sind", f"the formulas in {shown} are different"))
        if grew:
            parts.append(T("neue Zeilen mitzählen", "new rows are counted"))
        if shrank:
            parts.append(T("Zeilen aus dem Bereich gelöscht sind", "rows were deleted from the range"))
        for key, hits in sorted(sheets.items()):
            cells = sorted(hits)
            shown = ", ".join(address(p) for p in cells[:MAX_EXAMPLES]) + (" …" if len(cells) > MAX_EXAMPLES else "")
            name = self.names.get(key, key)
            parts.append(T(f"auf Blatt {name} {shown} geändert ist", f"{shown} on sheet {name} has changed") if len(cells) == 1
                         else T(f"sich auf Blatt {name} {len(cells)} Zellen geändert haben ({shown})",
                                f"{len(cells)} cells on sheet {name} have changed ({shown})"))
        if not parts:
            if volatile:
                return T("Davor hat sich nichts geändert, aber die Rechnung nutzt HEUTE(), ZUFALLSZAHL() oder Ähnliches.",
                         "Nothing before it has changed, but the calculation uses TODAY(), RAND() or similar."), []
            return T("Davor hat sich nichts geändert, was das erklärt. Eine der beiden Versionen war wohl nicht neu "
                     "berechnet (Berechnung auf manuell?). Ob die offene Mappe stimmt, zeigt der Reiter Hinweise.",
                     "Nothing before it has changed that explains this. One of the two versions was probably not "
                     "recalculated (calculation set to manual?). The Findings tab shows whether the open workbook is "
                     "right."), []
        text = T("Weil ", "Because ") + ", ".join(parts[:-1]) + (T(" und ", " and ") if len(parts) > 1 else "") + parts[-1] + "."
        return text, [address(p) for p in (inputs + formulas)[:MAX_EXAMPLES * 2]]

    def _wrong(self, chain: set[Pos]) -> list[str]:
        """Formeln auf dem Weg mit neuem Wert, die in Excel anders zeigen, als sie nachgerechnet ergeben: Dann liegt
        der Unterschied nicht an einer Änderung, sondern daran, dass Excel nicht neu gerechnet hat."""
        changed = {p for p, *_ in self.results}
        found = []
        for p in sorted(chain & changed):
            node = self.calc.nodes.get(p)
            if node is None or p in self.calc.unknown:
                continue
            if self._recalc is None:
                self._recalc = herleitung.Book(self.book)
            ours = self._recalc.run(self.new.name, node)[0]
            state, _ = herleitung._check(self.calc.shown.get(p), ours)
            if state == "anders" and isinstance(ours, float):
                shown = self.new.cells[p].value
                found.append(T(f"{self.label(p)} in Excel {shown} zeigt, nachgerechnet aber {herleitung._like(shown, ours)} ergibt",
                               f"{self.label(p)} shows {shown} in Excel but gives {herleitung._like(shown, ours)} "
                               "when recalculated"))
            if len(found) >= 2:
                break
        return found

    def _input_change(self, pos: Pos) -> str:
        """Für „Weil B4 …“: „von 200 auf 250 geändert ist“, „neu eingetragen ist (45)“, „gelöscht ist (vorher 45)“."""
        for p, before, after in self.inputs:
            if p == pos:
                if before is None:
                    return T(f"neu eingetragen ist ({after.value})", f"was newly entered ({after.value})")
                if after is None:
                    return T(f"gelöscht ist (vorher {before.value})", f"was deleted (before {before.value})")
                return T(f"von {before.value} auf {after.value} geändert ist",
                         f"changed from {before.value} to {after.value}")
        return T("geändert ist", "has changed")

    # --- Formeln ---

    def _replaced(self) -> list[dict]:
        items = []
        for pos, before, after in self.replaced[:MAX_ITEMS]:
            if after is None:
                text = T(f"{self.label(pos)}: Die Formel ={before.typed or before.formula} ist gelöscht.",
                         f"{self.label(pos)}: The formula ={before.typed or before.formula} has been deleted.")
                tip = T("Gewollt? Was darauf aufbaut, rechnet jetzt mit einer leeren Zelle.",
                        "Intended? Whatever builds on it now calculates with an empty cell.")
            else:
                text = T(f"{self.label(pos)}: Statt der Formel ={before.typed or before.formula} steht jetzt die feste Zahl {after.value}.",
                         f"{self.label(pos)}: Instead of the formula ={before.typed or before.formula} there is now the fixed "
                         f"number {after.value}.")
                tip = T("Ändern sich die Eingaben, rechnet diese Zelle nicht mehr mit. Steht die Formel noch in den Zeilen daneben?",
                        "If the inputs change, this cell no longer updates. Is the formula still in the neighbouring rows?")
            items.append({"kind": "ersetzt", "level": LEVEL["ersetzt"], "cell": address(pos), "cells": [address(pos)],
                          "text": text, "tip": tip})
        return items

    def _formulas(self) -> list[dict]:
        groups: dict[tuple, list[tuple[Pos, Pos, Cell | None, Cell]]] = {}
        for item in self.formulas:
            pos, old_pos, before, after = item
            was = "" if before is None else pattern(before.formula, old_pos) if before.formula else "zahl"
            groups.setdefault((pos[1], was, pattern(after.formula, pos)), []).append(item)
        items = []
        for (col, _, _), group in groups.items():
            pos, old_pos, before, after = group[0]
            now = readable(after.formula, pos, self.new.plan)
            header = self.new.plan.headers.get(col)
            where = T(f"Spalte {column_letters(col)}", f"column {column_letters(col)}") + (f" ({header})" if header else "")
            span = T(f"in {len(group)} Zeilen ({address(group[0][0])} bis {address(group[-1][0])})",
                     f"in {len(group)} rows ({address(group[0][0])} to {address(group[-1][0])})")
            if before is None:
                text = T(f"Neue Formel in {self.label(pos)}: {now}.", f"New formula in {self.label(pos)}: {now}.") \
                    if len(group) == 1 else T(f"Neue Formel in {where} {span}: {now}.", f"New formula in {where} {span}: {now}.")
            elif before.formula is None:
                text = (T(f"{self.label(pos)}: Statt der Zahl {before.value} steht jetzt die Formel {now}.",
                          f"{self.label(pos)}: Instead of the number {before.value} there is now the formula {now}.")
                        if len(group) == 1
                        else T(f"In {where} {span} steht jetzt eine Formel statt einer Zahl: {now}.",
                               f"{where[:1].upper()}{where[1:]} now has a formula instead of a number {span}: {now}."))
            elif len(group) == 1:
                was = readable(before.formula, old_pos, self.old.plan)
                text = T(f"Formel in {self.label(pos)}: vorher {was}, jetzt {now}.",
                         f"Formula in {self.label(pos)}: before {was}, now {now}.")
            else:
                was = readable(before.formula, old_pos, self.old.plan)
                text = T(f"Formel in {where} {span} geändert: vorher {was}, jetzt {now}.",
                         f"Formula in {where} {span} changed: before {was}, now {now}.")
            if len(group) == 1 and before is not None and not same(before, after):
                text += T(f" Wert: {_change(before, after)}.", f" Value: {_change(before, after)}.")
            items.append({"kind": "formel", "level": LEVEL["formel"], "cell": address(pos), "count": len(group),
                          "cells": [address(g[0]) for g in group[:MAX_EXAMPLES]], "text": text})
        return items

    # --- Eingaben und Zeilen ---

    def _inputs(self) -> list[dict]:
        by_col: dict[int, list[tuple[Pos, Cell | None, Cell | None]]] = {}
        for item in self.inputs:
            by_col.setdefault(item[0][1], []).append(item)
        items = []
        for col, group in sorted(by_col.items()):
            header = self.new.plan.headers.get(col)
            where = header or T(f"Spalte {column_letters(col)}", f"column {column_letters(col)}")
            empty = Cell(T("leer", "empty"))
            shown = "; ".join(f"{address(p)} {_change(a or empty, b or empty)}" for p, a, b in group[:MAX_EXAMPLES])
            more = T(f" und {len(group) - MAX_EXAMPLES} weitere", f" and {len(group) - MAX_EXAMPLES} more") \
                if len(group) > MAX_EXAMPLES else ""
            if len(group) == 1:
                p, a, b = group[0]
                text = f"{self.label(p)}: {_change(a or empty, b or empty)}."
            else:
                text = T(f"In {where} {len(group)} Werte geändert: {shown}{more}.",
                         f"{len(group)} values changed in {where}: {shown}{more}.")
            items.append({"kind": "eingabe", "level": "info", "cell": address(group[0][0]), "count": len(group),
                          "cells": [address(p) for p, *_ in group[:MAX_EXAMPLES]], "text": text})
        return items

    def _lines(self) -> list[dict]:
        items = []
        rows, cols, old, new = self.rows, self.cols, self.old, self.new
        added = [r for r in rows.added if r in new.rows]
        removed = [r for r in rows.removed if r in old.rows]
        if added:
            names = [new.plan.names[r] for r in added if r in new.plan.names]
            text = (T(f"{len(added)} {'Zeile' if len(added) == 1 else 'Zeilen'} neu: {_rows_text(added)}",
                      f"{len(added)} new {'row' if len(added) == 1 else 'rows'}: {_rows_text(added)}")
                    + (f" ({_names(names)})" if names else "") + ".")
            first = min(c for r, c in new.cells if r == added[0])
            items.append({"kind": "zeilen", "what": "neu", "count": len(added), "level": "info",
                          "cell": address((added[0], first)), "cells": [], "text": text})
        if removed:
            names = [old.plan.names[r] for r in removed if r in old.plan.names]
            text = (T(f"{len(removed)} {'Zeile' if len(removed) == 1 else 'Zeilen'} von vorher gibt es nicht mehr",
                      f"{len(removed)} {'row' if len(removed) == 1 else 'rows'} from before {'is' if len(removed) == 1 else 'are'} "
                      "gone")
                    + (f": {_names(names)}" if names else T(f" (vorher {_rows_text(removed)})", f" (before {_rows_text(removed)})"))
                    + ".")
            items.append({"kind": "zeilen", "what": "weg", "count": len(removed), "level": "info", "cell": None, "cells": [], "text": text})
        if rows.moved:
            n = len(rows.moved)
            text = T("1 Zeile steht" if n == 1 else f"{n} Zeilen stehen", "1 row is" if n == 1 else f"{n} rows are")
            items.append({"kind": "zeilen", "what": "sortiert", "count": n, "level": "info", "cell": None, "cells": [],
                          "text": T(f"{text} an anderer Stelle (umsortiert). Ich vergleiche trotzdem Zeile mit Zeile.",
                                    f"{text} in a different place (re-sorted). I still compare row by row.")})
        for kind, side, version in (("neu", cols.added, new), ("weg", cols.removed, old)):
            used, plan = [c for c in side if c in version.cols], version.plan
            if used:
                heads = [plan.headers.get(c) or column_letters(c) for c in used]
                if len(used) > 1:
                    text = T(f"Spalten {_names(heads)} sind neu." if kind == "neu" else
                             f"Spalten {_names(heads)} von vorher gibt es nicht mehr.",
                             f"Columns {_names(heads)} are new." if kind == "neu" else
                             f"Columns {_names(heads)} from before are gone.")
                else:
                    text = T(f"Spalte {_names(heads)} ist neu." if kind == "neu" else
                             f"Spalte {_names(heads)} von vorher gibt es nicht mehr.",
                             f"Column {_names(heads)} is new." if kind == "neu" else
                             f"Column {_names(heads)} from before is gone.")
                items.append({"kind": "spalten", "what": kind, "count": len(used), "level": "info",
                              "cell": f"{column_letters(used[0])}1" if kind == "neu" else None,
                              "cells": [], "text": text})
        return items


def _foreign_hits(node, other: dict[str, set[Pos]], found: dict[str, set[Pos]]) -> None:
    """Bezüge auf andere Blätter, in denen sich etwas geändert hat: je Blatt die geänderten Zellen darin."""
    if node.kind in ("xref", "xrange"):
        sheet, where = node.value
        changed = other.get(sheet.casefold())
        if changed:
            r1, c1, r2, c2 = where * 2 if node.kind == "xref" else where
            hits = {p for p in changed if r1 <= p[0] <= r2 and c1 <= p[1] <= c2}
            if hits:
                found.setdefault(sheet.casefold(), set()).update(hits)
    for child in node.args:
        if hasattr(child, "kind"):
            _foreign_hits(child, other, found)


def _rows_text(rows: list[int]) -> str:
    rows = sorted(rows)
    if len(rows) > 2 and rows[-1] - rows[0] == len(rows) - 1:
        return T(f"Zeilen {rows[0]} bis {rows[-1]}", f"rows {rows[0]} to {rows[-1]}")
    shown = ", ".join(str(r) for r in rows[:6]) + (" …" if len(rows) > 6 else "")
    return T(f"{'Zeile' if len(rows) == 1 else 'Zeilen'} {shown}", f"{'row' if len(rows) == 1 else 'rows'} {shown}")


def _names(names: list[str]) -> str:
    shown = ", ".join(T(f"„{n}“", f"“{n}”") for n in names[:6])
    return shown + (T(f" und {len(names) - 6} weitere", f" and {len(names) - 6} more") if len(names) > 6 else "")


def _sheet_diff(old: Side, new: Side, axes: dict[str, tuple[Axis, Axis, str]], book: dict[str, dict]) -> _Diff:
    rows, cols, _ = axes[old.name.casefold()]
    diff = _Diff(old, new, rows, cols, book)
    back_rows, back_cols = rows.back(), cols.back()
    seen_old = set()
    for pos, after in sorted(new.cells.items()):
        r, c = pos
        if r > MAX_ROWS or r not in back_rows or c not in back_cols:
            continue
        old_pos = (back_rows[r], back_cols[c])
        seen_old.add(old_pos)
        before = old.cells.get(old_pos)
        _classify(diff, pos, old_pos, before, after, axes)
    for old_pos, before in old.cells.items():
        if old_pos in seen_old or old_pos[0] > MAX_ROWS:
            continue
        r, c = old_pos
        if r in rows.pairs and c in cols.pairs:
            _classify(diff, (rows.pairs[r], cols.pairs[c]), old_pos, before, None, axes)
    for name in ("inputs", "formulas", "replaced", "results"):
        getattr(diff, name).sort(key=lambda item: item[0])
    return diff


def _classify(diff: _Diff, pos: Pos, old_pos: Pos, before: Cell | None, after: Cell | None, axes) -> None:
    empty_before = before is None or (before.formula is None and before.value == "")
    empty_after = after is None or (after.formula is None and after.value == "")
    if empty_before and empty_after:
        return
    if before is not None and before.formula:
        if after is None or not after.formula:
            diff.replaced.append((pos, before, None if empty_after else after))
            diff.causes[pos] = "ersetzt"
            return
        # Gleich, wenn die Formel relativ zu ihrer Zelle gleich ist (Zeile verschoben oder umsortiert) …
        moved_same = _norm(pattern(translate(before.formula, diff.old.name, axes, move=False), old_pos)) == _norm(pattern(after.formula, pos))
        # … oder wenn Excel sie beim Einfügen und Löschen von Zeilen so mitgenommen hätte.
        if not moved_same and _norm(translate(before.formula, diff.old.name, axes)) != _norm(after.formula):
            diff.formulas.append((pos, old_pos, before, after))
            diff.causes[pos] = "formel"
        elif not same(before, after):
            diff.results.append((pos, before, after))
        return
    if after is not None and after.formula:
        # Vorher eine Zahl oder leer, jetzt eine Formel.
        diff.formulas.append((pos, old_pos, None if empty_before else before, after))
        diff.causes[pos] = "formel"
        return
    if not same(before, after):
        diff.inputs.append((pos, None if empty_before else before, None if empty_after else after))
        diff.causes[pos] = "eingabe"
