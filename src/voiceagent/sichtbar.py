"""Was man sieht und was gezählt wird: SUMME, MITTELWERT und Co. zählen ausgeblendete und weggefilterte Zeilen mit,
TEILERGEBNIS(109; …) nur die sichtbaren. Wer eine gefilterte Liste ansieht und darunter die Summe liest, hält sie oft für
die Summe dessen, was er sieht. Der Prüfer rechnet beides und sagt den Unterschied: „B20 zählt auch die 3
weggefilterten Zeilen mit: Was der Filter zeigt, ergibt 4560, B20 zeigt 5790.“ Einsetzen schreibt TEILERGEBNIS.

Weggefiltert ist meist ein Versehen in der Summe (darum „prüfen“), von Hand ausgeblendet oft gewollt (darum nur ein
Hinweis). Eine eingeklappte Gruppe (Daten → Gruppieren, nur aus der Datei bekannt: Office.js sagt nicht, welche
Zeilen gruppiert sind) und eine Summe nur über ausgeblendete Zeilen bleiben still."""

from __future__ import annotations

import re
from bisect import bisect_left, bisect_right

from . import einsetzen, formel
from .analyse import Finding, Sheet, _checked, address, said
from .sprache import T
from .workbook import Layout
from .xlsx import Cell

Pos = tuple[int, int]
MAX_FOUND = 10
MAX_LISTED = 3
# Die Funktion und ihre Nummer in TEILERGEBNIS, die ausgeblendete Zeilen weglässt.
_CODES = {"AVERAGE": 101, "COUNT": 102, "COUNTA": 103, "MAX": 104, "MIN": 105, "PRODUCT": 106, "STDEV": 107,
          "STDEV.S": 107, "STDEV.P": 108, "SUM": 109, "VAR": 110, "VAR.S": 110, "VAR.P": 111}


class _Seen:
    """Eine Formel, die mit den ausgeblendeten Zeilen etwas anderes ergibt als ohne sie."""

    def __init__(self, pos: Pos, function: str, rows: list[int], away: bool, visible: float, shown: float,
                 sub: str) -> None:
        self.pos, self.function, self.rows, self.away = pos, function, rows, away
        self.visible, self.shown, self.sub = visible, shown, sub


def checks(cells: dict[Pos, Cell], plan: Layout, name: str | None, book: dict[str, dict[Pos, Cell]]
           ) -> tuple[list[Finding], set[str]]:
    """Die Hinweise und die Zellen, deren Summe über weggefilterte Zeilen hier schon erklärt ist (auswertung.py sagt
    dazu sonst nur, dass die Tabelle gefiltert ist)."""
    hidden = getattr(cells, "hidden_rows", None)
    if not hidden:
        return [], set()
    sheet = Sheet(cells, plan, name, book)
    quiet = getattr(cells, "grouped_rows", None) or frozenset()
    filtered = sheet.filtered_rows or frozenset()
    order = sorted(hidden - quiet)
    columns: dict[int, list[int]] = {}
    for (r, c), value in sheet.inputs.items():
        if value not in (None, ""):
            columns.setdefault(c, []).append(r)
    for rows in columns.values():
        rows.sort()
    calc, _ = sheet.evaluator()
    found: list[_Seen] = []
    covered: set[str] = set()
    for pos, node in sorted(sheet.nodes.items(), key=lambda item: (item[0][1], item[0][0])):
        if node is None or node.kind != "call" or node.value not in _CODES or pos[0] in hidden or not node.args:
            continue
        if not all(a.kind in ("range", "ref") for a in node.args):
            continue
        areas = [a.value if a.kind == "range" else (*a.value, *a.value) for a in node.args]
        rows = _counted(sheet, areas, order)
        if not rows:
            continue
        if any(r in filtered for r in rows):
            covered.add(address(pos))
        if not _visible(columns, areas, hidden):
            continue  # Eine Summe nur über die ausgeblendeten Zeilen ist so gemeint.
        seen = _seen(sheet, calc, pos, node, rows, filtered)
        if seen is not None:
            found.append(seen)
    return _hints(sheet, _last_per_column(found)), covered


def _counted(sheet: Sheet, areas: list[tuple[int, int, int, int]], order: list[int]) -> list[int]:
    """Die ausgeblendeten Zeilen in den Bereichen, in denen etwas steht."""
    rows = set()
    for r1, c1, r2, c2 in areas:
        for r in order[bisect_left(order, r1):bisect_right(order, r2)]:
            if any(sheet.inputs.get((r, c)) not in (None, "") for c in range(c1, c2 + 1)):
                rows.add(r)
    return sorted(rows)


def _visible(columns: dict[int, list[int]], areas: list[tuple[int, int, int, int]], hidden: frozenset[int]) -> bool:
    """Steht in den Bereichen auch in einer sichtbaren Zeile etwas?"""
    for r1, c1, r2, c2 in areas:
        for c in range(c1, c2 + 1):
            rows = columns.get(c, [])
            if any(r not in hidden for r in rows[bisect_left(rows, r1):bisect_right(rows, r2)]):
                return True
    return False


def _seen(sheet: Sheet, calc, pos: Pos, node, rows: list[int], filtered: frozenset[int]) -> _Seen | None:
    formula = sheet.cells[pos].formula or ""
    if "(" not in formula:
        return None
    # Nur Bezüge in der Formel: Ohne Leerzeichen bleibt sie dieselbe, und so schreibt Einsetzen sie zurück.
    args = re.sub(r"\s+", "", formula[formula.index("(") + 1:])
    sub = f"SUBTOTAL({_CODES[node.value]},{args}"
    try:
        whole = calc.run(node)
        visible = calc.run(formel.parse(sub))
    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError, IndexError, KeyError):
        return None
    if not isinstance(whole, float) or not isinstance(visible, float) or isinstance(whole, bool):
        return None
    if abs(whole - visible) <= 1e-9 * max(1.0, abs(whole)):
        return None  # In den ausgeblendeten Zeilen steht nur, was nichts ändert.
    shown = sheet.inputs.get(pos)
    shown = shown if isinstance(shown, float) and not isinstance(shown, bool) else whole
    return _Seen(pos, node.value, rows, all(r in filtered for r in rows), visible, shown, sub)


def _last_per_column(found: list[_Seen]) -> list[_Seen]:
    """Eine laufende Summe je Zeile meldet sich nur einmal, mit ihrer untersten Zelle. Summe, Mittelwert und Anzahl
    untereinander unter der Liste bleiben alle."""
    by_column: dict[tuple[int, str], list[_Seen]] = {}
    for item in found:
        by_column.setdefault((item.pos[1], item.function), []).append(item)
    out = []
    for items in by_column.values():
        out += items if len(items) <= MAX_LISTED else [max(items, key=lambda i: i.pos[0])]
    return sorted(out, key=lambda i: (i.pos[0], i.pos[1]))


def _hints(sheet: Sheet, found: list[_Seen]) -> list[Finding]:
    """Ein Hinweis je Gruppe von Formeln mit denselben ausgeblendeten Zeilen (Summe, Mittelwert und Anzahl unter einer
    Liste, oder die Summen ihrer Spalten nebeneinander)."""
    groups: dict[tuple, list[_Seen]] = {}
    for item in found:
        groups.setdefault((tuple(item.rows), item.away), []).append(item)
    out = []
    for (rows, away), items in groups.items():
        first, rows = items[0], list(rows)
        kind = "weggefiltert" if away else "ausgeblendet"
        english = "filtered-out" if away else "hidden"
        which = T(f"die {kind}e Zeile {rows[0]}", f"the {english} row {rows[0]}") if len(rows) == 1 else \
            T(f"die {len(rows)} {kind}en Zeilen ({_rows(rows)})", f"the {len(rows)} {english} rows ({_rows(rows)})") if away else \
            T(f"die {kind}en Zeilen {_rows(rows)}", f"the {english} rows {_rows(rows)}")
        seen = T("Was der Filter zeigt", "What the filter shows") if away else T("Was man sieht", "What you can see")
        cells = [address(i.pos) for i in items]
        also = T(f" Ebenso {_listed(cells[1:])}.", f" The same goes for {_listed(cells[1:])}.") if len(items) > 1 else ""
        text = T(f"{sheet.label(first.pos)} zählt auch {which} mit: {seen}, ergibt {said(first.visible)}, "
                 f"{cells[0]} zeigt {said(first.shown)}.{also}",
                 f"{sheet.label(first.pos)} also counts {which}: {seen[0].lower()}{seen[1:]} adds up to {said(first.visible)} "
                 f"while {cells[0]} shows {said(first.shown)}.{also}")
        fix = einsetzen.formulas(sheet, {i.pos: i.sub for i in items})
        checked = _checked(sheet, first.sub)
        tip = T(f"Soll {cells[0]} nur zählen, was zu sehen ist: Korrigiert in {cells[0]}: {checked}.",
                f"If {cells[0]} should count only what is visible: Fixed in {cells[0]}: {checked}.") if checked and fix else \
            T(f"Soll {cells[0]} nur zählen, was zu sehen ist, nimm TEILERGEBNIS mit 101 bis 111.",
              f"If {cells[0]} should count only what is visible, use SUBTOTAL with 101 to 111.")
        tip += T(" Sollen alle Zeilen zählen, passt es so.", " If all rows should count, it is fine as it is.")
        out.append(Finding(f"{'filter' if away else 'ausgeblendet'}:{cells[0]}", text, "pruefen" if away else "info",
                           tip, fix=fix, cell=",".join(cells)))
        if len(out) >= MAX_FOUND:
            break
    return out


def _listed(items: list[str]) -> str:
    if len(items) > MAX_LISTED:
        return ", ".join(items[:MAX_LISTED]) + T(f" und {len(items) - MAX_LISTED} weitere", f" and {len(items) - MAX_LISTED} more")
    return items[0] if len(items) == 1 else ", ".join(items[:-1]) + T(f" und {items[-1]}", f" and {items[-1]}")


def _rows(rows: list[int]) -> str:
    """8 bis 10; 12, 15 und 18; 12, 15, 18 und 4 weitere."""
    runs: list[tuple[int, int]] = []
    for r in rows:
        if runs and r == runs[-1][1] + 1:
            runs[-1] = (runs[-1][0], r)
        else:
            runs.append((r, r))
    parts: list[tuple[str, int]] = []
    for a, b in runs:
        parts += [(T(f"{a} bis {b}", f"{a} to {b}"), b - a + 1)] if b - a >= 2 else [(str(r), 1) for r in range(a, b + 1)]
    if len(parts) > MAX_LISTED:
        rest = sum(n for _, n in parts[MAX_LISTED:])
        return ", ".join(p for p, _ in parts[:MAX_LISTED]) + T(f" und {rest} weitere", f" and {rest} more")
    shown = [p for p, _ in parts]
    return shown[0] if len(shown) == 1 else ", ".join(shown[:-1]) + T(f" und {shown[-1]}", f" and {shown[-1]}")
