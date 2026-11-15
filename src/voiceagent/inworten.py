"""Mappe in Worten: was die ganze Mappe rechnet, von den Daten bis zu den Ergebnissen.

Für den Reiter „Mappe“ im Seitenfenster und den Prüfbericht. Ohne KI: Die Sätze kommen aus den Spaltenköpfen, den
Formeln (mit den Namen ihrer Spalten, dazu bereiche.in_words: welche Zeilen eine Formel nimmt) und dem eigenen
Formelrechner, der jede Formel mit den Werten aus Excel nachrechnet. Was Gegenprobe nicht in Worte fassen kann, steht
als Formel da, nicht als geratener Satz.
"""

from __future__ import annotations

import re
import time
from collections import Counter, deque
from dataclasses import dataclass, field

from . import auswertung, formel, herleitung
from .analyse import Sheet as Calc
from .analyse import _german, address
from .bereiche import _german_formula
from .formel import Node
from .sprache import T, english
from .workbook import Layout, layout, parts_of, pattern
from .xlsx import ANY_REF, Sheet, area_text, column_number, number_text, outside_literals

Pos = tuple[int, int]
Area = tuple[int, int, int, int]

MAX_SHEETS = 40
MAX_FIGURES = 30
MAX_RESULTS = 15
MAX_LISTED = 4
# So viele Formeln rechnet die Mappe in Worten höchstens nach, und so viele Zellen besucht die Suche nach den Eingaben
# eines Ergebnisses höchstens.
MAX_CHECKS = 5000
MAX_WALK = 20000
BUDGET = 8.0
# Ein einzelner Bezug (kein Bereich, kein anderes Blatt).
_SINGLE = re.compile(r"(?<![A-Za-z0-9_.$\\:!])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(:])")


@dataclass
class _Info:
    name: str
    sheet: Sheet
    plan: Layout
    calc: Calc
    rows: list[int]
    # Spalte der Tabelle → „eingabe“ oder „formel“.
    columns: dict[int, str]
    # Die Formeln der gerechneten Spalten (je Spalte eine Beschreibung, nicht je Zelle).
    computed: set[Pos]
    # Anderes Blatt → (Formeln, die dort lesen, Funktionen, gelesene Spalten).
    reads: dict[str, tuple[set[Pos], Counter, list[str]]] = field(default_factory=dict)
    # Bereiche dieses Blatts, die Formeln auf anderen Blättern lesen.
    read_from: list[tuple[str, Area]] = field(default_factory=list)


class _Checks:
    """Nachrechnen mit dem eigenen Formelrechner (wie die Herleitung), höchstens MAX_CHECKS Formeln."""

    def __init__(self, cells_of: dict[str, dict], infos: dict[str, _Info]) -> None:
        self.book = herleitung.Book(cells_of)
        # Die Formeln sind schon gelesen: nicht noch einmal.
        self.book._calcs.update({name: info.calc for name, info in infos.items()})
        self.left = MAX_CHECKS
        self.deadline = time.monotonic() + BUDGET
        self.skipped = 0
        # Bei sehr vielen Formeln in gerechneten Spalten: nur jede stride-te nachrechnen, gleichmäßig verteilt.
        many = sum(len(info.computed) for info in infos.values())
        self.total = sum(1 for info in infos.values() for c in info.sheet.cells.values() if c.formula)
        # Die einzelnen Formeln (Kennzahlen) immer, die Spalten mit dem, was übrig bleibt.
        room = max(MAX_CHECKS // 2, MAX_CHECKS - min(self.total - many, MAX_FIGURES * len(infos)))
        self.stride = max(1, -(-many // room))
        # Wie viele Formeln nachgerechnet dasselbe ergeben, etwas anderes, und welche Gegenprobe nicht nachrechnen kann.
        self.counts: Counter = Counter()
        self.wrong: set[tuple[str, Pos]] = set()

    def check(self, info: _Info, pos: Pos) -> tuple[str, str]:
        """(„ok“ | „anders“ | „unbekannt“, unser Wert, wenn anders)."""
        state, value = self._check(info, pos)
        if state == "zeit":
            self.skipped += 1
            return "unbekannt", ""
        self.counts[state] += 1
        if state == "anders":
            self.wrong.add((info.name, pos))
        return state, value

    def _check(self, info: _Info, pos: Pos) -> tuple[str, str]:
        node = info.calc.nodes.get(pos)
        if node is None or pos in info.calc.unknown:
            return "unbekannt", ""
        if self.left <= 0 or time.monotonic() > self.deadline:
            return "zeit", ""
        self.left -= 1
        ours, _ = self.book.run(info.name, node)
        state, value = herleitung._check(info.calc.shown.get(pos), ours)
        if state == "anders" and isinstance(ours, float):
            value = herleitung._like(info.sheet.cells[pos].value, ours)
        return state, value


def describe(book: str, sheets: list[Sheet], hints: dict[str, dict] | None = None) -> dict:
    """Die Mappe in Worten. sheets: alle Blätter mit Zellen (und Pivot-Tabellen, Diagrammen …), hints: je Blatt die Zahl
    der Hinweise ({"fehler": 2, "pruefen": 1}), soweit schon geprüft."""
    sheets = [s for s in sheets if s.cells or s.charts or s.pivots][:MAX_SHEETS]
    if not sheets:
        return {"state": "empty", "book": book, "text": T("In dieser Mappe steht noch nichts.", "This workbook is still empty.")}
    cells_of = {s.name: s.cells for s in sheets}
    infos = {s.name: _info(s, cells_of) for s in sheets}
    names = {n.casefold(): n for n in infos}
    for info in infos.values():
        _links(info, infos, names)
    checks = _Checks(cells_of, infos)
    order = _order(infos)
    described = [_sheet(infos[n], infos, checks, (hints or {}).get(n)) for n in order]
    taint = _taint(infos, checks.wrong)
    results = _results(infos, order, described, taint)
    return {
        "state": "ok",
        "book": book,
        "summary": _summary(infos, order, described, results, checks, taint),
        "flow": _flow(infos, order),
        "results": results,
        "sheets": described,
    }


# --- Blätter ---


def _info(sheet: Sheet, cells_of: dict[str, dict]) -> _Info:
    plan = layout(sheet)
    calc = Calc(sheet.cells, plan, sheet.name, cells_of)
    rows = auswertung._data_rows(calc) if plan.header_row is not None and plan.headers else []
    columns: dict[int, str] = {}
    computed: set[Pos] = set()
    for col in sorted(plan.headers):
        filled = [(r, col) for r in rows if (c := sheet.cells.get((r, col))) is not None and (c.formula or c.value)]
        # Ein paar Kennzahlen neben der Tabelle („Mittelwert Alter“ in K2:L3) sind keine Spalte der Tabelle.
        if not filled or len(filled) * 4 < len(rows):
            continue
        formulas = [p for p in filled if sheet.cells[p].formula]
        columns[col] = "formel" if formulas and len(formulas) * 2 >= len(filled) else "eingabe"
        if columns[col] == "formel":
            computed |= set(formulas)
    return _Info(sheet.name, sheet, plan, calc, rows, columns, computed)


def _foreign(node: Node, function: str = ""):
    """Bezüge auf andere Blätter: (Blatt, Zelle oder Bereich, die Funktion, in der sie stehen)."""
    if node.kind in ("xref", "xrange"):
        yield node.value[0], node.value[1], function
    for child in node.args:
        yield from _foreign(child, node.value if node.kind == "call" else function)


def _links(info: _Info, infos: dict[str, _Info], names: dict[str, str]) -> None:
    for pos, node in info.calc.nodes.items():
        if node is None:
            continue
        for other, ref, function in _foreign(node):
            target = names.get(other.casefold())
            if target is None or target == info.name:
                continue
            area = ref if len(ref) == 4 else (ref[0], ref[1], ref[0], ref[1])
            cells, functions, columns = info.reads.setdefault(target, (set(), Counter(), []))
            cells.add(pos)
            if function:
                functions[T(_german(function), function)] += 1
            for name in _columns(infos[target], area):
                if name not in columns:
                    columns.append(name)
            infos[target].read_from.append((info.name, area))


def _columns(info: _Info, area: Area) -> list[str]:
    """Die Spalten der Tabelle, die ein Bereich trifft, sonst der Bereich selbst."""
    r1, c1, r2, c2 = area
    plan = info.plan
    if plan.header_row is not None and r2 > plan.header_row and r1 <= plan.last_row + 1:
        found = [plan.headers[c] for c in sorted(plan.headers) if c1 <= c <= c2]
        if found:
            return found
    if (r1, c1) == (r2, c2):
        return [info.calc.label((r1, c1))]
    return [_where(area)]


def _where(area: Area) -> str:
    r1, c1, r2, c2 = area
    if r1 == 1 and r2 >= formel.MAX_ROW:
        return f"{address((1, c1))[:-1]}:{address((1, c2))[:-1]}"
    return area_text(area)


def _order(infos: dict[str, _Info]) -> list[str]:
    """Die Blätter so, dass ein Blatt nach denen kommt, aus denen es liest (sonst in der Reihenfolge der Mappe)."""
    names = list(infos)
    waiting = {n: {o for o in infos[n].reads if o != n} for n in names}
    order: list[str] = []
    while len(order) < len(names):
        ready = [n for n in names if n not in order and not (waiting[n] - set(order))]
        # Ein Kreis zwischen Blättern: das nächste in der Reihenfolge der Mappe.
        order.append(ready[0] if ready else next(n for n in names if n not in order))
    return order


def _sheet(info: _Info, infos: dict[str, _Info], checks: _Checks, hints: dict | None) -> dict:
    plan, cells = info.plan, info.sheet.cells
    formulas = [p for p, c in cells.items() if c.formula]
    out: dict = {"name": info.name, "cells": len(cells), "formulas": len(formulas)}
    if info.rows and info.columns:
        cols = sorted(info.columns)
        out["table"] = {
            "area": f"{address((plan.header_row, cols[0]))}:{address((info.rows[-1], cols[-1]))}",
            "rows": len(info.rows),
            "inputs": [plan.headers[c] for c in cols if info.columns[c] == "eingabe"],
            "computed": [plan.headers[c] for c in cols if info.columns[c] == "formel"],
        }
    out["columns"] = [_computed(info, col, checks) for col in sorted(info.columns) if info.columns[col] == "formel"]
    figures, more = _figures(info, checks)
    out["figures"], out["figures_more"] = figures, more
    out["reads"] = [{"sheet": other, "columns": columns[:8], "functions": [f for f, _ in functions.most_common(3)],
                     "formulas": len(cells_read)} for other, (cells_read, functions, columns) in info.reads.items()]
    readers = []
    for other in infos.values():
        if other.name != info.name and info.name in other.reads and other.name not in readers:
            readers.append(other.name)
    out["read_by"] = readers
    parts = parts_of([info.sheet], {info.name: plan})
    out["parts"] = [{"label": p.label, "text": p.text} for p in parts.values() if not p.key.startswith("name:")][:12]
    if hints:
        out["hints"] = {"fehler": hints.get("fehler") or 0, "pruefen": hints.get("pruefen") or 0}
    return out


def _computed(info: _Info, col: int, checks: _Checks) -> dict:
    """Eine gerechnete Spalte: ihre Formel in Worten, in wie vielen Zeilen sie so steht, und ob sie nachgerechnet stimmt."""
    cells, plan = info.sheet.cells, info.plan
    filled = [(r, col) for r in info.rows if (c := cells.get((r, col))) is not None and (c.formula or c.value)]
    formulas = [p for p in filled if cells[p].formula]
    # Dieselbe Formel: heruntergezogen (gleich relativ zur Zelle) oder in jede Zeile von Hand geschrieben mit demselben
    # Bezug unter der Tabelle (=B3/B6, =B4/B6 …). Es zählt, was mehr Zeilen zusammenfasst.
    rows = set(info.rows)
    groups = []
    for shape in (pattern, lambda formula, here: _fixed(formula, here, rows)):
        keys = {p: shape(cells[p].formula, p) for p in formulas}
        main = Counter(keys.values()).most_common(1)[0][0]
        groups.append([p for p in formulas if keys[p] == main])
    same = groups[0] if len(groups[0]) >= len(groups[1]) else groups[1]
    first = same[0]
    alike = set(same)
    odd = [(p, T("feste Zahl", "fixed number") if not cells[p].formula else T("andere Formel", "different formula"))
           for p in filled if p not in alike]
    wrong, unknown = [], 0
    for p in formulas[::checks.stride]:
        state, _ = checks.check(info, p)
        if state == "anders":
            wrong.append(address(p))
        unknown += state == "unbekannt"
    out = {
        "name": plan.headers[col],
        "area": f"{address(filled[0])}:{address(filled[-1])}",
        "cell": address(first),
        "formula": "=" + _german_formula(cells[first].typed or cells[first].formula),
        "words": info.calc.readable(first),
        "rows": len(filled),
        "same": len(same),
        "odd": [f"{address(p)} ({why})" for p, why in odd[:MAX_LISTED]],
        "odd_more": max(0, len(odd) - MAX_LISTED),
        "wrong": wrong[:MAX_LISTED],
        "wrong_count": len(wrong),
        "check": "anders" if wrong else ("unbekannt" if unknown else "ok"),
        "constants": _constants(info.calc.nodes.get(first)),
    }
    if _today(info.calc.nodes.get(first)):
        out["today"] = True
    moving = _moving(info, first, cells[first].formula, same[-1]) if same is groups[0] and len(groups[1]) < len(same) else ""
    if moving:
        out["moving"] = moving
    out["text"] = _column_text(out)
    return out


def _fixed(formula: str, here: Pos, rows: set[int]) -> str:
    """Die Formel wie pattern(), aber ein Bezug auf eine Zeile außerhalb der Daten gilt als fest (wie mit $)."""

    def relative(match: re.Match) -> str:
        dollar_col, letters, dollar_row, digits = match.groups()
        row = int(digits)
        col = f"S{column_number(letters)}" if dollar_col else f"S[{column_number(letters) - here[1]}]"
        return (f"Z{row}" if dollar_row or row not in rows else f"Z[{row - here[0]}]") + col

    return outside_literals(formula.replace(" ", ""), lambda part: ANY_REF.sub(relative, part))


def _moving(info: _Info, first: Pos, formula: str, last: Pos) -> str:
    """Ein Bezug ohne $ auf eine Zelle außerhalb der Datenzeilen (D9 unter der Tabelle): Heruntergezogen wandert er
    mit, E3 nimmt D10. Das beschreibt, was die Formel tut; ob es ein Fehler ist, sagen die Hinweise."""
    if last == first:
        return ""
    rows = set(info.rows)
    found: list[Pos] = []

    def look(part: str) -> str:
        for match in _SINGLE.finditer(part):
            _, letters, dollar_row, digits = match.groups()
            row = int(digits)
            if not dollar_row and row not in rows and row != first[0]:
                found.append((row, column_number(letters)))
        return part

    outside_literals(formula, look)
    if not found:
        return ""
    row, col = found[0]
    second = next(r for r in info.rows if r > first[0])
    step, end = second - first[0], last[0] - first[0]
    text = T(f"Der Bezug auf {info.calc.label((row, col))} hat kein $ vor der Zeile und wandert beim Herunterziehen mit: "
             f"{address((second, first[1]))} nimmt {address((row + step, col))}",
             f"The reference to {info.calc.label((row, col))} has no $ before the row and moves when filled down: "
             f"{address((second, first[1]))} takes {address((row + step, col))}")
    if last[0] != second:
        text += T(f", {address(last)} schon {address((row + end, col))}", f", {address(last)} already {address((row + end, col))}")
    return text + "."


def _today(node: Node | None) -> bool:
    return node is not None and any(c.value in _VOLATILE for c in formel.calls(node))


def _constants(node: Node | None) -> list[str]:
    if node is None:
        return []
    return [number_text(v) for v in sorted(set(formel.literals(node)) - {0.0, 1.0})][:MAX_LISTED]


def _loose(info: _Info) -> list[list[Pos]]:
    """Die übrigen Formeln, untereinander gleiche zusammengefasst (J12:J15 mit derselben Formel)."""
    cells = info.sheet.cells
    rest = sorted((p for p, c in cells.items() if c.formula and p not in info.computed), key=lambda p: (p[1], p[0]))
    groups: list[list[Pos]] = []
    for p in rest:
        last = groups[-1][-1] if groups else None
        if last is not None and last[1] == p[1] and last[0] + 1 == p[0] and \
                pattern(cells[last].formula, last) == pattern(cells[p].formula, p):
            groups[-1].append(p)
        else:
            groups.append([p])
    return sorted(groups, key=lambda g: g[0])


def _figures(info: _Info, checks: _Checks) -> tuple[list[dict], int]:
    groups = _loose(info)
    found = [_figure(info, group, checks) for group in groups[:MAX_FIGURES]]
    return found, max(0, len(groups) - MAX_FIGURES)


def _figure(info: _Info, group: list[Pos], checks: _Checks) -> dict:
    cells, calc = info.sheet.cells, info.calc
    first = group[0]
    label = calc.figure_label(first)
    # „B45 (Wert je Aktie (€))“: nur die äußere Klammer weg, die Einheit bleibt.
    rest = label[len(address(first)):].strip()
    name = rest[1:-1].strip() if rest.startswith("(") and rest.endswith(")") else rest.strip(" ()")
    words = auswertung.in_words(calc, first) if len(group) == 1 else ""
    if not words:
        readable = calc.readable(first)
        german = "=" + _german_formula(cells[first].typed or cells[first].formula)
        # Mit Zeilennamen („Alter 24100 bis 24136“) hilft die lesbare Formel nicht weiter.
        ranges = (" bis ", " to ") if english() else (" bis ",)
        if readable.replace(" ", "") != german.replace(" ", "") and not any(word in readable for word in ranges):
            words = readable
    states = [checks.check(info, p) for p in group]
    wrong = [(address(p), ours) for p, (state, ours) in zip(group, states) if state == "anders"]
    out = {
        "cell": address(first) if len(group) == 1 else f"{address(first)}:{address(group[-1])}",
        "name": name,
        "value": cells[first].value,
        "formula": "=" + _german_formula(cells[first].typed or cells[first].formula),
        "words": words,
        "check": "anders" if wrong else ("ok" if all(s == "ok" for s, _ in states) else "unbekannt"),
        "constants": _constants(calc.nodes.get(first)),
    }
    if _today(calc.nodes.get(first)):
        out["today"] = True
    if len(group) > 1:
        out["count"] = len(group)
    if wrong:
        out["wrong"] = [T(f"{cell} ergibt {ours}", f"{cell} gives {ours}") if ours else cell for cell, ours in wrong[:MAX_LISTED]]
    return out


# --- Die ganze Mappe ---


def _flow(infos: dict[str, _Info], order: list[str]) -> list[dict]:
    """Welches Blatt aus welchem liest, mit den Funktionen dafür (SVERWEIS, direkte Bezüge)."""
    flow = []
    for name in order:
        for other, (cells, functions, columns) in infos[name].reads.items():
            flow.append({"from": other, "to": name, "formulas": len(cells), "columns": columns[:6],
                         "functions": [f for f, _ in functions.most_common(3)]})
    return flow


def _read_elsewhere(infos: dict[str, _Info], info: _Info, pos: Pos) -> bool:
    return any(a[0] <= pos[0] <= a[2] and a[1] <= pos[1] <= a[3] for _, a in info.read_from)


def _results(infos: dict[str, _Info], order: list[str], described: list[dict], taint: dict) -> list[dict]:
    """Die Ergebnisse: Formeln, die keine andere Formel der Mappe weiterverwendet, mit den Eingaben, aus denen sie
    entstehen. Eine gerechnete Spalte, die niemand weiterverwendet, ist als Ganzes ein Ergebnis."""
    by_name = {d["name"]: d for d in described}
    found = []
    for name in order:
        info = infos[name]
        sheet = by_name[name]
        used = info.calc.dependents
        for figure in sheet["figures"]:
            cells = _cells_of(figure["cell"])
            if any(used.get(p) or _read_elsewhere(infos, info, p) for p in cells):
                continue
            found.append({"sheet": name, "kind": "kennzahl", "cell": figure["cell"], "name": figure["name"],
                          "value": figure["value"] if len(cells) == 1 else "", "check": figure["check"],
                          **_sources(infos, name, cells), **_upstream(taint, name, cells)})
        for column in sheet["columns"]:
            cells = _cells_of(column["area"])
            if any(used.get(p) or _read_elsewhere(infos, info, p) for p in cells):
                continue
            found.append({"sheet": name, "kind": "spalte", "cell": column["area"], "name": column["name"], "value": "",
                          "check": column["check"],
                          **_sources(infos, name, [formel.position(column["cell"])]),
                          **_upstream(taint, name, cells)})
    for item in found:
        item["text"] = _result_text(item)
    return found[:MAX_RESULTS]


def _result_text(item: dict) -> list[str]:
    """Woraus ein Ergebnis entsteht, in Sätzen."""
    lines = []
    groups = []
    for group in item["inputs"]:
        names = group["names"] + ([T(f"{group['more']} weiteren", f"{group['more']} more")] if group["more"] else [])
        own = group["sheet"] == item["sheet"]
        groups.append(_listing(names) + ("" if own else T(f" („{group['sheet']}“)", f" (“{group['sheet']}”)")))
    if groups:
        joined = groups[0] if len(groups) == 1 else T(f"{', '.join(groups[:-1])} sowie {groups[-1]}",
                                                      f"{', '.join(groups[:-1])} as well as {groups[-1]}")
        text = T("Entsteht aus ", "Comes from ") + joined
        if item["via"]:
            text += T(f", gerechnet über {_listing(item['via'])}", f", calculated via {_listing(item['via'])}")
        lines.append(text + ".")
    elif item["via"]:
        lines.append(T(f"Gerechnet über {_listing(item['via'])}.", f"Calculated via {_listing(item['via'])}."))
    else:
        lines.append(T("Entsteht nur aus festen Zahlen in der Formel.", "Comes only from fixed numbers in the formula."))
    if item["constants"]:
        lines.append(T(f"Feste Zahlen unterwegs: {_listing(item['constants'])}.",
                       f"Fixed numbers along the way: {_listing(item['constants'])}."))
    if item.get("upstream"):
        more = item["upstream_count"] - len(item["upstream"])
        cells = _listing(item["upstream"] + ([T(f"{more} weitere", f"{more} more")] if more else []))
        # Das Seitenfenster (pruefer.js) zeigt Zeilen, die mit „Unterwegs“ / „On the way“ beginnen, als Fehler.
        lines.append(T(f"Unterwegs ergibt nachgerechnet {cells} etwas anderes als in Excel: Dann stimmt womöglich auch "
                       "dieses Ergebnis nicht.",
                       f"On the way, {cells} {'gives' if item['upstream_count'] == 1 else 'give'} something different "
                       "from Excel when recalculated: then this result may be wrong too."))
    if item.get("today"):
        lines.append(T("Hängt vom heutigen Datum ab (HEUTE()): Später geöffnet kann etwas anderes herauskommen, ohne dass sich "
                       "eine Zahl in der Mappe ändert.",
                       "Depends on today's date (TODAY()): opened later, it can give a different result without any "
                       "number in the workbook changing."))
    if item.get("cut"):
        lines.append(T("Der Weg zurück ist sehr lang, ich nenne nur, was ich bis hierher gefunden habe.",
                       "The path back is very long; I only list what I found up to this point."))
    return lines


def _column_text(column: dict) -> list[str]:
    lines = []
    if column["same"] == column["rows"]:
        lines.append(T(f"In allen {column['rows']} Zeilen dieselbe Formel.", f"The same formula in all {column['rows']} rows.")
                     if column["rows"] > 1 else "")
    else:
        odd = column["odd"] + ([T(f"{column['odd_more']} weitere", f"{column['odd_more']} more")] if column["odd_more"] else [])
        lines.append(T(f"In {column['same']} von {column['rows']} Zeilen dieselbe Formel, anders: {_listing(odd)}.",
                       f"The same formula in {column['same']} of {column['rows']} rows, different: {_listing(odd)}."))
    if column.get("moving"):
        lines.append(column["moving"])
    if column["wrong_count"]:
        more = column["wrong_count"] - len(column["wrong"])
        wrong = _listing(column["wrong"] + ([T(f"{more} weitere", f"{more} more")] if more else []))
        # Das Seitenfenster (pruefer.js) zeigt Zeilen, die mit „Nachgerechnet anders“ / „Recalculated differently“
        # beginnen, als Fehler.
        lines.append(T(f"Nachgerechnet anders als in Excel: {wrong}.", f"Recalculated differently from Excel: {wrong}."))
    if column["constants"]:
        lines.append(T(f"Feste Zahlen in der Formel: {_listing(column['constants'])}.",
                       f"Fixed numbers in the formula: {_listing(column['constants'])}."))
    if column.get("today"):
        lines.append(T("Hängt vom heutigen Datum ab (HEUTE()).", "Depends on today's date (TODAY())."))
    return [line for line in lines if line]


def _taint(infos: dict[str, _Info], wrong: set[tuple[str, Pos]]) -> dict[tuple[str, Pos], list[tuple[str, Pos]]]:
    """Je Zelle die Formeln vor ihr (auch auf anderen Blättern), die nachgerechnet etwas anderes ergeben als in Excel:
    Mit ihnen stimmt womöglich auch, was aus ihnen folgt, nicht."""
    readers: dict[str, list[tuple[Area, tuple[str, Pos]]]] = {}
    names = {n.casefold(): n for n in infos}
    for info in infos.values():
        for pos, node in info.calc.nodes.items():
            if node is None:
                continue
            for other, ref, _ in _foreign(node):
                target = names.get(other.casefold())
                if target is not None:
                    area = ref if len(ref) == 4 else (ref[0], ref[1], ref[0], ref[1])
                    readers.setdefault(target, []).append((area, (info.name, pos)))
    origin: dict[tuple[str, Pos], list[tuple[str, Pos]]] = {}
    queue = deque((w, w) for w in sorted(wrong))
    steps = 0
    while queue and steps < MAX_WALK:
        (where, pos), source = queue.popleft()
        steps += 1
        after = [(where, p) for p in infos[where].calc.dependents.get(pos, ())]
        after += [reader for a, reader in readers.get(where, ()) if a[0] <= pos[0] <= a[2] and a[1] <= pos[1] <= a[3]]
        for item in after:
            found = origin.setdefault(item, [])
            if source in found or len(found) >= MAX_LISTED:
                continue
            found.append(source)
            queue.append((item, source))
    return origin


def _upstream(taint: dict, sheet: str, cells: list[Pos]) -> dict:
    own = {(sheet, p) for p in cells}
    found: list[tuple[str, Pos]] = []
    for p in cells:
        for source in taint.get((sheet, p), ()):
            if source not in own and source not in found:
                found.append(source)
    if not found:
        return {}
    names = [address(p) if where == sheet else f"{where}!{address(p)}" for where, p in found]
    return {"upstream": names[:MAX_LISTED], "upstream_count": len(names)}


def _cells_of(text: str) -> list[Pos]:
    if ":" not in text:
        return [formel.position(text)]
    r1, c1, r2, c2 = formel.area(text)
    return [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]


def _sources(infos: dict[str, _Info], sheet: str, start: list[Pos]) -> dict:
    """Woraus ein Ergebnis entsteht: die Eingaben (je Blatt die Spalten oder Zellen), die gerechneten Spalten und
    Kennzahlen dazwischen (die nächsten zuerst) und die festen Zahlen in den Formeln unterwegs."""
    names = {n.casefold(): n for n in infos}
    columns: dict[str, set[int]] = {}
    loose: dict[str, list[Pos]] = {}
    via: list[str] = []
    today = False
    constants: set[float] = set()
    queue = deque((sheet, p) for p in start)
    first = set(queue)
    seen: set[tuple[str, Pos]] = set()
    # Eine Zelle einer gerechneten Spalte steht für alle mit derselben Formel: Ihre Vorgänger liegen in denselben
    # Spalten, nur in anderen Zeilen. So bleibt der Weg auch bei 20 000 Zeilen kurz.
    expanded: set[tuple[str, int, str]] = set()
    steps = 0
    cut = False
    while queue:
        if steps >= MAX_WALK:
            cut = True
            break
        where, pos = queue.popleft()
        if (where, pos) in seen:
            continue
        seen.add((where, pos))
        info = infos[where]
        cell = info.sheet.cells.get(pos)
        if cell is None:
            continue
        if cell.formula is None:
            if cell.value == "":
                continue
            if info.plan.in_table(*pos) and pos[0] in info.rows:
                found = columns.setdefault(where, set())
                steps += pos[1] not in found
                found.add(pos[1])
            else:
                loose.setdefault(where, []).append(pos)
                steps += 1
            continue
        node = info.calc.nodes.get(pos)
        if node is None:
            continue
        if pos in info.computed:
            shape = (where, pos[1], pattern(cell.formula, pos))
            if shape in expanded:
                continue
            expanded.add(shape)
        steps += 1
        if (where, pos) not in first:
            _add(via, _via_name(info, pos, where != sheet))
        constants |= set(formel.literals(node)) - {0.0, 1.0}
        today = today or _today(node)
        queue += [(where, p) for p in sorted(info.calc.precedents.get(pos, ()), key=lambda p: (p[1], p[0]))]
        for other, ref, _ in _foreign(node):
            target = names.get(other.casefold())
            if target is None:
                continue
            area = ref if len(ref) == 4 else (ref[0], ref[1], ref[0], ref[1])
            queue += [(target, p) for p in _inside(infos[target], area)]
    inputs: dict[str, list[str]] = {}
    # Das eigene Blatt zuerst, dann die anderen in der Reihenfolge der Mappe; je Blatt die Spalten von links.
    for where in sorted(set(columns) | set(loose), key=lambda n: (n != sheet, list(infos).index(n))):
        found = inputs.setdefault(where, [])
        for col in sorted(columns.get(where, ())):
            _add(found, infos[where].plan.headers[col])
        for name in _runs(infos[where], loose.get(where, [])):
            _add(found, name)
    out = {"inputs": [{"sheet": s, "names": n[:8], "more": max(0, len(n) - 8)} for s, n in inputs.items()],
           "via": via[:6], "constants": [number_text(v) for v in sorted(constants)][:MAX_LISTED]}
    if today:
        out["today"] = True
    if cut:
        out["cut"] = True
    return out


def _add(items: list[str], name: str) -> None:
    if name and name not in items:
        items.append(name)


def _inside(info: _Info, area: Area) -> list[Pos]:
    r1, c1, r2, c2 = area
    if (r2 - r1 + 1) * (c2 - c1 + 1) <= len(info.sheet.cells):
        return [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1) if (r, c) in info.sheet.cells]
    return [p for p in info.sheet.cells if r1 <= p[0] <= r2 and c1 <= p[1] <= c2]


# Funktionen, deren Ergebnis sich ohne Änderung an der Mappe ändert.
_VOLATILE = {"TODAY", "NOW"}


def _runs(info: _Info, cells: list[Pos]) -> list[str]:
    """Eingaben außerhalb der Tabelle: untereinander stehende als Bereich mit dem Wort darüber („Patienten
    (B21:B32)“), einzelne mit dem Wort daneben („B1 (Zinssatz)“), Zahlen einer Pivot-Tabelle als diese."""
    names = []
    pivots = [p for p in cells if info.plan.is_derived(*p)]
    if pivots:
        names.append(T("Zahlen einer Pivot-Tabelle", "numbers from a pivot table"))
    rest = sorted((p for p in cells if p not in pivots), key=lambda p: (p[1], p[0]))
    runs: list[list[Pos]] = []
    for p in rest:
        if runs and runs[-1][-1][1] == p[1] and runs[-1][-1][0] + 1 == p[0]:
            runs[-1].append(p)
        else:
            runs.append([p])
    for run in sorted(runs, key=lambda r: r[0]):
        if len(run) == 1:
            names.append(info.calc.label(run[0]))
            continue
        above = info.sheet.cells.get((run[0][0] - 1, run[0][1]))
        where = f"{address(run[0])}:{address(run[-1])}"
        word = above.value if above is not None and above.formula is None and isinstance(info.calc.shown.get(
            (run[0][0] - 1, run[0][1])), str) else ""
        names.append(f"{word} ({where})" if word else where)
    return names


def _via_name(info: _Info, pos: Pos, elsewhere: bool) -> str:
    if pos in info.computed:
        name = info.plan.headers[pos[1]]
    else:
        label = info.calc.figure_label(pos)
        name = label if label != address(pos) else address(pos)
    return f"{info.name}: {name}" if elsewhere else name


def _summary(infos: dict[str, _Info], order: list[str], described: list[dict], results: list[dict], checks: _Checks,
             taint: dict) -> list[str]:
    """Die ersten Sätze: wie groß die Mappe ist, was nachgerechnet stimmt, wie die Blätter zusammenhängen."""
    formulas = sum(d["formulas"] for d in described)
    lines = [T(f"{len(order)} {'Blatt' if len(order) == 1 else 'Blätter'} mit Inhalt, {_count(formulas, 'Formel', 'Formeln')}.",
               f"{len(order)} {'sheet' if len(order) == 1 else 'sheets'} with content, "
               f"{_count(formulas, 'formula', 'formulas')}.")]
    tables = [T(f"„{d['name']}“ ({d['table']['rows']} Zeilen)", f"“{d['name']}” ({_count(d['table']['rows'], 'row', 'rows')})")
              for d in described if d.get("table") and d["table"]["inputs"]]
    if tables:
        lines.append(T(f"Daten stehen in {_listing(tables)}.", f"Data is in {_listing(tables)}."))
    # Zuerst die Formeln, vor denen nichts anders rechnet: Die übrigen rechnen oft nur mit deren Zahl weiter.
    sheets = list(order)
    wrong = [f"{name}!{address(pos)}" for name, pos in
             sorted(checks.wrong, key=lambda w: (w in taint, sheets.index(w[0]), w[1][1], w[1][0]))]
    ok, bad, unknown = checks.counts["ok"], checks.counts["anders"], checks.counts["unbekannt"]
    listed = f"{_listing(wrong[:MAX_LISTED])}{' …' if bad > MAX_LISTED else ''}" if wrong else ""
    if checks.stride > 1:
        # Sehr viele Formeln: nur eine Auswahl nachgerechnet, und das steht dabei.
        head = T(f"Nachgerechnet habe ich {ok + bad + unknown} der {checks.total} Formeln (in den gerechneten Spalten "
                 f"jede {checks.stride}.)",
                 f"I recalculated {ok + bad + unknown} of the {checks.total} formulas (in the calculated columns "
                 f"every {_ordinal(checks.stride)})")
        if bad:
            lines.append(T(f"{head}: {bad} {'ergibt' if bad == 1 else 'ergeben'} etwas anderes als in Excel ({listed}), "
                           f"{ok} dasselbe.",
                           f"{head}: {bad} {'gives' if bad == 1 else 'give'} something different from Excel ({listed}), "
                           f"{ok} the same."))
        elif ok:
            lines.append(T(f"{head}: Alle, die Gegenprobe nachrechnen kann, ergeben dasselbe wie in Excel.",
                           f"{head}: all that Gegenprobe can recalculate give the same as in Excel."))
    elif bad:
        lines.append(T(f"Nachgerechnet {'ergibt' if bad == 1 else 'ergeben'} {_count(bad, 'Formel', 'Formeln')} etwas anderes "
                       f"als in Excel ({listed}), {_count(ok, 'Formel ergibt', 'Formeln ergeben')} dasselbe.",
                       f"Recalculated, {_count(bad, 'formula gives', 'formulas give')} something different from Excel "
                       f"({listed}), {_count(ok, 'formula gives', 'formulas give')} the same."))
    elif ok:
        lines.append(T(f"Nachgerechnet {'ergibt die Formel' if ok == 1 else f'ergeben alle {ok} Formeln'}, die Gegenprobe "
                       "nachrechnen kann, dasselbe wie in Excel.",
                       f"Recalculated, {'the formula' if ok == 1 else f'all {ok} formulas'} that Gegenprobe can "
                       f"recalculate {'gives' if ok == 1 else 'give'} the same as in Excel."))
    if unknown:
        today = any(c.get("today") for d in described for c in d["columns"] + d["figures"])
        lines.append(T(f"{_count(unknown, 'Formel rechnet', 'Formeln rechnet')} Gegenprobe nicht nach"
                       f"{' (darunter solche mit HEUTE(), das jeden Tag etwas anderes ergibt)' if today else ''}.",
                       f"Gegenprobe does not recalculate {_count(unknown, 'formula', 'formulas')}"
                       f"{' (including some with TODAY(), which gives a different result every day)' if today else ''}."))
    links = []
    for name in order:
        for other, (_, functions, _) in infos[name].reads.items():
            how = f" ({_listing([f for f, _ in functions.most_common(2)])})" if functions else ""
            links.append(T(f"„{name}“ liest aus „{other}“{how}", f"“{name}” reads from “{other}”{how}"))
    if links:
        lines.append(_listing(links) + ".")
    elif len(order) > 1:
        lines.append(T("Die Blätter rechnen je für sich, keine Formel liest aus einem anderen Blatt.",
                       "The sheets calculate separately; no formula reads from another sheet."))
    counted = [d for d in described if "hints" in d]
    per = [T(f"„{d['name']}“ ({_levels(d['hints'])})", f"“{d['name']}” ({_levels(d['hints'])})")
           for d in counted if d["hints"]["fehler"] or d["hints"]["pruefen"]]
    if per:
        lines.append(T(f"Fehler und Stellen zum Prüfen gibt es auf {_listing(per)}.",
                       f"There are errors and places to check on {_listing(per)}."))
    elif counted and len(counted) == len(described):
        lines.append(T("Fehler oder Stellen zum Prüfen hat Gegenprobe auf keinem Blatt gefunden.",
                       "Gegenprobe found no errors or places to check on any sheet."))
    if len(results) == 1:
        lines.append(T("1 Ergebnis, das keine andere Formel weiterverwendet, steht unten mit dem, woraus es entsteht.",
                       "1 result that no other formula uses is listed below with what it comes from."))
    elif results:
        lines.append(T(f"{len(results)} Ergebnisse, die keine andere Formel weiterverwendet, stehen unten mit dem, woraus "
                       "sie entstehen.",
                       f"{len(results)} results that no other formula uses are listed below with what they come from."))
    if checks.skipped:
        lines.append(T(f"{_count(checks.skipped, 'Formel', 'Formeln')} habe ich nicht mehr nachgerechnet, weil die Mappe "
                       "so groß ist.",
                       f"I did not recalculate {_count(checks.skipped, 'formula', 'formulas')} because the workbook "
                       "is so large."))
    return lines


def _count(n: int, one: str, many: str) -> str:
    return f"{n} {one if n == 1 else many}"


def _ordinal(n: int) -> str:
    """3 → „3rd“, für „every 3rd“."""
    suffix = "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def _levels(hints: dict) -> str:
    parts = [T(f"{hints['fehler']} Fehler", _count(hints["fehler"], "error", "errors")) if hints["fehler"] else "",
             T(f"{hints['pruefen']} zum Prüfen", f"{hints['pruefen']} to check") if hints["pruefen"] else ""]
    return ", ".join(p for p in parts if p)


def _listing(items: list[str]) -> str:
    return items[0] if len(items) == 1 else f"{', '.join(items[:-1])}{T(' und ', ' and ')}{items[-1]}"
