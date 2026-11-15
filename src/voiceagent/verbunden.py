"""Verbundene Zellen in einer Liste (Start → Verbinden und zentrieren): Excel zeigt den Wert über alle Zeilen des
Verbunds, er steht aber nur in der ersten Zelle. ZÄHLENWENN, SUMMEWENN, SVERWEIS, Filter und Pivot-Tabellen sehen in
den übrigen Zeilen eine leere Zelle, und das sieht man der Liste nicht an.

Ein Hinweis nur, wo das etwas ändert: Formeln, die mit dem Wert in jeder Zeile anders rechnen, oder ein Filter über
der Liste. Verbundene Überschriften, Titel und Zellen außerhalb einer Liste bleiben, wie sie sind, ebenso Finanzpläne,
in denen jede Zeile eine eigene Größe ist. Einsetzen hebt den Verbund auf und schreibt den Wert in jede Zeile;
Rückgängig verbindet die Zellen wieder."""

from __future__ import annotations

import re

from .analyse import Finding, Refs, Sheet, address, changed_by, more_formulas, said
from .sprache import T
from .workbook import Layout
from .xlsx import Area, AutoFilter, Cell, area_text, in_area

Pos = tuple[int, int]
MAX_LISTED = 3
# Formeln, die je Zeile nach einem Wert suchen oder zählen: Mit ihnen ist eine leere Zelle unter dem Verbund sicher
# falsch. ANZAHL2 über der Spalte kann dagegen gewollt die Gruppen zählen, das ist nur zu prüfen.
_SEARCHING = re.compile(r"\b(?:COUNTIFS?|SUMIFS?|AVERAGEIFS?|MAXIFS|MINIFS|VLOOKUP|HLOOKUP|XLOOKUP|LOOKUP|X?MATCH|"
                        r"FILTER|SUMPRODUCT|DCOUNTA?|DSUM|DAVERAGE|UNIQUE|GROUPBY|PIVOTBY)\(", re.I)


def checks(cells: dict[Pos, Cell], plans: list[Layout], merged: list[Area], name: str | None = None,
           book: dict | None = None, filters: list[AutoFilter] | None = None) -> list[Finding]:
    """plans: die Tabellen des Blatts (die erste ist die Haupttabelle). merged: die verbundenen Bereiche."""
    groups: dict[tuple[int, int], list[Area]] = {}
    for area in merged:
        top, left, bottom, right = area
        if bottom == top:
            continue  # nur nebeneinander: eine Überschrift über mehreren Spalten
        index = next((i for i, plan in enumerate(plans)
                      if not plan.across and plan.in_table(top, left) and plan.in_table(bottom, left)), None)
        anchor = cells.get((top, left))
        if index is None or anchor is None or anchor.formula is not None or not anchor.value.strip():
            continue
        # Steht in einer der übrigen Zellen etwas (aus einer Datei, die das zulässt), ist es kein reiner Verbund.
        if any((r, c) in cells for r in range(top, bottom + 1) for c in range(left, right + 1) if (r, c) != (top, left)):
            continue
        groups.setdefault((index, left), []).append(area)
    if not groups:
        return []
    sheet = Sheet(cells, plans[0], name, book)
    found = []
    for (index, col), areas in sorted(groups.items(), key=lambda g: (g[0][1], g[0][0])):
        item = _column(sheet, plans, index, col, sorted(areas), filters or [])
        if item is not None:
            found.append(item)
    return found


def _column(sheet: Sheet, plans: list[Layout], index: int, col: int, areas: list[Area],
            filters: list[AutoFilter]) -> Finding | None:
    plan = plans[index]
    fill: dict[Pos, object] = {}
    for top, left, bottom, _ in areas:
        value = sheet.inputs.get((top, left))
        for r in range(top + 1, bottom + 1):
            fill[(r, left)] = value
    direct = _readers(sheet, set(fill), [(t + 1, left, b, left) for t, left, b, _ in areas])
    changed = changed_by(sheet, direct, fill)
    filtered = any(in_area(f.area, top, col) for f in filters for top, *_ in areas)
    if not changed and not filtered:
        return None
    head = plan.column(col)
    first = areas[0]
    if len(areas) == 1:
        top, _, bottom, _ = first
        value = sheet.shown.get((top, col), "")
        lead = T(f"{area_text(first)} ({head}) ist verbunden: Excel zeigt „{value}“ in allen {bottom - top + 1} Zeilen, es "
                 f"steht aber nur in {address((top, col))}. Für Formeln {'ist' if bottom == top + 1 else 'sind'} "
                 f"{_cells_said(top + 1, bottom, col)} leer.",
                 f"{area_text(first)} ({head}) is merged: Excel shows “{value}” in all {bottom - top + 1} rows, but "
                 f"it is only stored in {address((top, col))}. For formulas, {_cells_said(top + 1, bottom, col)} "
                 f"{'is' if bottom == top + 1 else 'are'} empty.")
    else:
        listed = ", ".join(area_text(a) for a in areas[:MAX_LISTED])
        rest = len(areas) - MAX_LISTED
        more = T(" und ein weiterer", " and one more") if rest == 1 else T(f" und {rest} weitere", f" and {rest} more") if rest > 1 else ""
        lead = T(f"In Spalte {head} sind {len(areas)} Bereiche verbunden ({listed}{more}): Excel zeigt den Wert in jeder "
                 f"Zeile, er steht aber nur in der ersten Zelle jedes Bereichs. Für Formeln sind die übrigen {len(fill)} "
                 "Zellen leer.",
                 f"In column {head}, {len(areas)} ranges are merged ({listed}{more}): Excel shows the value in every "
                 f"row, but it is only stored in the first cell of each range. For formulas, the other {len(fill)} "
                 "cells are empty.")
    effect = ""
    if changed:
        shown = [T(f"{_label(sheet, plans, p)} {_value(b)} statt {_value(a)}", f"{_label(sheet, plans, p)} gives {_value(b)} instead of {_value(a)}")
                 for p, b, a in changed[:MAX_LISTED]]
        tail = more_formulas(len(changed) - MAX_LISTED) if len(changed) > MAX_LISTED else ""
        effect = T(f" Darum ergibt {', '.join(shown) + tail if tail else _listed(shown)}.",
                   f" So {'; '.join(shown) + tail if tail else _listed(shown, '; ')}.")
    if filtered:
        effect += T(" Filterst du die Liste nach dieser Spalte, fallen diese Zeilen weg.",
                    " If you filter the list by this column, these rows drop out.")
    text = T(f"{lead}{effect} Auch Pivot-Tabellen sehen dort nichts. Einsetzen hebt den Verbund auf und schreibt den Wert "
             "in jede Zeile.",
             f"{lead}{effect} Pivot tables see nothing there either. Applying the fix unmerges the cells and writes the value "
             "into every row.")
    searching = any(_SEARCHING.search(sheet.cells[p].formula or "") for p, *_ in changed if p in direct)
    level = "fehler" if searching else "pruefen"
    cell = f"{address((first[0], col))}:{address((areas[-1][2], col))}"
    return Finding(f"merged:{address((first[0], col))}", text, level, fix=_fix(sheet, fill, areas), cell=cell)


def _label(sheet: Sheet, plans: list[Layout], pos: Pos) -> str:
    """„G2 (Fälle Kardiologie)“: mit den Namen der Tabelle, in der die Formel steht, auch der Auswertung daneben."""
    plan = next((p for p in plans if p.in_table(*pos) and not p.across), sheet.plan)
    if plan is sheet.plan:
        return sheet.label(pos)
    main, sheet.plan = sheet.plan, plan
    try:
        return sheet.label(pos)
    finally:
        sheet.plan = main


def _cells_said(top: int, bottom: int, col: int) -> str:
    if top == bottom:
        return address((top, col))
    return T(f"{address((top, col))} bis {address((bottom, col))}", f"{address((top, col))} to {address((bottom, col))}")


def _listed(parts: list[str], english: str = ", ") -> str:
    """„A, B und C“; english: womit die englische Fassung trennt („; “, wenn die Teile mit einer Zahl enden)."""
    return parts[0] if len(parts) == 1 else T(", ", english).join(parts[:-1]) + T(f" und {parts[-1]}", f" and {parts[-1]}")


def _overlap(a: Area, b: Area) -> bool:
    return a[0] <= b[2] and b[0] <= a[2] and a[1] <= b[3] and b[1] <= a[3]


def _readers(sheet: Sheet, spots: set[Pos], boxes: list[Area]) -> set[Pos]:
    """Die Formeln, die eine der leeren Zellen unter einem Verbund lesen, einzeln oder über einen Bereich."""
    found = set()
    for pos, areas in sheet.ranges.items():
        if pos in sheet.unknown:
            continue
        before = sheet.precedents.get(pos)
        singles = before.singles if isinstance(before, Refs) else before or ()
        if any(_overlap(a, b) for a in areas for b in boxes) or any(p in spots for p in singles):
            found.add(pos)
    return found


def _value(value) -> str:
    return T(f"„{value}“", f"“{value}”") if isinstance(value, str) else said(value)


def _fix(sheet: Sheet, fill: dict[Pos, object], areas: list[Area]) -> dict | None:
    """Einsetzen: den Verbund aufheben (unmerge) und den Wert der ersten Zelle in jede Zeile schreiben, eine Zahl oder
    ein Datum im Format der ersten Zelle."""
    from . import einsetzen

    if not fill or len(fill) > einsetzen.MAX_CELLS or any(v is None for v in fill.values()):
        return None
    targets = sorted(fill, key=lambda p: (p[1], p[0]))
    # Zum Anzeigen: der ganze verbundene Teil der Spalte („A2 bis A21“), geschrieben wird nur in die leeren Zellen.
    found = {"range": f"{address(areas[0][:2])}:{address((areas[-1][2], areas[0][1]))}", "cells": {address(p): fill[p] for p in targets},
             "was": {address(p): {"text": ""} for p in targets}, "unmerge": [area_text(a) for a in areas]}
    if len(areas) == 1:
        found["german"] = T(f"„{sheet.shown.get(areas[0][:2], '')}“ in jede Zeile", f"“{sheet.shown.get(areas[0][:2], '')}” in every row")
    else:
        found["german"] = T("den Wert jedes Verbunds in jede Zeile", "the value of each merged range in every row")
    numbers = {address(p): address(_anchor(p, areas)) for p in targets if isinstance(fill[p], float)}
    if numbers:
        found["format"] = numbers
    if any(isinstance(v, str) for v in fill.values()):
        found["text"] = True
    return found


def _anchor(pos: Pos, areas: list[Area]) -> Pos:
    return next(a[:2] for a in areas if in_area(a, *pos))
