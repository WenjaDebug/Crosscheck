"""Berichte mit Zwischensummen: „Summe Nord“ unter den Zeilen von Nord, „Summe Süd“ unter Süd und ganz unten „Gesamt“.

Zwei Fehler, die in solchen Berichten typisch sind und die die allgemeinen Prüfungen nicht sehen:

- Die Gesamtsumme zählt die Zwischensummen einzeln zusammen (=C5+C9+C13) und lässt eine aus: Summe West fehlt.
- Die Gesamtsumme summiert alles darüber (=SUMME(C2:C17)), die Zwischensummen mit: Jeder Wert zählt doppelt.
  TEILERGEBNIS(9; …) darf das, es lässt andere Teilergebnisse im Bereich weg (so baut Excel selbst solche Berichte,
  Daten → Teilergebnis).

Zwischensummen sind SUMMEN (oder TEILERGEBNIS(9; …)) über einen Bereich ihrer eigenen Spalte direkt darüber, deren
Bereiche sich nicht überschneiden, mindestens zwei in einer Spalte. Alles ohne KI.
"""

from __future__ import annotations

import math
import re

from . import formel
from .analyse import Finding, Sheet, _fmt, address
from .bereiche import _german_formula
from .formel import Node
from .sprache import T
from .summen import _is_total
from .xlsx import column_letters

Pos = tuple[int, int]

MAX_FINDINGS = 3
_SUBTOTALS = ("SUBTOTAL", "AGGREGATE")


def _own_sum(node: Node | None, pos: Pos) -> tuple[int, int] | None:
    """Die erste und letzte Zeile, wenn die Zelle SUMME oder TEILERGEBNIS(9; …) über ihre eigene Spalte darüber ist."""
    if node is None or node.kind != "call":
        return None
    args = node.args
    if node.value == "SUBTOTAL" and len(args) == 2 and args[0].kind == "num" and int(args[0].value) in (9, 109):
        args = args[1:]
    elif node.value != "SUM" or len(args) != 1:
        return None
    area = args[0]
    if area.kind != "range":
        return None
    r1, c1, r2, c2 = area.value
    if c1 != c2 or c1 != pos[1] or r2 >= pos[0]:
        return None
    return r1, r2


def _row_name(sheet: Sheet, row: int, col: int) -> str:
    """Wie die Zeile heißt: ihr Name in der Tabelle oder der erste Text links in der Zeile („Summe West“)."""
    name = sheet.plan.names.get(row)
    if name:
        return name.strip()
    for c in range(1, col):
        cell = sheet.cells.get((row, c))
        if cell is not None and cell.formula is None and cell.value.strip() and any(ch.isalpha() for ch in cell.value):
            return cell.value.strip()
    return ""


def _named(sheet: Sheet, pos: Pos) -> str:
    name = _row_name(sheet, *pos)
    return f"{name} ({address(pos)})" if name else address(pos)


def _listing(items: list[str]) -> str:
    return items[0] if len(items) == 1 else ", ".join(items[:-1]) + T(f" und {items[-1]}", f" and {items[-1]}")


def _addends(node: Node) -> list[Pos] | None:
    """Die Zellen einer reinen Summe einzelner Zellen (=C5+C9+C13 oder =SUMME(C5;C9;C13)), sonst None."""
    if node.kind == "ref":
        return [node.value]
    if node.kind == "op" and node.value == "+":
        left, right = _addends(node.args[0]), _addends(node.args[1])
        return None if left is None or right is None else left + right
    if node.kind == "call" and node.value == "SUM" and node.args and all(a.kind == "ref" for a in node.args):
        return [a.value for a in node.args]
    return None


def _groups(sheet: Sheet) -> dict[int, list[tuple[int, int, int]]]:
    """Je Spalte die Zwischensummen (Zeile, erste, letzte), wenn es mindestens zwei sind."""
    by_col: dict[int, list[tuple[int, int, int]]] = {}
    for pos, node in sheet.nodes.items():
        rows = _own_sum(node, pos)
        if rows is not None:
            by_col.setdefault(pos[1], []).append((pos[0], *rows))
    out = {}
    for col, sums in by_col.items():
        # Eine Summe, in deren Bereich eine andere steht, fasst zusammen: die Gesamtsumme, keine Zwischensumme.
        inner = [s for s in sums if not any(o is not s and s[1] <= o[0] <= s[2] for o in sums)]
        inner.sort()
        if len(inner) >= 2 and all(a[0] < b[1] for a, b in zip(inner, inner[1:])):
            out[col] = inner
    return out


def _value(sheet: Sheet, pos: Pos) -> float | None:
    value = sheet.inputs.get(pos)
    return value if isinstance(value, float) and not isinstance(value, bool) and math.isfinite(value) else None


def _fix(sheet: Sheet, calc, english: str, expected: float) -> str:
    """Die deutsche Formel, wenn sie nachgerechnet das Erwartete ergibt."""
    try:
        got = calc.run(formel.parse(english))
    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
        return ""
    if not isinstance(got, float) or abs(got - expected) > 1e-6 * max(1.0, abs(expected)):
        return ""
    return _german_formula(english)


def checks(sheet: Sheet) -> tuple[list[Finding], set[str]]:
    """Die Hinweise und die Schlüssel allgemeinerer Hinweise zu denselben Zellen, die sie ersetzen."""
    found: list[Finding] = []
    covered: set[str] = set()
    groups = _groups(sheet)
    if not groups:
        return found, covered
    calc, _ = sheet.evaluator()
    for col, subtotals in sorted(groups.items()):
        rows = {s[0] for s in subtotals}
        last = max(rows)
        for (r, c), node in sorted(sheet.nodes.items()):
            if c != col or r <= last or node is None or len(found) >= MAX_FINDINGS:
                continue
            if not _is_total(_row_name(sheet, r, c)):
                continue
            pos = (r, c)
            got = _value(sheet, pos)
            values = [_value(sheet, (s, col)) for s in sorted(rows)]
            if got is None or any(v is None for v in values):
                continue
            expected = math.fsum(values)
            finding = _missing(sheet, calc, pos, node, sorted(rows), got, expected)
            finding = finding or _doubled(sheet, calc, pos, node, subtotals, got, expected)
            if finding is not None:
                found.append(finding)
                # „B5 und B9 gehören nicht dazu“ (label:) sagt dasselbe ohne die Ursache: doppelt gezählt.
                covered |= {f"double:{address(pos)}", f"parts:{address(pos)}", f"label:{address(pos)}"}
            break
    return found, covered


def _missing(sheet: Sheet, calc, pos: Pos, node: Node, rows: list[int], got: float, expected: float) -> Finding | None:
    """„Gesamt“ =C5+C9+C13 zählt Summe West (C17) nicht mit."""
    cells = _addends(node)
    col = pos[1]
    if cells is None or any(p[1] != col for p in cells):
        return None
    read = {p[0] for p in cells}
    if not read <= set(rows) or len(read) < 2 or len(read) == len(rows):
        return None
    left_out = [(r, col) for r in rows if r not in read]
    if abs(got - expected) <= 1e-6 * max(1.0, abs(expected)):
        return None
    letters = column_letters(col)
    formula = sheet.cells[pos].formula or ""
    if re.match(r"^=?\s*SUM\(", formula, re.I):
        english = "SUM(" + ",".join(f"{letters}{r}" for r in rows) + ")"
    else:
        english = "+".join(f"{letters}{r}" for r in rows)
    fixed = _fix(sheet, calc, english, expected)
    if not fixed:
        return None
    total = _named(sheet, pos)
    text = T(f"{total} zählt nur {len(read)} der {len(rows)} Zwischensummen zusammen: "
             f"{_listing([_named(sheet, p) for p in left_out])} {'fehlt' if len(left_out) == 1 else 'fehlen'}. "
             f"So ergibt {address(pos)} {_fmt(got)} statt {_fmt(expected)}.",
             f"{total} adds up only {len(read)} of the {len(rows)} subtotals: "
             f"{_listing([_named(sheet, p) for p in left_out])} {'is' if len(left_out) == 1 else 'are'} missing. "
             f"So {address(pos)} gives {_fmt(got)} instead of {_fmt(expected)}.")
    return Finding(f"subtotal:{address(pos)}", text, "fehler",
                   T(f"Korrigiert: ={fixed}, nachgerechnet {_fmt(expected)}.",
                     f"Fixed: ={fixed}, recalculated this gives {_fmt(expected)}."))


def _doubled(sheet: Sheet, calc, pos: Pos, node: Node, subtotals: list[tuple[int, int, int]], got: float,
             expected: float) -> Finding | None:
    """„Gesamt“ =SUMME(C2:C17) zählt die Zeilen und ihre Zwischensummen: alles doppelt. TEILERGEBNIS darf das."""
    rows = _own_sum(node, pos)
    if rows is None or node.value in _SUBTOTALS and all(sheet.nested((s[0], pos[1])) for s in subtotals):
        return None
    first, last = rows
    inside = [s for s in subtotals if first <= s[0] <= last]
    if len(inside) != len(subtotals) or abs(got - expected) <= 1e-6 * max(1.0, abs(expected)):
        return None
    # Jede Zahl im Bereich gehört zu einer Gruppe: Sonst fehlt der Summe der Zwischensummen etwas.
    grouped = {r for s in subtotals for r in range(s[1], s[0] + 1)}
    if any(_value(sheet, (r, pos[1])) is not None and r not in grouped for r in range(first, last + 1)):
        return None
    letters = column_letters(pos[1])
    english = "SUM(" + ",".join(f"{letters}{s[0]}" for s in subtotals) + ")"
    fixed = _fix(sheet, calc, english, expected)
    if not fixed:
        return None
    text = T(f"{_named(sheet, pos)} summiert {letters}{first}:{letters}{last} und zählt damit die Zeilen und ihre "
             f"{len(inside)} Zwischensummen ({_listing([_named(sheet, (s[0], pos[1])) for s in inside])}): Jeder Wert "
             f"zählt doppelt. So ergibt {address(pos)} {_fmt(got)} statt {_fmt(expected)}.",
             f"{_named(sheet, pos)} sums {letters}{first}:{letters}{last} and so counts the rows and their "
             f"{len(inside)} subtotals ({_listing([_named(sheet, (s[0], pos[1])) for s in inside])}): every value "
             f"counts twice. So {address(pos)} gives {_fmt(got)} instead of {_fmt(expected)}.")
    tip = T(f"Korrigiert: ={fixed}, nachgerechnet {_fmt(expected)}. Oder nimm überall TEILERGEBNIS(9; …): Das lässt "
            "die Zwischensummen im Bereich von selbst weg.",
            f"Fixed: ={fixed}, recalculated this gives {_fmt(expected)}. Or use SUBTOTAL(9, …) everywhere: it "
            "automatically leaves out the subtotals in the range.")
    return Finding(f"subtotal:{address(pos)}", text, "fehler", tip)
