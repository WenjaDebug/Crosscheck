"""Beträge in Tausend und in Einheiten, die in einer Formel zusammengezählt werden.

„Umsatz (T€)“ minus „Kosten (€)“ geht nur, wenn eine der beiden Spalten umgerechnet wird. Sonst stimmt entweder der
Spaltenkopf nicht (die Zahlen sind gleich groß, beide in Euro) oder die Formel (die Zahlen sind wirklich in Tausend,
dann fehlt *1000). Welches von beiden, sagen die Zahlen selbst: wie groß sie gegeneinander sind.

Ohne KI: Der Prüfer liest die Einheit aus den Spaltenköpfen, die Formeln, die Spalten mit + und − verbinden, und
vergleicht die Größenordnung der Zahlen.
"""

from __future__ import annotations

import math
import re
import statistics

from . import formel
from .analyse import Finding, Sheet, _fmt, address
from .sprache import T
from .workbook import _left_label, pattern
from .xlsx import column_letters

Pos = tuple[int, int]

_CURRENCY = {"€": "EUR", "eur": "EUR", "euro": "EUR", "chf": "CHF", "fr.": "CHF", "sfr": "CHF", "usd": "USD", "$": "USD",
             "gbp": "GBP", "£": "GBP"}
_FACTOR = {"t": 1e3, "tsd": 1e3, "tsd.": 1e3, "k": 1e3, "tausend": 1e3, "mio": 1e6, "mio.": 1e6, "mn": 1e6, "m": 1e6,
           "millionen": 1e6, "mrd": 1e9, "mrd.": 1e9, "milliarden": 1e9}
_CUR = r"€|eur(?:o)?|chf|sfr|fr\.|usd|\$|gbp|£"
_PRE = r"t|tsd\.?|k|mio\.?|mn|m|mrd\.?|tausend|millionen|milliarden"
# „T€“, „TEUR“, „Tsd. €“, „kCHF“, „Mio. CHF“, „€ Tsd.“, „in Tausend €“, „€“: was davor oder dahinter steht, zählt.
_UNIT = re.compile(rf"(?<![^\W\d_])(?:(?P<pre>{_PRE})\s*(?P<cur1>{_CUR})|(?P<cur2>{_CUR})\s*(?P<post>{_PRE})(?![^\W\d_])"
                   rf"|(?P<cur3>{_CUR}))(?![^\W\d_])", re.IGNORECASE)
_ADDITIVE = {"SUM", "AVERAGE", "MIN", "MAX", "MEDIAN"}
MAX_EACH = 2


def _unit(header: str) -> tuple[str, float, re.Match] | None:
    """„Umsatz (T€)“ → ("EUR", 1000, Fundstelle), „Kosten (€)“ → ("EUR", 1, …). Nur eine Einheit im Namen."""
    found = list(_UNIT.finditer(header))
    if len(found) != 1:
        return None
    m = found[0]
    if m.group("cur3"):
        # „M“ oder „T“ allein ist keine Einheit, „€“ allein schon.
        return _CURRENCY[m.group("cur3").casefold()], 1.0, m
    cur = (m.group("cur1") or m.group("cur2")).casefold()
    factor = _FACTOR[(m.group("pre") or m.group("post")).casefold()]
    return _CURRENCY[cur], factor, m


def _terms(node, sheet_name: str) -> list[int] | None:
    """Die Spalten, die eine Formel nur zusammenzählt oder abzieht (B5−C5, G4+B5, B5−MITTELWERT($B$5:$B$16)).
    None, sobald etwas mal oder durch etwas rechnet: dann ändert sich die Einheit."""
    if node.kind == "ref":
        return [node.value[1]]
    if node.kind == "num":
        return []
    if node.kind == "neg":
        return _terms(node.args[0], sheet_name)
    if node.kind == "op" and node.value in ("+", "-") and len(node.args) == 2:
        left, right = _terms(node.args[0], sheet_name), _terms(node.args[1], sheet_name)
        return left + right if left is not None and right is not None else None
    if node.kind == "call" and node.value in _ADDITIVE and node.args:
        cols = []
        for arg in node.args:
            if arg.kind == "range" and arg.value[1] == arg.value[3]:
                cols.append(arg.value[1])
            elif arg.kind == "ref":
                cols.append(arg.value[1])
            else:
                return None
        return cols
    return None


def checks(sheet: Sheet, rows: list[int]) -> list[Finding]:
    plan = sheet.plan
    if plan.header_row is None or len(rows) < 3:
        return []
    heads = _heads(sheet)
    data = set(rows)
    units = {c: u for c, h in heads.items() if (u := _unit(h))}
    if len(units) < 2:
        return []
    # Welche Spalten in Formeln zusammengezählt werden: je Formelspalte einmal (heruntergezogen ist es dieselbe).
    links: dict[int, set[int]] = {}
    uses: dict[int, list[Pos]] = {}
    shapes: set[tuple[int, str]] = set()
    for pos, node in sorted(sheet.nodes.items(), key=lambda item: (item[0][1], item[0][0])):
        if node is None or pos in sheet.unknown or pos[0] not in data or pos[1] not in heads:
            continue
        cols = _terms(node, sheet.name)
        if not cols:
            continue
        group = set(cols) | {pos[1]}
        shape = (pos[1], pattern(sheet.cells[pos].formula, pos))
        if len(group) < 2:
            continue
        for c in group:
            links.setdefault(c, set()).update(group - {c})
        if shape not in shapes:
            shapes.add(shape)
            for c in set(cols) - {pos[1]}:
                uses.setdefault(c, []).append(pos)
    _, now = sheet.evaluator()
    found = []
    for col, (cur, factor, match) in sorted(units.items()):
        partners = {p for p in links.get(col, ()) if p in units and units[p][0] == cur}
        other = {p for p in partners if units[p][1] != factor}
        # Allein gegen alle: mindestens zwei andere Spalten in einer anderen Größe, keine in derselben.
        if len(other) < 2 or len(other) != len(partners) or len({units[p][1] for p in other}) != 1:
            continue
        theirs = units[next(iter(other))]
        mine = _size(sheet, now, col, rows)
        sizes = [s for p in other if (s := _size(sheet, now, p, rows))]
        if mine is None or not sizes:
            continue
        ratio = math.log10(mine / statistics.median(sizes))
        expected = math.log10(theirs[1] / factor)
        header = heads[col]
        partner_names = [T(f"{column_letters(p)} „{heads[p]}“", f"{column_letters(p)} “{heads[p]}”") for p in sorted(other)]
        listed = ", ".join(partner_names[:-1]) + T(" und ", " and ") + partner_names[-1]
        said = match.group(0)
        if abs(ratio) <= 1 and abs(ratio - expected) >= 2:
            # Gleich groß wie die Spalten in Euro: Der Spaltenkopf stimmt nicht.
            spelled = _spelled(heads, other, theirs)
            new = header[:match.start()] + spelled + header[match.end():]
            head = (plan.header_row, col)
            cell = sheet.cells.get(head)
            fix = None
            if cell is not None and not cell.formula and (cell.value or "").strip() == header.strip():
                fix = {"range": address(head), "cells": {address(head): new}, "was": {address(head): {"text": cell.value}},
                       "german": new, "text": True}
            example = uses[col][0] if uses.get(col) else None
            typical, others, meant = (_fmt(round(mine)), _fmt(round(statistics.median(sizes))),
                                      _fmt(round(mine * factor / theirs[1])))
            text = T(f"{address(head)} heißt „{header}“, die Zahlen darunter sind aber nicht in {_name(said)}: Sie sind so groß "
                     f"wie die in {listed}, mit denen die Formeln sie ohne Umrechnung zusammenzählen oder abziehen"
                     + (f" (etwa {address(example)} ={_german(sheet, example)})" if example else "")
                     + f". Typisch steht in {column_letters(col)} {typical}, in den anderen "
                     f"{others}. Wären es {_name(said)}, stünde {typical} für "
                     f"{meant} {_name(spelled)}.",
                     f"{address(head)} is called “{header}”, but the numbers below it are not in {_name(said)}: they are "
                     f"the same size as those in {listed}, and the formulas add or subtract them without converting"
                     + (f" (e.g. {address(example)} ={_german(sheet, example)})" if example else "")
                     + f". Typically {column_letters(col)} has {typical}, the others {others}. If they were "
                     f"{_name(said)}, {typical} would stand for {meant} {_name(spelled)}.")
            tip = T(f"Korrigiert in {address(head)}: {new}.", f"Corrected in {address(head)}: {new}.") if fix else \
                T(f"Schreib in {address(head)} {spelled} statt {said}.", f"In {address(head)}, write {spelled} instead of {said}.")
            found.append(Finding(f"unit:{address(head)}", text, "fehler", tip, fix=fix, cell=address(head)))
        elif abs(ratio - expected) <= 1 and abs(ratio) >= 2:
            # Wirklich in Tausend: Die Formeln zählen Tausender und Einer zusammen.
            first = uses[col][0] if uses.get(col) else None
            if first is None:
                continue
            scale = factor / theirs[1]
            changes = _converted(sheet, col, scale, uses[col], data)
            fix = None
            if changes:
                from . import einsetzen

                fix = einsetzen.formulas(sheet, changes)
            letter, typical, others = column_letters(col), _fmt(round(mine)), _fmt(round(statistics.median(sizes)))
            text = T(f"{address(first)} zählt {letter} „{header}“ mit {listed} zusammen, ohne umzurechnen: "
                     f"{letter} ist in {_name(said)}, die anderen nicht. Typisch {typical} gegen "
                     f"{others}.",
                     f"{address(first)} adds {letter} “{header}” together with {listed} without converting: {letter} is "
                     f"in {_name(said)}, the others are not. Typically {typical} versus {others}.")
            if fix:
                cells = sorted(changes, key=lambda p: (p[1], p[0]))
                example = f"{address(cells[0])} ={_german_text(changes[cells[0]])}"
                text += T(f" Einsetzen rechnet {letter} in den {len(cells)} Formeln mal {_fmt(scale)} um, "
                          f"etwa {example}.",
                          f" Applying the fix multiplies {letter} by {_fmt(scale)} in the {len(cells)} "
                          f"formula{'' if len(cells) == 1 else 's'}, e.g. {example}.")
                tip = ""
            else:
                tip = T(f"Rechne {letter} vorher um, mal {_fmt(scale)}, oder die anderen Spalten "
                        f"in {_name(said)}.",
                        f"Convert {letter} first (multiply by {_fmt(scale)}) or convert the other columns to "
                        f"{_name(said)}.")
            found.append(Finding(f"unit:{address(first)}", text, "fehler" if fix else "pruefen", tip, fix=fix))
        if len(found) >= MAX_EACH:
            break
    flagged = {f.key.partition(":")[2] for f in found} | {address((plan.header_row, c)) for c in units
                                                        if any(f.key == f"unit:{address((plan.header_row, c))}" for f in found)}
    found += _currencies(sheet, units, uses, data, flagged)
    found += _title(sheet, units, flagged)
    return found


_RATE = re.compile(r"kurs|wechselkurs|umrechnung|fx|exchange|rate|devisen", re.IGNORECASE)


def _currencies(sheet: Sheet, units: dict, uses: dict, data: set[int], flagged: set[str]) -> list[Finding]:
    """CHF und EUR in einer Formel nur zusammengezählt: ohne Kurs. Steht der Kurs in einer Zelle, deren Name beide
    Währungen nennt, rechnet Einsetzen damit um."""
    heads = _heads(sheet)
    found = []
    seen = set()
    for pos, node in sorted(sheet.nodes.items(), key=lambda item: (item[0][1], item[0][0])):
        if node is None or pos in sheet.unknown or pos[0] not in data or pos[1] not in units or pos[1] in seen:
            continue
        cols = _terms(node, sheet.name)
        if not cols:
            continue
        mine = units[pos[1]][0]
        foreign = sorted({c for c in cols if c in units and units[c][0] != mine})
        if not foreign or address(pos) in flagged:
            continue
        seen.add(pos[1])
        other = units[foreign[0]][0]
        rate = None
        for where, cell in sheet.cells.items():
            label = _left_label(sheet.cells, *where) or ""
            v = sheet.inputs.get(where)
            if cell.formula is None and isinstance(v, float) and v > 0 and _RATE.search(label) and \
                    {mine, other} <= {_CURRENCY.get(w.casefold(), "") for w in re.findall(r"[A-Za-z€$£]+|€", label)}:
                rate = where
                break
        changes = None
        if rate is not None and len(foreign) == 1:
            fixed = f"${column_letters(rate[1])}${rate[0]}"
            changes = {}
            row = pos
            while row in sheet.nodes and row[0] in data and sheet.cells[row].formula and \
                    pattern(sheet.cells[row].formula, row) == pattern(sheet.cells[pos].formula, pos):
                ref = re.compile(rf"(?<![A-Za-z0-9_.$!:])(\$?{column_letters(foreign[0])}\$?{row[0]})(?![0-9A-Za-z(!:])")
                text = sheet.cells[row].formula or ""
                if len(ref.findall(text)) != 1:
                    changes = None
                    break
                changes[row] = ref.sub(lambda m: f"{m.group(1)}*{fixed}", text)
                row = (row[0] + 1, row[1])
        fix = None
        if changes:
            from . import einsetzen

            fix = einsetzen.formulas(sheet, changes)
        names = T(" und ".join(f"{column_letters(c)} „{heads[c]}“" for c in foreign),
                  " and ".join(f"{column_letters(c)} “{heads[c]}”" for c in foreign))
        text = T(f"{address(pos)} ({heads[pos[1]]}) zählt {names} in {other} mit Beträgen in {mine} zusammen, "
                 "ohne Wechselkurs.",
                 f"{address(pos)} ({heads[pos[1]]}) adds {names} in {other} to amounts in {mine} without an exchange "
                 "rate.")
        if fix:
            text += T(f" Der Kurs steht in {address(rate)} ({_left_label(sheet.cells, *rate)}): Einsetzen rechnet damit um.",
                      f" The rate is in {address(rate)} ({_left_label(sheet.cells, *rate)}): applying the fix converts "
                      f"with it.")
            tip = ""
        else:
            tip = T(f"Rechne {other} zuerst mit dem Kurs in {mine} um.",
                    f"First convert {other} to {mine} with the exchange rate.")
        found.append(Finding(f"currency:{address(pos)}", text, "fehler" if fix else "pruefen", tip, fix=fix))
        if len(found) >= MAX_EACH:
            break
    return found


def _title(sheet: Sheet, units: dict, flagged: set[str]) -> list[Finding]:
    """Ein Satz über der Tabelle („Alle Beträge in Euro“, „in Tausend CHF“), der einem Spaltenkopf widerspricht."""
    plan = sheet.plan
    heads = _heads(sheet)
    said = None
    for (r, c), cell in sorted(sheet.cells.items()):
        if r >= plan.header_row or cell.formula or not isinstance(cell.value, str):
            continue
        if re.search(r"beträge|betraege|zahlen|werte|angaben|alle|amounts|figures|in\s", cell.value, re.IGNORECASE):
            text = re.sub(r"(?i)\bin\s+tausend\b", "Tsd.", cell.value)
            unit = _unit(text)
            if unit:
                said = ((r, c), unit, cell.value)
                break
    if said is None:
        return []
    where, (cur, factor, _), sentence = said
    found = []
    for col, (c2, f2, m) in sorted(units.items()):
        head = address((plan.header_row, col))
        if c2 == cur and f2 != factor and head not in flagged:
            found.append(Finding(f"title:{head}",
                                 T(f"{address(where)} sagt „{sentence.strip()}“, der Kopf {head} aber "
                                   f"„{heads[col]}“. Eines von beiden stimmt nicht.",
                                   f"{address(where)} says “{sentence.strip()}”, but the heading {head} says "
                                   f"“{heads[col]}”. One of the two is wrong."), "pruefen",
                                 T("Prüf, in welcher Einheit die Zahlen wirklich sind, und pass das andere an.",
                                   "Check which unit the numbers are really in, and adjust the other one."),
                                 cell=head))
    return found[:MAX_EACH]


def _converted(sheet: Sheet, col: int, scale: float, firsts: list[Pos], data: set[int]) -> dict[Pos, str] | None:
    """Die Formeln, die col zusammenzählen, mit col umgerechnet: B5-C5 → B5*1000-C5. Rechnet eine Formel nur mit
    col (=B5, =B5-MITTELWERT($B$5:$B$16)), wird das Ganze umgerechnet. None, wenn es für eine nicht sicher geht."""
    factor = _fmt(scale).replace(",", ".")
    out: dict[Pos, str] = {}
    for first in firsts:
        shape = pattern(sheet.cells[first].formula, first)
        run = [first]
        while (nxt := (run[-1][0] + 1, first[1])) in sheet.nodes and nxt[0] in data and sheet.cells[nxt].formula and \
                pattern(sheet.cells[nxt].formula, nxt) == shape:
            run.append(nxt)
        for pos in run:
            text = sheet.cells[pos].formula or ""
            cols = _terms(sheet.nodes[pos], sheet.name) or []
            if cols and all(c == col for c in cols):
                out[pos] = f"({text})*{factor}"
                continue
            if any(a.kind == "range" and a.value[1] == col for c in formel.calls(sheet.nodes[pos]) for a in c.args):
                return None
            ref = re.compile(rf"(?<![A-Za-z0-9_.$!:])(\$?{column_letters(col)}\$?\d+)(?![0-9A-Za-z(!:])")
            if not ref.search(text):
                return None
            out[pos] = ref.sub(lambda m: f"{m.group(1)}*{factor}", text)
    return out or None


def _german_text(text: str) -> str:
    from .bereiche import _german_formula

    return _german_formula(text)


def _heads(sheet: Sheet) -> dict[int, str]:
    """Die Spaltenköpfe der Tabelle und dazu Köpfe in derselben Zeile über Spalten daneben („Kumuliert (€)“ in G4)."""
    plan = sheet.plan
    heads = dict(plan.headers)
    if plan.header_row is None:
        return heads
    for (r, c), cell in sheet.cells.items():
        if r == plan.header_row and c not in heads and not cell.formula and isinstance(cell.value, str) and \
                re.search(r"[^\W\d_]", cell.value):
            heads[c] = cell.value.strip()
    return heads


def _size(sheet: Sheet, now, col: int, rows: list[int]) -> float | None:
    values = []
    for r in rows:
        try:
            v = now((r, col))
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
            continue
        if isinstance(v, float) and not isinstance(v, bool) and v != 0:
            values.append(abs(v))
    return statistics.median(values) if len(values) >= 3 else None


def _spelled(headers: dict[int, str], cols: set[int], unit: tuple) -> str:
    """Die Einheit so geschrieben wie in den anderen Spalten („€“, nicht „EUR“)."""
    for c in sorted(cols):
        u = _unit(headers[c])
        if u and u[:2] == unit[:2]:
            return u[2].group(0)
    return unit[0]


def _name(said: str) -> str:
    thousand, thousand_chf = T("Tausend Euro", "thousand euros"), T("Tausend Franken", "thousand francs")
    words = {"t€": thousand, "teur": thousand, "k€": thousand, "keur": thousand,
             "tchf": thousand_chf, "kchf": thousand_chf, "€": T("Euro", "euros"), "eur": T("Euro", "euros"),
             "chf": T("Franken", "francs")}
    return words.get(re.sub(r"[\s.]", "", said.casefold()), said)


def _german(sheet: Sheet, pos: Pos) -> str:
    from .bereiche import _german_formula

    return _german_formula(sheet.cells[pos].formula or "")


# --- Zeilen mit verschiedener Einheit in einer Summe ---


def _row_scale(sheet: Sheet, pos: Pos) -> tuple[str, float] | None:
    """Die Einheit einer Zeile aus ihrer Einheitenspalte links („Mio. EUR“, „Tsd. EUR“), sonst None."""
    for c in range(pos[1] - 1, 0, -1):
        cell = sheet.cells.get((pos[0], c))
        if cell is None or cell.formula or not isinstance(cell.value, str) or not cell.value.strip() \
                or isinstance(sheet.inputs.get((pos[0], c)), (int, float)):
            continue
        text = cell.value.strip()
        if len(text) > 24:
            return None
        unit = _unit(text)
        return (unit[0], unit[1]) if unit else None
    return None


def _cells_summed(node) -> list[Pos] | None:
    """Die Zellen einer Spalte, die eine Formel nur zusammenzählt (SUMME(C6:C8), C6+C7+C8)."""
    if node.kind == "ref":
        return [node.value]
    if node.kind == "op" and node.value == "+" and len(node.args) == 2:
        left, right = _cells_summed(node.args[0]), _cells_summed(node.args[1])
        return left + right if left is not None and right is not None else None
    if node.kind == "call" and node.value == "SUM" and node.args:
        cells = []
        for arg in node.args:
            if arg.kind == "range" and arg.value[1] == arg.value[3] and arg.value[2] - arg.value[0] < 200:
                cells += [(r, arg.value[1]) for r in range(arg.value[0], arg.value[2] + 1)]
            elif arg.kind == "ref":
                cells.append(arg.value)
            else:
                return None
        return cells
    return None


def mixed_rows(sheet: Sheet) -> list[Finding]:
    """„Umsatz gesamt (Mio. EUR)“ = SUMME(C6:C8), aber C7 steht laut Einheitenspalte in Tsd. EUR: 38.200 statt 38,2
    zählen mit. Gemeldet, wenn die umgerechnete Zahl so groß ist wie die übrigen (dann ist es wirklich die Einheit).
    Ein Hinweis je Summenzeile, mit der Korrektur für alle ihre Spalten."""
    from . import einsetzen

    groups: dict[tuple, list] = {}
    for pos, node in sorted(sheet.nodes.items(), key=lambda item: (item[0][0], item[0][1])):
        if node is None or pos in sheet.unknown:
            continue
        cells = _cells_summed(node)
        if not cells or len(cells) < 2 or any(c[1] != pos[1] for c in cells):
            continue
        mine = _row_scale(sheet, pos)
        if mine is None:
            continue
        scales = {c: _row_scale(sheet, c) for c in cells}
        if any(s is None or s[0] != mine[0] for s in scales.values()):
            continue
        odd = [c for c, s in scales.items() if s[1] != mine[1]]
        if not odd or len(odd) * 2 >= len(cells):
            continue
        values = {c: sheet.inputs.get(c) for c in cells}
        if not all(isinstance(v, (int, float)) and not isinstance(v, bool) and v for v in values.values()):
            continue
        typical = statistics.median([abs(values[c]) for c in cells if c not in odd])
        if not all(1 / 30 <= abs(values[c]) * scales[c][1] / mine[1] / typical <= 30 for c in odd):
            continue
        terms = []
        for c in cells:
            ratio = scales[c][1] / mine[1]
            ref = f"{column_letters(c[1])}{c[0]}"
            terms.append(ref if ratio == 1 else f"{ref}/{_fmt(1 / ratio)}" if ratio < 1 else f"{ref}*{_fmt(ratio)}")
        expected = sum(values[c] * scales[c][1] / mine[1] for c in cells)
        key = (pos[0], tuple(sorted({r for r, _ in odd})))
        groups.setdefault(key, []).append((pos, odd, values, "SUM(" + ",".join(terms) + ")", expected))
    found: list[Finding] = []
    for (row, _), items in groups.items():
        fix = einsetzen.formulas(sheet, {pos: text for pos, _, _, text, _ in items})
        if fix is None:
            continue
        pos, odd, values, _, expected = items[0]
        listed = ", ".join(f"{address(c)} in {_unit_text(sheet, c)} ({_fmt(values[c])})" for c in odd)
        first = sheet.cells.get((pos[0], 1))
        name = f" ({first.value.strip()})" if first is not None and isinstance(first.value, str) and first.value.strip() \
            and pos[1] > 1 else ""
        also = ", ".join(address(p) for p, *_ in items[1:])
        more = T(f" Genauso in {also}.", f" The same in {also}.") if len(items) > 1 else ""
        verb = "is" if len(odd) == 1 else "are"
        found.append(Finding(
            f"unitmix:{address(pos)}",
            T(f"{address(pos)}{name} zählt Zahlen in {_unit_text(sheet, pos)} zusammen, aber {listed} steht in einer "
              f"anderen Einheit und zählt unumgerechnet mit. So zeigt {address(pos)} {_fmt(sheet.inputs.get(pos))} statt "
              f"{_fmt(expected)}.{more}",
              f"{address(pos)}{name} adds up numbers in {_unit_text(sheet, pos)}, but {listed} {verb} in a different "
              f"unit and {verb} counted without conversion. So {address(pos)} shows {_fmt(sheet.inputs.get(pos))} "
              f"instead of {_fmt(expected)}.{more}"),
            "fehler", T("Rechne die Zeile in die Einheit der Summe um.", "Convert the row to the unit of the total."),
            fix=fix))
        if len(found) >= MAX_EACH:
            break
    return found


def _unit_text(sheet: Sheet, pos: Pos) -> str:
    for k in range(pos[1] - 1, 0, -1):
        cell = sheet.cells.get((pos[0], k))
        if cell is not None and not cell.formula and isinstance(cell.value, str) and cell.value.strip() \
                and not isinstance(sheet.inputs.get((pos[0], k)), (int, float)):
            return cell.value.strip()
    return ""


# --- Ein Satz, der genau ein Zehntel oder das Zehnfache der übrigen Jahre ist ---

_RATE_NAME = re.compile(r"satz|quote|rate|anteil|marge|%", re.IGNORECASE)


def shifted_rates(sheet: Sheet) -> list[Finding]:
    """„Ertragsteuersatz %“ 0,30 im ersten Jahr und 0,03 im zweiten: genau ein Zehntel, ein verrutschtes Komma? Nur
    Eingaben, die eine Formel liest; mit der Folge in der Formel."""
    found: list[Finding] = []
    rows: dict[int, list[Pos]] = {}
    for pos, cell in sheet.cells.items():
        value = sheet.inputs.get(pos)
        if cell.formula or pos in sheet.nodes or isinstance(value, bool) or not isinstance(value, (int, float)):
            continue
        if 0 < abs(value) <= 1.5:
            rows.setdefault(pos[0], []).append(pos)
    for row, cells in sorted(rows.items()):
        if len(cells) < 2 or len(found) >= MAX_EACH:
            continue
        cells.sort()
        label = " ".join(str(sheet.cells[(row, c)].value) for c in range(1, cells[0][1]) if (row, c) in sheet.cells
                         and not sheet.cells[(row, c)].formula and not isinstance(sheet.inputs.get((row, c)), (int, float)))
        if not _RATE_NAME.search(label):
            continue
        values = [sheet.inputs[c] for c in cells]
        for i, pos in enumerate(cells):
            others = values[:i] + values[i + 1:]
            if len(set(others)) != 1:
                continue
            base, mine = others[0], values[i]
            factor = next((f for f in (10, 100, 0.1, 0.01) if math.isclose(mine * f, base, rel_tol=1e-9)), None)
            users = sorted(sheet.dependents.get(pos, ()))
            if factor is None or not users or len(cells) > 2 and len(set(values)) != 2:
                continue
            if len(cells) == 2 and i == 0:
                continue  # Bei zwei Jahren weicht das spätere ab.
            near = cells[i - 1] if i else cells[1]
            word = T("ein Zehntel", "a tenth of") if factor == 10 else T("ein Hundertstel", "a hundredth of") \
                if factor == 100 else T("das Zehnfache", "ten times") if factor == 0.1 else \
                T("das Hundertfache", "a hundred times")
            user = users[0]
            calc, _ = sheet.evaluator()
            consequence = ""
            node = sheet.nodes.get(user)
            if node is not None:
                saved = sheet.inputs[pos]
                try:
                    sheet.inputs[pos] = base
                    _, value = sheet.evaluator(chain=set(sheet.nodes))
                    meant = value(user)
                except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError):
                    meant = None
                finally:
                    sheet.inputs[pos] = saved
                now = sheet.inputs.get(user)
                if isinstance(meant, (int, float)) and isinstance(now, (int, float)) and not math.isclose(meant, now):
                    consequence = T(f" {address(user)} rechnet damit und zeigt {_fmt(now)}, mit {_fmt(base)} wären es {_fmt(meant)}.",
                                    f" {address(user)} uses it and shows {_fmt(now)}; with {_fmt(base)} it would be "
                                    f"{_fmt(meant)}.")
            found.append(Finding(
                f"rateshift:{address(pos)}",
                T(f"{address(pos)} ({label.strip()}) ist mit {_fmt(mine)} genau {word} von {address(near)} ({_fmt(base)}). "
                  f"Ist das Komma verrutscht?{consequence}",
                  f"{address(pos)} ({label.strip()}) is {_fmt(mine)} and thus exactly {word} {address(near)} ({_fmt(base)}). "
                  f"Is the decimal point in the wrong place?{consequence}"),
                "pruefen", T(f"Prüf den Wert in {address(pos)}; ist es derselbe Satz, schreib {_fmt(base)}.",
                             f"Check the value in {address(pos)}; if it is the same rate, enter {_fmt(base)}.")))
            break
    return found
