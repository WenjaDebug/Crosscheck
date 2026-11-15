"""Modelle mit Anteilen, Annahmen und eigenen Prüfungen: was dort typisch schiefgeht, ohne KI.

- Anteile, die 100 % ergeben müssen: Ergeben in einer Übergangsmatrix (oder einer Verteilung je Zeile) alle Zeilen
  genau 100 % und eine nicht, ist dort eine Zahl falsch. Der Hinweis sagt, wie viel zu viel oder zu wenig es ist und
  was davon abhängt (die Verteilung der Folgejahre, die eigene Prüfung des Modells).
- Abzinsen mit der falschen Annahme: =C29/(1+$B$6)^A29 nimmt die Verlustquote statt des Diskontsatzes in B7, und B7
  liest keine Formel. Die Korrektur setzt die richtige Annahme ein, in allen Zellen, die so rechnen.
- Die eigene Prüfung des Modells (=WENN(ABS(F26-1)<0,0001;"OK";"Fehler")) zeigt „Fehler“: Er sagt, welche Zelle
  sie prüft und was dort steht, wenn kein genauerer Hinweis die Ursache schon nennt.
"""

from __future__ import annotations

import math
import re

from . import einsetzen, formel
from .analyse import Finding, Pos, Sheet, _fmt, _german, address
from .formel import Node
from .sprache import T
from .workbook import _left_label
from .xlsx import column_letters, number_text
from .zeilenmodell import _to_cell

MAX_FOUND = 3
# So genau muss eine Zeile 100 % ergeben, damit sie als „genau 100 %“ zählt.
_EXACT = 1e-9
# Ab so viel Abweichung ist eine Zeile falsch (nicht nur gerundet).
_OFF = 1e-6
# Wörter für einen Zins zum Abzinsen.
_RATE = re.compile(r"diskont|abzins|zins|wacc|kapitalkost|discount|interest|rendite|hurdle", re.IGNORECASE)
_OK = {"ok", "o.k.", "stimmt", "passt", "korrekt", "richtig", "✓", "✔", "true", "wahr", "ja", "yes"}
_BAD = re.compile(r"fehler|falsch|error|prüfen|pruefen|abweich|nicht ok|✗|✘|achtung|stimmt nicht|check", re.IGNORECASE)


def checks(sheet: Sheet, before: list[Finding]) -> list[Finding]:
    """Die neuen Hinweise. Zeigt die eigene Prüfung des Modells „Fehler“ wegen Fehlern, die Hinweise schon nennen
    (in before oder den neuen), sagen es diese Hinweise mit, statt eines eigenen."""
    found = _hundred(sheet, before)
    found += _discount(sheet)
    found += _lenient(sheet)
    for item in _self_checks(sheet):
        at = item.key.partition(":")[2]
        if any(_named(f.text, at) for f in found + before):
            continue
        causes = _causes(sheet, formel.position(at), found + before)
        shown = sheet.inputs.get(formel.position(at))
        if causes and _still_failing(sheet, formel.position(at), causes):
            # Mit allen Korrekturen dieser Hinweise zeigt die Prüfung immer noch „Fehler“: Es gibt noch eine Ursache,
            # die keiner nennt. Dann bleibt der eigene Hinweis.
            item.text += T(" Die übrigen Hinweise erklären das nicht ganz: Auch mit ihren Korrekturen bleibt sie dabei.",
                           " The other notes don't fully explain this: even with their fixes, the check still fails.")
            found.append(item)
            continue
        causes = _needed(sheet, formel.position(at), causes)
        for cause in causes:
            cause.text += (T(f" Darum zeigt auch die eigene Prüfung des Modells in {at} „{shown}“.",
                             f" That is also why the model's own check in {at} shows “{shown}”.") if len(causes) == 1
                           else T(f" Auch deshalb zeigt die eigene Prüfung des Modells in {at} „{shown}“.",
                                  f" This is also one reason why the model's own check in {at} shows “{shown}”."))
        if not causes:
            found.append(item)
    return found


def _needed(sheet: Sheet, pos: Pos, causes: list[Finding]) -> list[Finding]:
    """Nur die Ursachen, ohne deren Korrektur die Prüfung wieder „Fehler“ zeigt. Ein falscher Kurs, der beide Seiten
    eines Vergleichs gleich ändert, ist keine Ursache für dessen „Fehler“."""
    fixed = [c for c in causes if getattr(c, "fix", None)]
    if not fixed or len(causes) < 2:
        return causes
    if len(fixed) < len(causes) and _still_failing(sheet, pos, fixed):
        return causes  # Ohne die Hinweise ohne Korrektur geht die Prüfung nicht auf: Sie zählen mit.
    # Die Hinweise ohne Korrektur braucht es nicht, wenn die übrigen schon reichen.
    kept = [c for c in fixed if len(fixed) == 1 or not _passes(sheet, pos, [o for o in fixed if o is not c])]
    return kept or causes


def _passes(sheet: Sheet, pos: Pos, causes: list[Finding]) -> bool:
    """Zeigt die Prüfung mit den Korrekturen dieser Hinweise sicher „OK“? Unklar zählt als nein."""
    from .sensitivitaet import _with_fixes

    found = _branches(sheet.nodes.get(pos))
    after = _with_fixes(sheet, causes, [pos]) if causes else None
    return found is not None and after is not None and bool(after[0]) and after[1][0] == found[0]


def _still_failing(sheet: Sheet, pos: Pos, causes: list[Finding]) -> bool:
    """Zeigt die Prüfung in pos auch mit allen Korrekturen dieser Hinweise noch ihr „Fehler“? Ohne Korrekturen zum
    Nachrechnen: nein (wie bisher)."""
    from .sensitivitaet import _with_fixes

    einsetzen.attach(sheet, causes)
    # Hinweise ohne Korrektur lassen sich nicht nachrechnen: Dann gilt die Prüfung als erklärt.
    if not all(getattr(c, "fix", None) for c in causes):
        return False
    after = _with_fixes(sheet, causes, [pos])
    if after is None or not after[0]:
        return False
    found = _branches(sheet.nodes.get(pos))
    return found is not None and after[1][0] == found[1]


def _named(text: str, where: str) -> bool:
    return re.search(rf"(?<![A-Z0-9]){where}(?![0-9])", text) is not None


def _upstream(sheet: Sheet, pos: Pos, limit: int = 5000) -> set[Pos]:
    seen, todo = set(), [pos]
    while todo and len(seen) < limit:
        here = todo.pop()
        before = set(sheet.precedents.get(here, ()))
        for area in sheet.ranges.get(here, ()):
            if (area[2] - area[0] + 1) * (area[3] - area[1] + 1) <= limit:
                before.update(_cells(area))
        for p in before:
            if p not in seen:
                seen.add(p)
                todo.append(p)
    return seen


def _causes(sheet: Sheet, check: Pos, findings: list[Finding]) -> list[Finding]:
    """Die Fehler, deren Zelle die Prüfung (über ihre Formeln) liest."""
    earlier = _upstream(sheet, check)
    rows = sheet.memo.get("modell.rows", {})
    return [item for item in findings if item.level == "fehler" and (at := _at(item)) is not None
            and (at in earlier or earlier & rows.get(item.key, set()))]


def _at(item: Finding) -> Pos | None:
    tail = item.key.partition(":")[2]
    return formel.position(tail) if re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]*", tail) else None


def _pct(share: float) -> str:
    return f"{number_text(round(share * 100, 2))} %"


def _value(sheet: Sheet, pos: Pos):
    v = sheet.inputs.get(pos)
    return v if isinstance(v, float) and not isinstance(v, bool) and math.isfinite(v) else None


# --- Anteile, die 100 % ergeben müssen ---


def _sum_area(sheet: Sheet, pos: Pos) -> tuple[int, int, int, int] | None:
    """Der Bereich von =SUMME(B11:E11), wenn er eine Zeile oder Spalte ist."""
    node = sheet.nodes.get(pos)
    if node is None or node.kind != "call" or node.value != "SUM" or len(node.args) != 1 or node.args[0].kind != "range":
        return None
    r1, c1, r2, c2 = node.args[0].value
    if (r1 == r2) == (c1 == c2) or max(r2 - r1, c2 - c1) < 1:
        return None
    return r1, c1, r2, c2


def _cells(area: tuple[int, int, int, int]) -> list[Pos]:
    r1, c1, r2, c2 = area
    return [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]


def _groups(sheet: Sheet) -> list[list[tuple[Pos, tuple[int, int, int, int]]]]:
    """Summen untereinander (je eine Zeile) oder nebeneinander (je eine Spalte), die gleich gebaut sind."""
    sums = {pos: area for pos in sheet.nodes if (area := _sum_area(sheet, pos)) is not None}
    groups: dict[tuple, list] = {}
    for pos, (r1, c1, r2, c2) in sums.items():
        if r1 == r2 and r1 == pos[0]:
            key = ("row", pos[1], c1 - pos[1], c2 - pos[1])
        elif c1 == c2 and c1 == pos[1]:
            key = ("col", pos[0], r1 - pos[0], r2 - pos[0])
        else:
            continue
        groups.setdefault(key, []).append((pos, (r1, c1, r2, c2)))
    out = []
    for key, items in groups.items():
        items.sort(key=lambda it: it[0][0] if key[0] == "row" else it[0][1])
        run: list = []
        for item in items:
            step = item[0][0] if key[0] == "row" else item[0][1]
            if run and step != (run[-1][0][0] if key[0] == "row" else run[-1][0][1]) + 1:
                out.append(run)
                run = []
            run.append(item)
        out.append(run)
    return [run for run in out if len(run) >= 3]


def _shares(sheet: Sheet, run: list, target: float) -> bool:
    values = [_value(sheet, pos) for pos, _ in run]
    if not any(v is not None and abs(v - target) <= target * _EXACT for v in values):
        return False
    parts = [_value(sheet, p) for _, area in run for p in _cells(area)]
    return all(v is not None and 0 <= v <= target for v in parts)


def _downstream(sheet: Sheet, start: list[Pos], limit: int = 2000) -> list[Pos]:
    seen, todo = set(start), list(start)
    out = []
    while todo and len(out) < limit:
        for after in sheet.dependents.get(todo.pop(), ()) or ():
            if after not in seen:
                seen.add(after)
                out.append(after)
                todo.append(after)
    return out


def _hundred(sheet: Sheet, before: list[Finding]) -> list[Finding]:
    found = []
    groups = _groups(sheet)
    # Summen, die ebenfalls 100 % ergeben sollten: In ihrer Gruppe ergibt eine genau 100 %, alle Teile sind Anteile.
    totals = {pos for run in groups for pos, _ in run if _shares(sheet, run, 1.0)} | \
             {pos for run in groups for pos, _ in run if _shares(sheet, run, 100.0)}
    for run in groups:
        values = {pos: _value(sheet, pos) for pos, _ in run}
        if any(v is None for v in values.values()):
            continue
        target = 1.0 if sum(abs(v - 1) <= _EXACT for v in values.values()) * 2 > len(run) else None
        if target is None and sum(abs(v - 100) <= 100 * _EXACT for v in values.values()) * 2 > len(run):
            target = 100.0
        if target is None:
            continue
        exact = [pos for pos, v in values.items() if abs(v - target) <= target * _EXACT]
        if len(exact) < 2:
            continue
        for pos, area in run:
            v = values[pos]
            parts = _cells(area)
            numbers = [_value(sheet, p) for p in parts]
            if abs(v - target) <= target * _OFF or any(n is None or n < 0 or n > target for n in numbers):
                continue
            # Die Ursache steht in der Zeile selbst (eine getippte Zahl), nicht in einer Formel davor.
            if all(p in sheet.nodes for p in parts):
                continue
            others = {at for f in before if f.level == "fehler" and (at := _at(f)) is not None}
            found.append(_hundred_finding(sheet, pos, area, v, target, exact, totals, others))
            if len(found) >= MAX_FOUND:
                return found
    return found


def _name(sheet: Sheet, area: tuple[int, int, int, int]) -> str:
    """„„C““ für die Zeile C einer Matrix (der Name links daneben), sonst die Zeilennummer oder der Spaltenbuchstabe."""
    r1, c1, r2, _ = area
    if r1 == r2:
        name = _left_label(sheet.cells, r1, c1)
        return T(f"„{name}“", f"“{name}”") if name else str(r1)
    head = sheet.cells.get((r1 - 1, c1))
    name = head.value.strip() if head is not None and head.formula is None and head.value else ""
    return T(f"„{name}“", f"“{name}”") if name else column_letters(c1)


def _hundred_finding(sheet: Sheet, pos: Pos, area, v: float, target: float, exact: list[Pos], totals: set[Pos],
                     others: set[Pos]) -> Finding:
    span = f"{address(area[:2])}:{address(area[2:])}"
    shown = _pct(v) if target == 1 else _fmt(v)
    want = "100 %" if target == 1 else "100"
    gap = v - target
    amount = number_text(round(abs(gap) * (100 if target == 1 else 1), 4))
    much = T((f"{amount} Prozentpunkte" if target == 1 else amount) + (" zu viel" if gap > 0 else " zu wenig"),
             (f"{amount} percentage points" if target == 1 else amount) + (" too high" if gap > 0 else " too low"))
    row = area[0] == area[2]
    names = [_name(sheet, _sum_area(sheet, p)) for p in exact]
    listed = names[0] if len(names) == 1 else T(f"{', '.join(names[:-1])} und {names[-1]}",
                                                f"{', '.join(names[:-1])} and {names[-1]}")
    one = len(names) == 1
    kind = ("Zeile" if row else "Spalte") + ("" if len(names) == 1 else ("n" if row else "n"))
    verb = "ergibt" if len(names) == 1 else "ergeben"
    text = T(f"{'Zeile' if row else 'Spalte'} {_name(sheet, area)} ergibt in {address(pos)} {shown} statt {want}: "
             f"Die Zahlen in {span} sind zusammen {much}. Die {kind} {listed} {verb} genau {want}.",
             f"{'Row' if row else 'Column'} {_name(sheet, area)} adds up to {shown} in {address(pos)} instead of "
             f"{want}: together the numbers in {span} are {much}. {'Row' if row else 'Column'}{'' if one else 's'} "
             f"{listed} add{'s' if one else ''} up to exactly {want}.")
    # Folgen, die nur von dieser Zeile kommen: Liest eine Zelle auch einen anderen Fehler, sagt er nicht „darum“.
    later = [p for p in _downstream(sheet, _cells(area)) if not others or not (_upstream(sheet, p) & others)]
    off = sorted(p for p in later if p in totals and p != pos and (x := _value(sheet, p)) is not None
                 and abs(x - target) > target * _OFF)
    if off:
        last = off[-1]
        where = address(off[0]) if len(off) == 1 else T(f"{address(off[0])} bis {address(last)}",
                                                         f"{address(off[0])} to {address(last)}")
        x = _value(sheet, last)
        got = _pct(x) if target == 1 else _fmt(x)
        text += (T(f" Darum ergibt auch {where} {got} statt {want}.",
                   f" That is why {where} also adds up to {got} instead of {want}.") if len(off) == 1
                 else T(f" Darum ergeben auch {where} nicht {want} ({address(last)}: {got}).",
                        f" That is why {where} also don't add up to {want} ({address(last)}: {got})."))
    tip = T(f"Prüf die Zahlen in {span} gegen die Quelle. Welche davon falsch ist, sieht er nicht: Sie müssen zusammen "
            f"genau {want} ergeben.",
            f"Check the numbers in {span} against the source. Which of them is wrong can't be seen from here: together "
            f"they must add up to exactly {want}.")
    # Die Zahlen der Zeile, nicht die Summe, rechnen weiter (in die Folgejahre und die eigene Prüfung).
    sheet.memo.setdefault("modell.rows", {})[f"hundred:{address(pos)}"] = set(_cells(area))
    return Finding(f"hundred:{address(pos)}", text, "fehler", tip)


# --- Die eigene Prüfung des Modells ---


def _branches(node: Node | None) -> tuple[str, str] | None:
    """=WENN(Bedingung;"OK";"Fehler") → ("OK", "Fehler"): was die Prüfung zeigt, wenn es stimmt und wenn nicht."""
    if node is None or node.kind != "call" or node.value != "IF" or len(node.args) != 3:
        return None
    a, b = node.args[1], node.args[2]
    if a.kind != "str" or b.kind != "str":
        return None
    if a.value.strip().casefold() in _OK and _BAD.search(b.value):
        return a.value, b.value
    if b.value.strip().casefold() in _OK and _BAD.search(a.value):
        return b.value, a.value
    return None


def _failing(sheet: Sheet, pos: Pos) -> bool:
    found = _branches(sheet.nodes.get(pos))
    return found is not None and sheet.inputs.get(pos) == found[1]


def _self_checks(sheet: Sheet) -> list[Finding]:
    found = []
    for pos in sorted(sheet.nodes):
        if not _failing(sheet, pos):
            continue
        read = sorted({p for p in sheet.precedents.get(pos, ()) if p != pos})
        label = _left_label(sheet.cells, *pos)
        named = f"{address(pos)} ({label})" if label else address(pos)
        parts = [T(f"{address(p)} zeigt {_shown(sheet, p)}", f"{address(p)} shows {_shown(sheet, p)}") for p in read[:3]]
        text = T(f"Die eigene Prüfung des Modells in {named} zeigt „{sheet.inputs.get(pos)}“.",
                 f"The model's own check in {named} shows “{sheet.inputs.get(pos)}”.")
        if parts:
            cells = ", ".join(address(p) for p in read[:3])
            text += T(f" Sie liest {cells}: ", f" It reads {cells}: ") + T(", ", "; ").join(parts) + "."
        found.append(Finding(f"selfcheck:{address(pos)}", text, "fehler",
                             T("Such die Ursache in den Zellen, die sie liest, und in deren Eingaben.",
                               "Look for the cause in the cells it reads and in their inputs.")))
        if len(found) >= MAX_FOUND:
            break
    return found


_EXACT_WORDS = re.compile(r"ohne\s+(?:rundungs)?toleranz|keine\s+(?:rundungs)?toleranz|ganz genau|exakt|auf den cent|"
                          r"without tolerance|no tolerance|exact(?:ly)?\b", re.IGNORECASE)
_LOOSE = re.compile(r"ROUND\(((?:[^()]|\([^()]*\))+),\s*-?\d+\)\s*=\s*0(?![\d.])|ABS\(((?:[^()]|\([^()]*\))+)\)\s*<=?\s*[\d.]+")


def _lenient(sheet: Sheet) -> list[Finding]:
    """„Bankabgleich (ohne Rundungstoleranz)“ = WENN(RUNDEN(F49;0)=0;"OK";…): Die Prüfung rundet doch und zeigt OK,
    obwohl F49 0,4 zeigt. Gemeldet, wenn der Name keine Toleranz will und die Abweichung nicht 0 ist."""
    found = []
    calc, _ = sheet.evaluator()
    for pos in sorted(sheet.nodes):
        found_branches = _branches(sheet.nodes.get(pos))
        if found_branches is None:
            continue
        try:
            shows = calc.run(sheet.nodes[pos])
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, KeyError, IndexError):
            continue
        if shows != found_branches[0]:
            continue
        label = " ".join(str(sheet.cells[(pos[0], c)].value).strip() for c in range(1, pos[1])
                         if (pos[0], c) in sheet.cells and not sheet.cells[(pos[0], c)].formula
                         and isinstance(sheet.cells[(pos[0], c)].value, str)).strip()
        if not label or not _EXACT_WORDS.search(label):
            continue
        formula = (sheet.cells[pos].formula or "").lstrip("=")
        m = _LOOSE.search(formula)
        if m is None:
            continue
        inner = m.group(1) or m.group(2)
        try:
            off = calc.run(formel.parse(inner))
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, KeyError, IndexError):
            continue
        if not isinstance(off, (int, float)) or isinstance(off, bool) or off == 0:
            continue
        strict = formula[:m.start()] + f"{inner}=0" + formula[m.end():]
        fix = einsetzen.formulas(sheet, {pos: strict})
        function = m.group(0).split('(')[0]
        found.append(Finding(
            f"lenient:{address(pos)}",
            T(f"{address(pos)} ({label}) soll ohne Toleranz prüfen, lässt aber mit {_german(function)}(…) eine "
              f"Abweichung durchgehen: {inner} ist {_fmt(off)}, und die Prüfung zeigt trotzdem „{found_branches[0]}“.",
              f"{address(pos)} ({label}) is supposed to check without tolerance, but lets a difference through with "
              f"{function}(…): {inner} is {_fmt(off)} and the check still shows “{found_branches[0]}”."),
            "fehler", T(f"Prüf genau auf 0: ={_german_text(strict)}.", f"Check for exactly 0: ={_german_text(strict)}."),
            fix=fix))
        if len(found) >= MAX_FOUND:
            break
    return found


def _german_text(english: str) -> str:
    from .einsetzen import _german_formula

    return _german_formula(english)


def _shown(sheet: Sheet, pos: Pos) -> str:
    v = sheet.inputs.get(pos)
    return _fmt(v) if v is not None else T("nichts", "nothing")


# --- Abzinsen mit der falschen Annahme ---


def _plus_one(node: Node) -> Pos | None:
    """(1+$B$7) → B7."""
    if node.kind == "op" and node.value == "+" and len(node.args) == 2:
        a, b = node.args
        if a.kind == "num" and a.value == 1.0 and b.kind == "ref":
            return b.value
        if b.kind == "num" and b.value == 1.0 and a.kind == "ref":
            return a.value
    return None


def _rates(node: Node) -> list[Pos]:
    """Die Zellen, mit denen die Formel abzinst: /(1+R)^t, /(1+R), *(1+R)^-t, NBW(R;…), BW(R;…)."""
    out = []
    if node.kind == "op" and node.value == "/" and len(node.args) == 2:
        right = node.args[1]
        base = right.args[0] if right.kind == "op" and right.value == "^" else right
        rate = _plus_one(base)
        if rate is not None:
            out.append(rate)
    if node.kind == "op" and node.value == "^" and len(node.args) == 2 and node.args[1].kind == "neg":
        rate = _plus_one(node.args[0])
        if rate is not None:
            out.append(rate)
    if node.kind == "call" and node.value in ("NPV", "PV", "XNPV") and node.args and node.args[0].kind == "ref":
        out.append(node.args[0].value)
    for arg in node.args if isinstance(node.args, tuple) else ():
        if isinstance(arg, Node):
            out += _rates(arg)
    return out


def _input(sheet: Sheet, pos: Pos) -> bool:
    v = _value(sheet, pos)
    return pos not in sheet.nodes and v is not None


def _block(sheet: Sheet, pos: Pos) -> list[Pos]:
    """Die Annahmen untereinander, in deren Block pos steht (Zahlen ohne Formel, je mit einem Namen links)."""
    r, c = pos
    top = r
    while top > 1 and _input(sheet, (top - 1, c)):
        top -= 1
    bottom = r
    while _input(sheet, (bottom + 1, c)):
        bottom += 1
    return [(row, c) for row in range(top, bottom + 1)]


def _discount(sheet: Sheet) -> list[Finding]:
    users: dict[Pos, list[Pos]] = {}
    for pos, node in sheet.nodes.items():
        if node is None:
            continue
        for rate in dict.fromkeys(_rates(node)):
            users.setdefault(rate, []).append(pos)
    found = []
    for rate, cells in sorted(users.items()):
        if not _input(sheet, rate):
            continue
        label = _left_label(sheet.cells, *rate)
        if not label or _RATE.search(label):
            continue
        spare = [p for p in _block(sheet, rate) if p != rate and not sheet.dependents.get(p)
                 and _RATE.search(_left_label(sheet.cells, *p)) and 0 < _value(sheet, p) < 1]
        if len(spare) != 1:
            continue
        found += _discount_findings(sheet, rate, label, spare[0], sorted(cells, key=lambda p: (p[1], p[0])))
        if len(found) >= MAX_FOUND:
            break
    return found


def _runs(cells: list[Pos]) -> list[list[Pos]]:
    out: list[list[Pos]] = []
    for pos in cells:
        if out and pos == (out[-1][-1][0] + 1, out[-1][-1][1]):
            out[-1].append(pos)
        else:
            out.append([pos])
    return out


def _rate_text(sheet: Sheet, pos: Pos, label: str) -> str:
    v = _value(sheet, pos)
    return f"{address(pos)} ({label}, {_pct(v) if v is not None and abs(v) < 1 else _fmt(v)})"


def _discount_findings(sheet: Sheet, rate: Pos, label: str, spare: Pos, cells: list[Pos]) -> list[Finding]:
    calc, _ = sheet.evaluator()
    found = []
    for run in _runs(cells):
        changes = {p: _to_cell(sheet.cells[p].formula, rate, spare) for p in run}
        fix = einsetzen.formulas(sheet, changes)
        if fix is None:
            continue
        first = run[0]
        span = address(first) if len(run) == 1 else f"{address(first)}:{address(run[-1])}"
        plan = sheet.plan
        head = plan.headers.get(first[1], "") if plan.in_table(*first) else _left_label(sheet.cells, *first)
        named = f"{span} ({head})" if head else span
        verb = "zinst" if len(run) == 1 else "zinsen"
        used, meant = _rate_text(sheet, rate, label), _rate_text(sheet, spare, _left_label(sheet.cells, *spare))
        text = T(f"{named} {verb} mit {used} ab statt mit "
                 f"{meant}. {address(spare)} liest keine Formel.",
                 f"{named} {'discounts' if len(run) == 1 else 'discount'} with {used} instead of {meant}. No formula "
                 f"reads {address(spare)}.")
        try:
            want = calc.run(formel.parse(changes[first]))
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError):
            want = None
        now = _value(sheet, first)
        if isinstance(want, float) and now is not None and want != now:
            text += T(f" {address(first)} zeigt darum {_fmt(now)} statt {_fmt(want)}.",
                      f" {address(first)} therefore shows {_fmt(now)} instead of {_fmt(want)}.")
        german = fix["german"]
        tip = T(f"Korrigiert: {german}.", f"Corrected: {german}.") if len(run) == 1 else \
            T(f"Korrigiert in {address(first)}: {german}, dann bis {address(run[-1])} herunterziehen.",
              f"Corrected in {address(first)}: {german}, then fill down to {address(run[-1])}.")
        found.append(Finding(f"discount:{address(first)}", text, "fehler", tip, fix=fix))
    return found
