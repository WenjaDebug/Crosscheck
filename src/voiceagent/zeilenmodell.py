"""Zeilenmodelle: Finanzpläne, Budgets und Prognosen, in denen jede Zeile eine Größe ist (Umsatz, EBITDA) und jede
Spalte ein Jahr (Layout.across). Dort rechnet eine Zeile über die Jahre gleich, nicht eine Spalte nach unten.

Die typischen Fehler solcher Modelle, jeder mit der Korrektur zum Einsetzen:
- Eine feste Zahl in einem Jahr, wo die übrigen Jahre derselben Zeile mit Formel rechnen (der Umsatz 2028E
  eingetippt). Sie rechnet nicht mit, wenn sich eine Annahme ändert.
- Eine Annahme mit $ festgehalten: Jedes Jahr rechnet mit dem Wachstum von 2026E, und die eigenen Werte der übrigen
  Jahre liest keine Formel.
- Eine SUMME, die in jedem Jahr eine Zeile ihres Abschnitts auslässt (der Operative Cashflow ohne die Veränderung
  des Working Capital): ein Hinweis für die ganze Zeile statt einer je Jahr. Fehlt die Zeile nur in einem Jahr,
  sagt es schon die Prüfung der Formelmuster (analyse._patterns).
"""

from __future__ import annotations

import math
import re
from collections import Counter

from . import einsetzen, formel
from .analyse import MAX_LISTED_MOVED, PATTERN_SHARE, Finding, Pos, Sheet, _checked, _moved_text, _sensible, _shows, address
from .sprache import T
from .workbook import pattern
from .xlsx import ANY_REF, column_letters, column_number, outside_literals, shift

# So viele Hinweise je Prüfung höchstens.
MAX_FOUND = 3


def checks(sheet: Sheet, before: list[Finding]) -> tuple[list[Finding], set[str]]:
    """Die Hinweise für ein Zeilenmodell und die Schlüssel der Hinweise aus before, die sie ersetzen."""
    if not sheet.plan.across:
        return typed_down(sheet) + _column_gaps(sheet)[:MAX_FOUND], set()
    patterns = {f.key.partition(":")[2]: f for f in before if f.key.startswith("pattern:")}
    gaps, covered = [], set()
    for item in _gaps(sheet):
        at = item.key.partition(":")[2]
        cells = set((item.fix or {}).get("cells") or {}) | {at}
        # Fehlt die Zeile nur in einem Jahr, sagt der Hinweis zum Muster dasselbe, mit dem, was die Zelle zeigt.
        if cells == {at} and at in patterns:
            continue
        gaps.append(item)
        # Lassen die meisten Jahre die Zeile aus, hielte das Muster die vollständige Summe für den Fehler.
        row = formel.position(at)[0]
        covered |= {key for key in (f"pattern:{c}" for c in patterns) if formel.position(key[8:])[0] == row
                    and key[8:] not in cells and _is_sum(sheet, formel.position(key[8:]))}
    return (_typed(sheet)[:MAX_FOUND] + _frozen(sheet)[:MAX_FOUND] + gaps[:MAX_FOUND] + _short_lines(sheet)[:MAX_FOUND]
            + _mixed_years(sheet)[:MAX_FOUND] + _walks(sheet)[:MAX_FOUND] + _drift(sheet)[:MAX_FOUND]
            + _npv_today(sheet)[:MAX_FOUND]), covered


def _column_gaps(sheet: Sheet) -> list[Finding]:
    """Eine SUMME unter ihrer Spalte, die eine Zeile der Tabelle auslässt (Dez in =SUMME(C18:C28)): mit der
    nachgerechneten Korrektur und dem, was die Summe darum zeigt. Ohne Korrektur sagt es task.py allein."""
    from .task import sum_gaps

    plan = sheet.plan
    if plan.header_row is None or plan.last_row is None:
        return []
    rows = list(range(plan.header_row + 1, plan.last_row + 1))
    found: list[Finding] = []
    for row, col, top, bottom, left_out in sum_gaps(sheet.cells, plan, rows):
        if not left_out or (row, col) not in sheet.nodes:
            continue
        start, end = min(top, *left_out), max(bottom, *left_out)
        # Nur wenn der längere Bereich nichts anderes dazunimmt als die ausgelassenen Zeilen.
        if any(_number(sheet, (r, col)) is not None for r in range(start, end + 1)
               if not top <= r <= bottom and r not in left_out):
            continue
        fix = einsetzen.stretched(sheet, [(row, col)], "SUM", (top, col, bottom, col), (start, end))
        if fix is None:
            continue
        pos = (row, col)
        changed = fix["cells"][address(pos)][1:]
        names = _listed([plan.word(r) or T(f"Zeile {r}", f"row {r}") for r in left_out])
        now, right = sheet.readable(pos), sheet.readable(pos, changed)
        text = T(f"Die Summe in {sheet.label(pos)} zählt {names} nicht mit: {now} statt {right}.",
                 f"The sum in {sheet.label(pos)} does not include {names}: {now} instead of {right}.")
        text = _moved_text(sheet, text, [(r, col) for r in left_out][:MAX_LISTED_MOVED], len(left_out), [], 0)
        text += _shows(sheet, pos, changed)
        checked = _checked(sheet, changed)
        found.append(Finding(f"gap:{address(pos)}", text, "fehler",
                             T(f"Korrigiert: {checked}.", f"Corrected: {checked}.") if checked else
                             T("Zieh den Bereich bis zur letzten Zeile.", "Extend the range to the last row."), fix=fix))
    return found


def _is_sum(sheet: Sheet, pos: Pos) -> bool:
    node = sheet.nodes.get(pos)
    return node is not None and node.kind == "call" and node.value == "SUM"


def _columns(sheet: Sheet) -> list[int]:
    plan = sheet.plan
    return sorted(c for c in plan.headers if c != plan.name_column)


def _rows(sheet: Sheet) -> range:
    plan = sheet.plan
    return range(plan.header_row + 1, plan.last_row + 1)


def _number(sheet: Sheet, pos: Pos) -> float | None:
    value = sheet.inputs.get(pos)
    if isinstance(value, float) and not isinstance(value, bool) and math.isfinite(value):
        return value
    return None


def _formulas(sheet: Sheet, r: int) -> list[int]:
    return [c for c in _columns(sheet) if (cell := sheet.cells.get((r, c))) is not None and cell.formula
            and sheet.nodes.get((r, c)) is not None and (r, c) not in sheet.unknown]


def kind(header: str) -> str:
    """Die Art einer Spalte: „2028E“ und „2027E“ sind dieselbe, „Jan“ und „Feb“ auch (Monate), „Summe 2026-2030“,
    „CAGR“ oder „Budget“ eine andere."""
    from .bereiche import _MONTHS, _QUARTERS

    out = []
    from . import woerter

    for token in re.findall(r"[^\W_]+", woerter.spaced(header).strip().casefold()):
        if token in _MONTHS or token in _QUARTERS:
            out.append("m")
        else:
            out.append(re.sub(r"\d+", "9", re.sub(r"[^\W\d_]+", "a", token)))
    return " ".join(out)


_kind = kind


def _num(value: float) -> str:
    from .bereiche import _num

    return _num(value)


def _value(sheet: Sheet, formula: str):
    calc, _ = sheet.evaluator()
    try:
        return calc.run(formel.parse(formula))
    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError, IndexError, KeyError):
        return None


def _span(sheet: Sheet, r: int, columns: list[int]) -> str:
    return f"{address((r, columns[0]))}:{address((r, columns[-1]))}"


def _listed(items: list[str]) -> str:
    """„A, B und C“ (englisch „A, B and C“)."""
    return items[0] if len(items) == 1 else ", ".join(items[:-1]) + T(f" und {items[-1]}", f" and {items[-1]}")


def _years(sheet: Sheet, columns: list[int]) -> str:
    return _listed([sheet.plan.headers[c] for c in columns])


def _cells(cells: list[Pos]) -> str:
    return _listed([address(p) for p in cells])


# --- Eine feste Zahl, wo die übrigen Jahre rechnen ---


# Steht ein Jahr in so vielen Zeilen als feste Zahl da, ist es wohl als Werte eingefügt: ein Hinweis für die Spalte.
PASTED = 3


def _typed(sheet: Sheet) -> list[Finding]:
    """Feste Zahlen, wo die übrigen Jahre derselben Zeile mit Formel rechnen: je Zeile ein Hinweis, oder einer für
    ein ganzes Jahr, das als Werte eingefügt ist."""
    plan = sheet.plan
    columns = _columns(sheet)
    rows: dict[int, tuple[dict[Pos, str], list[int]]] = {}
    computed = _computed_columns(sheet, columns)
    for r in _rows(sheet):
        formulas = _formulas(sheet, r)
        if len(formulas) < 2:
            continue
        # Links der ersten Formel stehen die Ist-Zahlen (2025A): Die sind fest, das gehört so. Eine Spalte anderer
        # Art („Budget“, „Vorgabe“) ist eine Eingabe. Rechnet die Spalte aber in den übrigen Zeilen (Periode 1 ist
        # eine Prognose wie die anderen), ist eine feste Zahl dort genauso falsch wie in Periode 3.
        kinds = {_kind(plan.headers[c]) for c in formulas}
        typed = [c for c in columns if (c > formulas[0] or c in computed) and (r, c) in sheet.cells
                 and not sheet.cells[(r, c)].formula and _number(sheet, (r, c)) is not None
                 and _kind(plan.headers[c]) in kinds]
        if not typed or len(typed) >= len(formulas):
            continue
        shapes = {c: pattern(sheet.cells[(r, c)].formula, (r, c), "Zeile") for c in formulas}
        common, count = Counter(shapes.values()).most_common(1)[0]
        if count < 2 or count < PATTERN_SHARE * len(formulas):
            continue
        model = [c for c in formulas if shapes[c] == common]
        # Eine Zeile, die nur sich selbst fortschreibt (=C8, =C8*1,02), ist eine Reihe von Annahmen: Eine eingetippte
        # Zahl darin ist eine neue Annahme, kein vergessener Bezug.
        singles, areas, _ = formel.references(sheet.nodes[(r, model[0])], sheet.name)
        if all(p[0] == r for p in singles) and all(a[0] == a[2] == r for a in areas):
            continue
        changes: dict[Pos, str] = {}
        for c in typed:
            example = min(model, key=lambda m: (abs(m - c), m > c))
            expected = shift(sheet.cells[(r, example)].formula, 0, c - example)
            if not _sensible(sheet, expected) or not isinstance(_value(sheet, expected), float):
                changes = {}
                break
            changes[(r, c)] = expected
        if changes:
            rows[r] = changes, formulas
    pasted = Counter(c for changes, _ in rows.values() for _, c in changes)
    found: list[Finding] = []
    for col, count in sorted(pasted.items()):
        if count < PASTED:
            continue
        changes = {p: f for chosen, _ in rows.values() for p, f in chosen.items() if p[1] == col}
        for r, _ in changes:
            rows[r] = ({p: f for p, f in rows[r][0].items() if p[1] != col}, rows[r][1])
        found += _typed_finding(sheet, changes, None)
    for r, (changes, formulas) in sorted(rows.items()):
        if changes:
            found += _typed_finding(sheet, changes, formulas)
    return found


def _rounded(have: float, want: float) -> bool:
    """Ist have die Zahl want, nur auf die Stellen gerundet, mit denen have dasteht (12249,36 für 12249,3601)? Nicht
    0,1155 für 0,119: Das ist eine andere Zahl, auch wenn der Abstand klein ist."""
    if abs(have - want) <= 1e-9 * max(1.0, abs(want)):
        return True
    text = repr(float(have))
    if "e" in text:
        return False
    places = len(text.split(".")[1].rstrip("0")) if "." in text else 0
    return places <= 6 and abs(round(want, places) - have) <= 1e-9 * max(1.0, abs(want))


def _computed_columns(sheet: Sheet, columns: list[int]) -> set[int]:
    """Die Jahre, die im Modell rechnen: In fast allen Zeilen mit Formeln steht dort eine Formel. Eine Spalte mit
    Ist-Zahlen (2025A) ist fest, eine Prognosespalte nicht."""
    lines = [r for r in _rows(sheet) if len(_formulas(sheet, r)) >= 2]
    if len(lines) < 4:
        return set()
    out = set()
    for c in columns:
        have = [r for r in lines if (r, c) in sheet.cells]
        if len(have) >= 4 and sum(1 for r in have if sheet.cells[(r, c)].formula) >= 0.8 * len(have):
            out.add(c)
    return out


def _typed_finding(sheet: Sheet, changes: dict[Pos, str], formulas: list[int] | None) -> list[Finding]:
    """Der Hinweis zu festen Zahlen in einer Zeile (formulas: ihre Spalten mit Formel) oder, ohne formulas, in einer
    Spalte, die als Werte eingefügt ist."""
    plan = sheet.plan
    fix = einsetzen.formulas(sheet, changes)
    if fix is None:
        return []
    where = sorted(changes)
    first = where[0]
    have = [_number(sheet, p) for p in where]
    want = [_value(sheet, changes[p]) for p in where]
    # Auf Rappen gerundet eingefügt (12249,36 statt 12249,3601) ist dieselbe Zahl, wenn ein ganzes Jahr als Werte
    # eingefügt ist. Eine einzelne eingetippte Zahl, die nur gerundet stimmt (49 statt 48,9974), ist ein Fehler.
    exact = [abs(a - b) <= 1e-9 * max(1.0, abs(b)) for a, b in zip(have, want)]
    close = [_rounded(a, b) for a, b in zip(have, want)]
    same = all(close) if formulas is None else all(exact)
    rounded = formulas is not None and not same and all(close)
    readable = sheet.readable(first, changes[first])
    if formulas is None:
        col = first[1]
        span = f"{address(first)}:{address(where[-1])}"
        text = T(f"In {plan.headers[col]} ({span}) stehen {len(where)} feste Zahlen, wo "
                 f"die übrigen Jahre derselben Zeilen mit Formel rechnen, etwa {sheet.label(first)}: hier wäre das "
                 f"{readable}. Das Jahr ist wohl als Werte eingefügt.",
                 f"In {plan.headers[col]} ({span}) there are {len(where)} fixed numbers where the other years of the "
                 f"same rows calculate with a formula, e.g. {sheet.label(first)}: here it would be {readable}. The year "
                 f"was probably pasted as values.")
        if not same:
            at = close.index(False)
            text += T(f" Die Zahlen stimmen nicht mehr mit den Formeln überein, etwa {sheet.label(where[at])}: "
                      f"{_num(have[at])} statt {_num(want[at])}.",
                      f" The numbers no longer match the formulas, e.g. {sheet.label(where[at])}: "
                      f"{_num(have[at])} instead of {_num(want[at])}.")
        sandwiched = True
    else:
        r = first[0]
        one = len(where) == 1
        span = f"{address((r, formulas[0]))}:{address((r, max(formulas[-1], where[-1][1])))}"
        subject = sheet.label(first) if one else f"{_cells(where)} ({plan.word(r)}, {_years(sheet, [p[1] for p in where])})"
        numbers, wanted = ", ".join(_num(v) for v in have), ", ".join(_num(v) for v in want)
        # Englisch mit „;“ zwischen den Zahlen: Aus „1,5, 2,5“ wird sonst nicht „1.5; 2.5“.
        numbers_en, wanted_en = "; ".join(_num(v) for v in have), "; ".join(_num(v) for v in want)
        text = T(f"{subject} "
                 f"{'ist eine feste Zahl' if one else 'sind feste Zahlen'}, {numbers}. "
                 f"Die übrigen Jahre in {span} rechnen mit einer Formel, hier wäre das {readable}",
                 f"{subject} {'is a fixed number' if one else 'are fixed numbers'}: {numbers_en}. The other years in "
                 f"{span} calculate with a formula; here it would be {readable}")
        text += T(" mit derselben Zahl.", ", with the same number.") if same else T(f" mit {wanted}.", f", with {wanted_en}.")
        if rounded:
            text += T(f" {'Die feste Zahl ist' if one else 'Die festen Zahlen sind'} nur gerundet und "
                      f"{'rechnet' if one else 'rechnen'} nicht mit, wenn sich eine Annahme oder ein Vorjahr ändert.",
                      f" {'The fixed number is' if one else 'The fixed numbers are'} only rounded and "
                      f"{'does' if one else 'do'} not update when an assumption or a prior year changes.")
        # Eine einzelne feste Zahl oder eine zwischen zwei Jahren mit Formel ist ein Versehen. Die letzten Jahre fest
        # (eine Vorgabe der Geschäftsleitung) können gewollt sein.
        sandwiched = one or all(c < formulas[-1] for _, c in where)
    if same:
        text += T(f" {'Die feste Zahl rechnet' if len(where) == 1 else 'Die festen Zahlen rechnen'} aber nicht mit, wenn "
                  "sich eine Annahme oder ein Vorjahr ändert.",
                  f" But {'the fixed number does' if len(where) == 1 else 'the fixed numbers do'} not update when an "
                  "assumption or a prior year changes.")
    checked = _checked(sheet, changes[first])
    tip = T(f"Korrigiert: {checked}.", f"Corrected: {checked}.") if checked else ""
    level = "fehler" if sandwiched and not same else "pruefen"
    return [Finding(f"typed:{address(first)}", text, level, tip, fix=fix)]


# --- Eine Annahme mit $ festgehalten ---


def _fixed_refs(formula: str) -> dict[Pos, bool]:
    """Die Bezüge mit festgehaltener Spalte ($C$6, $C6), die nicht Ecke eines Bereichs sind: Zelle → Zeile auch fest."""
    found: dict[Pos, bool] = {}

    def look(part: str) -> str:
        for match in ANY_REF.finditer(part):
            before, after = part[:match.start()], part[match.end():]
            if before.endswith(":") or after.startswith(":") or before.endswith("!"):
                continue
            if match.group(1):
                found[(int(match.group(4)), column_number(match.group(2)))] = bool(match.group(3))
        return part

    outside_literals(formula, look)
    return found


def _replaced(formula: str, target: Pos, col: int) -> str:
    """Die Formel mit dem Bezug auf target ($C$6) als Bezug auf dieselbe Zeile in Spalte col (D$6 oder D6)."""

    def change(part: str) -> str:
        def one(match: re.Match) -> str:
            before, after = part[:match.start()], part[match.end():]
            if before.endswith(":") or after.startswith(":") or before.endswith("!") or not match.group(1):
                return match.group(0)
            if (int(match.group(4)), column_number(match.group(2))) != target:
                return match.group(0)
            return f"{column_letters(col)}{match.group(3)}{match.group(4)}"

        return ANY_REF.sub(one, part)

    return outside_literals(formula, change)


def _frozen(sheet: Sheet) -> list[Finding]:
    plan = sheet.plan
    found: list[Finding] = []
    for r in _rows(sheet):
        formulas = _formulas(sheet, r)
        if len(formulas) < 3:
            continue
        refs = {c: _fixed_refs(sheet.cells[(r, c)].formula) for c in formulas}
        counts = Counter(p for c in formulas for p in refs[c])
        for target, count in counts.most_common():
            if count < 3 or count < PATTERN_SHARE * len(formulas):
                break
            row, col = target
            if row == r or not plan.in_table(row, col):
                continue
            users = [c for c in formulas if target in refs[c]]
            others = [c for c in users if c != col]
            # Jedes übrige Jahr hat in der Zeile der Annahme einen eigenen Wert, und keine Formel liest ihn.
            own = [c for c in others if _number(sheet, (row, c)) is not None]
            if len(own) < 2 or len(own) < len(others) or any(sheet.dependents.get((row, c)) for c in own):
                continue
            # Auch das Jahr der Annahme selbst ($C$6 → C$6): Sonst rechnete es danach anders als die übrigen Jahre.
            changes = {(r, c): _replaced(sheet.cells[(r, c)].formula, target, c) for c in users}
            if any(not _sensible(sheet, f) for f in changes.values()):
                continue
            fix = einsetzen.formulas(sheet, changes)
            if fix is None:
                continue
            same = all(_number(sheet, (row, c)) == _number(sheet, target) for c in own)
            first = (r, others[0])
            name, assumption = plan.word(r), plan.word(row)
            ref = ("$" + column_letters(col) + ("$" if refs[first[1]][target] else "") + str(row))
            values = f"{address((row, own[0]))}:{address((row, own[-1]))}"
            text = T(f"{_span(sheet, r, users)} ({name}) {'rechnen' if len(users) > 1 else 'rechnet'} in jedem Jahr mit "
                     f"{assumption} {plan.headers[col]} ({ref}). Die Werte der übrigen Jahre "
                     f"({values}) liest keine Formel.",
                     f"{_span(sheet, r, users)} ({name}) {'use' if len(users) > 1 else 'uses'} {assumption} "
                     f"{plan.headers[col]} ({ref}) in every year. No formula reads the values of the other years "
                     f"({values}).")
            if same:
                text += T(" Heute sind sie gleich, eine Änderung dort rechnet aber nicht mit.",
                          " Today they are the same, but a change there is not picked up.")
            checked = _checked(sheet, changes[first])
            tip = T(f"Korrigiert: jedes Jahr mit seinem eigenen Wert, {address(first)} {checked}.",
                    f"Corrected: each year with its own value, {address(first)} {checked}.") if checked else ""
            found.append(Finding(f"frozen:{address(first)}", text, "pruefen" if same else "fehler", tip, fix=fix))
            break
    return found


# --- Eine SUMME, die in jedem Jahr eine Zeile auslässt ---


def _near(left_out: list[int], top: int, bottom: int, row: int) -> list[int]:
    """Die ausgelassenen Zeilen, die zu den Teilen gehören: zwischen den Teilen und der Summe (der letzte Posten vor
    dem Operativen Cashflow) oder direkt vor dem ersten Teil. Was weiter weg steht, ist eine andere Größe."""
    out = set(left_out)
    if row > bottom:
        keep = {r for r in out if bottom < r < row}
        r = top - 1
        while r in out:
            keep.add(r)
            r -= 1
    else:
        keep = {r for r in out if row < r < top}
        r = bottom + 1
        while r in out:
            keep.add(r)
            r += 1
    return sorted(keep)


def _gaps(sheet: Sheet) -> list[Finding]:
    from .task import sum_gaps

    plan = sheet.plan
    rows = sorted(r for r in plan.names if plan.header_row < r <= plan.last_row)
    by_row: dict[int, list[tuple[int, int, int, list[int]]]] = {}
    for row, col, top, bottom, left_out in sum_gaps(sheet.cells, plan, rows):
        # Eine Zeile, die eine andere Formel liest (die Periode im Diskontfaktor), fließt dort ins Modell ein: Sie ist
        # kein vergessener Teil der Summe. Ein vergessener Posten hängt meist lose da, keine Formel liest ihn.
        left_out = _near([r for r in left_out if not sheet.dependents.get((r, col))], top, bottom, row)
        if left_out:
            by_row.setdefault(row, []).append((col, top, bottom, left_out))
    found: list[Finding] = []
    for row, items in sorted(by_row.items()):
        items.sort()
        left = sorted({r for *_, out in items for r in out})
        changes: dict[Pos, str] = {}
        for col, top, bottom, out in items:
            start, end = min(top, *out), max(bottom, *out)
            # Nur wenn der längere Bereich nichts anderes dazunimmt als die ausgelassenen Zeilen.
            if any(_number(sheet, (r, col)) is not None for r in range(start, end + 1)
                   if not top <= r <= bottom and r not in out):
                changes = {}
                break
            fix = einsetzen.stretched(sheet, [(row, col)], "SUM", (top, col, bottom, col), (start, end))
            if fix is None:
                changes = {}
                break
            changes[(row, col)] = fix["cells"][address((row, col))][1:]
        fix = einsetzen.formulas(sheet, changes) if changes else None
        first = (row, items[0][0])
        names = _listed([plan.word(r) or T(f"Zeile {r}", f"row {r}") for r in left])
        many = len(items) > 1
        cols = [c for c, *_ in items]
        spread = _span(sheet, row, cols) if cols == _columns(sheet)[_columns(sheet).index(cols[0]):][:len(cols)] \
            else _cells([(row, c) for c in cols])
        where = (T(f"Die Summen in {spread} ({plan.word(row)})", f"The sums in {spread} ({plan.word(row)})") if many
                 else T(f"Die Summe in {sheet.label(first)}", f"The sum in {sheet.label(first)}"))
        text = T(f"{where} {'zählen' if many else 'zählt'} {names} nicht mit: {sheet.readable(first)}",
                 f"{where} {'do' if many else 'does'} not include {names}: {sheet.readable(first)}")
        text += T(f" statt {sheet.readable(first, changes[first])}.", f" instead of {sheet.readable(first, changes[first])}.") \
            if fix else "."
        if fix:
            # Was fehlt und was die Summe darum zeigt, wie beim Hinweis zum Muster („So fehlt G34 … mit -61,77.“).
            if not many:
                col, *_, out = items[0]
                text = _moved_text(sheet, text, [(r, col) for r in out][:MAX_LISTED_MOVED], len(out), [], 0)
            text += _shows(sheet, first, changes[first], "Zeile")
        tip = (T(f"Korrigiert: {_checked(sheet, changes[first])}.", f"Corrected: {_checked(sheet, changes[first])}.")
               if fix and _checked(sheet, changes[first])
               else T("Prüfe, ob die Zeile zum Abschnitt gehört.", "Check whether the row belongs to the section."))
        found.append(Finding(f"gap:{address(first)}", text, "fehler", tip, fix=fix))
    return found


# --- Bezüge von außerhalb auf die Jahre: ein Bereich, der ein Jahr auslässt, und das falsche Jahr ---


def _periods(sheet: Sheet) -> list[int]:
    """Die Spalten der Jahre (oder Monate): die Art, die am häufigsten ist („Jahr 1“ … „Jahr 5“, nicht „Basisjahr“)."""
    memo = sheet.memo.setdefault("zeilenmodell", {})
    if "periods" not in memo:
        plan = sheet.plan
        columns = _columns(sheet)
        kinds = Counter(_kind(plan.headers[c]) for c in columns)
        common = kinds.most_common(1)[0][0] if kinds else None
        memo["periods"] = [c for c in columns if _kind(plan.headers[c]) == common]
    return memo["periods"]


def _lines(sheet: Sheet) -> set[int]:
    """Die Zeilen des Modells, die über die Jahre gehen (mindestens zwei Jahre gefüllt). Eine Kennzahl darunter in
    der Spalte des ersten Jahres (der Kapitalwert in B23) ist keine Zeile, sondern rechnet über die Zeilen."""
    memo = sheet.memo.setdefault("zeilenmodell", {})
    if "lines" not in memo:
        periods = _periods(sheet)
        memo["lines"] = {r for r in _rows(sheet)
                         if sum(1 for c in periods if (r, c) in sheet.nodes or _number(sheet, (r, c)) is not None) >= 2}
    return memo["lines"]


def _shape(sheet: Sheet, pos: Pos) -> str | None:
    """Wie eine Zelle der Zeile rechnet (die Formel relativ zur Zelle), „#“ für eine feste Zahl, sonst None."""
    cell = sheet.cells.get(pos)
    if cell is None:
        return None
    if cell.formula:
        return None if pos in sheet.unknown else pattern(cell.formula, pos, "Zeile")
    return "#" if _number(sheet, pos) is not None else None


_AREA_TEXT = re.compile(r"(?<![A-Za-z0-9_.!$\\])(\$?)([A-Z]{1,3})(\$?)(\d{1,7}):(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(!])")


def _widened(formula: str, area: tuple[int, int, int, int], new: tuple[int, int, int, int]) -> str:
    """Die Formel mit dem Bereich area (C32:F32) als new (C32:G32), die $ bleiben, wie sie sind."""

    def change(part: str) -> str:
        def one(match: re.Match) -> str:
            d1, a, d2, b, d3, c, d4, d = match.groups()
            if (int(b), column_number(a), int(d), column_number(c)) != area:
                return match.group(0)
            r1, c1, r2, c2 = new
            return f"{d1}{column_letters(c1)}{d2}{r1}:{d3}{column_letters(c2)}{d4}{r2}"

        return _AREA_TEXT.sub(one, part)

    return outside_literals(formula, change)


# Funktionen, die über alle Jahre einer Zeile rechnen sollen: die Summe der Barwerte, der Mittelwert, der Barwert.
_OVER_YEARS = {"SUM", "AVERAGE", "SUMPRODUCT", "NPV", "PRODUCT", "MIN", "MAX", "COUNT", "MEDIAN", "IRR"}
_VERB = {"SUM": "zählt nur", "AVERAGE": "mittelt nur über", "NPV": "zinst nur", "IRR": "rechnet nur mit"}
_VERB_EN = {"SUM": "only adds up", "AVERAGE": "only averages over", "NPV": "only discounts", "IRR": "only uses"}


def _left_words(sheet: Sheet, pos: Pos) -> str:
    from .analyse import _left_label

    return _left_label(sheet.cells, *pos) or ""


def _short_lines(sheet: Sheet) -> list[Finding]:
    """Eine Formel außerhalb des Modells (die Bewertung unter der Projektion) rechnet über eine Zeile, lässt aber ihr
    letztes oder erstes Jahr aus: =SUMME(C32:F32) für die Summe der Barwerte, wo die Barwerte bis G32 gehen. Ein
    Bereich, der gewollt nur einen Teil nimmt, sagt das im Namen („Summe Jahr 1-4“, „Umsatz 2026“) oder nimmt genau
    die Spalten eines Jahres (Q1 2026 bis Q4 2026)."""
    periods = _periods(sheet)
    if len(periods) < 3:
        return []
    rows = _lines(sheet)
    found: list[Finding] = []
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown or (pos[0] in rows and pos[1] in periods):
            continue
        for call in formel.calls(node):
            if call.value not in _OVER_YEARS:
                continue
            for arg in call.args:
                if arg.kind != "range":
                    continue
                item = _short_line(sheet, pos, call.value, arg.value, rows, periods)
                if item is not None:
                    found.append(item)
                    break
            else:
                continue
            break
    return found


def _short_line(sheet: Sheet, pos: Pos, function: str, area: tuple[int, int, int, int], rows: set[int],
                periods: list[int]) -> Finding | None:
    plan = sheet.plan
    r, c1, r2, c2 = area
    if r != r2 or r not in rows or c2 <= c1 or r == pos[0]:
        return None
    inside = list(range(c1, c2 + 1))
    if any(c not in periods for c in inside):
        return None
    shapes = {c: _shape(sheet, (r, c)) for c in periods}
    have = {shapes[c] for c in inside}
    if None in have:
        return None
    # Was außen anschließt und genauso rechnet, gehört zur Zeile: bis zum ersten Jahr, das anders rechnet.
    after, before = [], []
    for c in periods[periods.index(c2) + 1:]:
        if shapes[c] not in have:
            break
        after.append(c)
    for c in reversed(periods[:periods.index(c1)]):
        if shapes[c] not in have:
            break
        before.append(c)
    before.reverse()
    out = before + after
    if not out:
        return None
    # Liest die Formel die fehlenden Jahre anders (=SUMME(C32:F32)+G32*0,5), ist das gewollt.
    if any((r, c) in sheet.precedents.get(pos, ()) and not c1 <= c <= c2
           for c in out if isinstance(sheet.precedents.get(pos), set)):
        return None
    heads = [plan.headers[c] for c in inside]
    label = _left_words(sheet, pos).casefold()
    # Der Name sagt, welche Jahre gemeint sind („Summe Jahr 1-4“, „Umsatz 2026“).
    digits = set(re.findall(r"\d+", label))
    if digits & set(re.findall(r"\d+", heads[0] + " " + heads[-1])):
        return None
    # Genau die Spalten eines Jahres (Q1 2026 bis Q4 2026) neben dem nächsten.
    years = [set(re.findall(r"(?<!\d)(?:19|20)\d\d(?!\d)", h)) for h in heads]
    common = set.intersection(*years) if years else set()
    if common and not any(common & set(re.findall(r"(?<!\d)(?:19|20)\d\d(?!\d)", plan.headers[c])) for c in out):
        return None
    new = (r, min(c1, *out), r, max(c2, *out))
    formula = _widened(sheet.cells[pos].formula, area, new)
    if formula == sheet.cells[pos].formula or not _sensible(sheet, formula) \
            or not isinstance(_value(sheet, formula), float):
        return None
    fix = einsetzen.formulas(sheet, {pos: formula})
    if fix is None:
        return None
    name = plan.word(r) or T(f"Zeile {r}", f"row {r}")
    span = f"{address((r, c1))}:{address((r, c2))}"
    verb = T(_VERB.get(function, "nimmt nur"), _VERB_EN.get(function, "only takes"))
    text = T(f"{sheet.label(pos)} {verb} {span} ({name} {heads[0]} bis {heads[-1]}), ",
             f"{sheet.label(pos)} {verb} {span} ({name} {heads[0]} to {heads[-1]}), ")
    if after and before:
        text += T(f"die Zeile geht aber von {plan.headers[periods[0]]} bis {plan.headers[after[-1]]}.",
                  f"but the row runs from {plan.headers[periods[0]]} to {plan.headers[after[-1]]}.")
    elif after:
        text += T(f"die Zeile geht aber bis {plan.headers[after[-1]]}.", f"but the row runs to {plan.headers[after[-1]]}.")
    else:
        text += T(f"die Zeile fängt aber schon mit {plan.headers[before[0]]} an.",
                  f"but the row already starts with {plan.headers[before[0]]}.")
    text = _moved_text(sheet, text, [(r, c) for c in out][:MAX_LISTED_MOVED], len(out), [], 0)
    text += _shows(sheet, pos, formula)
    checked = _checked(sheet, formula)
    tip = T(f"Korrigiert: {checked}.", f"Corrected: {checked}.") if checked else ""
    return Finding(f"line:{address(pos)}", text, "fehler", tip, fix=fix)


def _to_column(formula: str, target: Pos, col: int) -> str:
    """Die Formel mit jedem einzelnen Bezug auf target (F31, $F$31) als Bezug auf dieselbe Zeile in Spalte col."""

    def change(part: str) -> str:
        def one(match: re.Match) -> str:
            before, after = part[:match.start()], part[match.end():]
            if before.endswith(":") or after.startswith(":") or before.endswith("!") or not match.group(2):
                return match.group(0)
            if (int(match.group(4)), column_number(match.group(2))) != target:
                return match.group(0)
            return f"{match.group(1)}{column_letters(col)}{match.group(3)}{match.group(4)}"

        return ANY_REF.sub(one, part)

    return outside_literals(formula, change)


def _direct(sheet: Sheet, pos: Pos, grid: set[Pos]) -> set[Pos]:
    """Die einzelnen Zellen der Jahre, die eine Formel selbst liest (keine Bereiche)."""
    node = sheet.nodes.get(pos)
    if node is None:
        return set()
    singles, _, _ = formel.references(node, sheet.name)
    return {p for p in singles if p in grid}


def _mixed_years(sheet: Sheet) -> list[Finding]:
    """Eine Rechnung außerhalb des Modells nimmt ein Jahr zu früh: Der Terminalwert stammt aus Jahr 5 (über
    =G30*(1+g)), abgezinst wird er aber mit der Periode von Jahr 4 (=B37/(1+B15)^F31). In einer Rechnung, die sonst
    nur das letzte Jahr liest, ist ein einzelner Bezug auf das Jahr davor fast immer ein Versehen."""
    periods = _periods(sheet)
    if len(periods) < 3:
        return []
    last, before_last = periods[-1], periods[-2]
    lines = _lines(sheet)
    grid = {(r, c) for r in lines for c in periods}
    outside = [p for p, n in sheet.nodes.items() if n is not None and p not in sheet.unknown and p not in grid]
    if not outside:
        return []
    direct = {p: _direct(sheet, p, grid) for p in outside}
    reads: dict[Pos, frozenset[int]] = {}

    def columns(pos: Pos, seen: frozenset = frozenset()) -> frozenset[int]:
        """Die Jahre, die eine Formel außerhalb selbst oder über andere Formeln außerhalb liest."""
        if pos in reads:
            return reads[pos]
        if pos in seen:
            return frozenset()
        found = {c for _, c in direct.get(pos, ())}
        before = sheet.precedents.get(pos, ())
        for p in (before if isinstance(before, set) else ()):
            if p in direct:
                found |= columns(p, seen | {pos})
        reads[pos] = frozenset(found)
        return reads[pos]

    found: list[Finding] = []
    done: set[Pos] = set()
    for pos in sorted(outside):
        own = {c for _, c in direct[pos]}
        before = sheet.precedents.get(pos, ())
        if len(own) != 1 or not isinstance(before, set):
            continue
        rest = set()
        for p in before:
            if p in direct:
                rest |= columns(p)
        if len(rest) != 1 or own == rest or own | rest != {last, before_last}:
            continue
        if own == {before_last}:
            wrong, source = pos, next((p for p in before if p in direct and columns(p) == {last}), None)
        else:
            # Das Jahr davor kommt über eine frühere Formel: die, die es selbst liest.
            wrong, source = _origin(sheet, pos, direct, columns, before_last), pos
        if wrong is None or source is None or wrong in done:
            continue
        item = _mixed_finding(sheet, wrong, source, direct, columns, before_last, last)
        if item is not None:
            done.add(wrong)
            found.append(item)
    return found


def _origin(sheet: Sheet, pos: Pos, direct: dict, columns, year: int) -> Pos | None:
    """Die Formel in der Kette vor pos, die das Jahr year selbst liest (und sonst keines)."""
    todo, seen = [pos], set()
    while todo:
        here = todo.pop()
        if here in seen:
            continue
        seen.add(here)
        before = sheet.precedents.get(here, ())
        for p in sorted(before if isinstance(before, set) else ()):
            if p not in direct:
                continue
            if {c for _, c in direct[p]} == {year} and columns(p) == {year}:
                return p
            todo.append(p)
    return None


def _path(sheet: Sheet, start: Pos, direct: dict, year: int) -> list[Pos]:
    """Der Weg von start über Formeln außerhalb bis zur Zelle des Jahres year, die sie liest (B37 → B36 → G30)."""
    todo, seen = [[start]], set()
    while todo:
        path = todo.pop(0)
        here = path[-1]
        hit = sorted(p for p in direct.get(here, ()) if p[1] == year)
        if hit:
            return path + [hit[0]]
        if here in seen:
            continue
        seen.add(here)
        before = sheet.precedents.get(here, ())
        todo += [path + [p] for p in sorted(before if isinstance(before, set) else ()) if p in direct]
    return [start]


def _mixed_finding(sheet: Sheet, wrong: Pos, source: Pos, direct: dict, columns, year: int, last: int) -> Finding | None:
    plan = sheet.plan
    odd = sorted(p for p in direct[wrong] if p[1] == year)
    formula = sheet.cells[wrong].formula
    for p in odd:
        formula = _to_column(formula, p, last)
    if formula == sheet.cells[wrong].formula or not _sensible(sheet, formula) \
            or not isinstance(_value(sheet, formula), float):
        return None
    fix = einsetzen.formulas(sheet, {wrong: formula})
    if fix is None:
        return None
    named = _cells_labelled(sheet, odd)
    if source != wrong:
        # Der Rest der Rechnung kommt über source: B37 rechnet über B36 mit G30.
        path = _path(sheet, source, direct, last)
        text = T(f"{sheet.label(wrong)} rechnet mit {named}. Der Rest der Rechnung stammt aber aus "
                 f"{plan.headers[last]}: {sheet.label(path[0])} rechnet",
                 f"{sheet.label(wrong)} uses {named}. But the rest of the calculation comes from "
                 f"{plan.headers[last]}: {sheet.label(path[0])} calculates")
        if len(path) > 2:
            text += T(" über ", " via ") + _cells(path[1:-1])
        text += T(f" mit {sheet.label(path[-1])}.", f" with {sheet.label(path[-1])}.") if len(path) > 1 else \
            T(" damit.", " with it.")
    else:
        later = next((p for p in sorted(sheet.dependents.get(wrong, ())) if p in direct and columns(p) >= {last}), None)
        uses = sorted(p for p in direct.get(later, ()) if p[1] == last) if later else []
        text = T(f"{sheet.label(wrong)} rechnet mit {named}.", f"{sheet.label(wrong)} uses {named}.")
        if later and uses:
            text += T(f" Die Rechnung, in die es einfließt, nimmt sonst {plan.headers[last]}: {sheet.label(later)} "
                      f"rechnet mit {_cells_labelled(sheet, uses)}.",
                      f" The calculation it feeds into otherwise takes {plan.headers[last]}: {sheet.label(later)} "
                      f"uses {_cells_labelled(sheet, uses)}.")
    right = _cells_labelled(sheet, [(p[0], last) for p in odd])
    text += T(f" Richtig wäre {right}.", f" Correct would be {right}.")
    text += _shows(sheet, wrong, formula)
    checked = _checked(sheet, formula)
    tip = T(f"Korrigiert: {checked}.", f"Corrected: {checked}.") if checked else ""
    return Finding(f"year:{address(wrong)}", text, "fehler", tip, fix=fix)


def _cells_labelled(sheet: Sheet, cells: list[Pos]) -> str:
    return _listed([sheet.label(p) for p in cells])


# --- Eine Annahme je Jahr aus einer Liste, und ein Bezug ohne $, der beim Ziehen mitrutscht ---


def _refs(formula: str) -> list[tuple[str, int, str, int]]:
    """Die einzelnen Bezüge einer Formel der Reihe nach ($, Spalte, $, Zeile), ohne Bereiche und andere Blätter."""
    found: list[tuple[str, int, str, int]] = []

    def look(part: str) -> str:
        for match in ANY_REF.finditer(part):
            before, after = part[:match.start()], part[match.end():]
            if before.endswith(":") or after.startswith(":") or before.endswith("!"):
                continue
            found.append((match.group(1), column_number(match.group(2)), match.group(3), int(match.group(4))))
        return part

    outside_literals(formula, look)
    return found


def _skeleton(formula: str) -> str:
    return outside_literals(formula.replace("$", ""), lambda part: ANY_REF.sub("R", part))


def _walks(sheet: Sheet) -> list[Finding]:
    """Jedes Jahr liest seine Annahme aus einer Liste untereinander (C22 =$B$5, D22 =$B$6 … für Umsatzwachstum
    Jahr 1 bis Jahr 5). Liest ein Jahr dieselbe Zelle wie ein anderes oder springt eine über, rechnet es mit der
    Annahme eines anderen Jahres, und die eigene liest keine Formel."""
    periods = _periods(sheet)
    found: list[Finding] = []
    for r in _rows(sheet):
        cols = [c for c in _formulas(sheet, r) if c in periods]
        if len(cols) < 4:
            continue
        formulas = {c: sheet.cells[(r, c)].formula for c in cols}
        if len({_skeleton(f) for f in formulas.values()}) != 1:
            continue
        refs = {c: _refs(formulas[c]) for c in cols}
        if len({len(v) for v in refs.values()}) != 1:
            continue
        for i in range(len(refs[cols[0]])):
            targets = {c: refs[c][i] for c in cols}
            if len({t[1] for t in targets.values()}) != 1:
                continue
            index = {c: periods.index(c) for c in cols}
            offsets = Counter(t[3] - index[c] for c, t in targets.items())
            start, fitting = offsets.most_common(1)[0]
            if fitting < len(cols) - 1 or fitting == len(cols) or len(offsets) != 2:
                continue
            odd = next(c for c, t in targets.items() if t[3] - index[c] != start)
            have = targets[odd]
            want = (start + index[odd], have[1])
            if _number(sheet, want) is None or sheet.dependents.get(want):
                continue
            formula = _to_cell(formulas[odd], (have[3], have[1]), want)
            if not _sensible(sheet, formula) or not isinstance(_value(sheet, formula), float):
                continue
            fix = einsetzen.formulas(sheet, {(r, odd): formula})
            if fix is None:
                continue
            twice = [c for c, t in targets.items() if c != odd and (t[3], t[1]) == (have[3], have[1])]
            first, last = (start + index[cols[0]], have[1]), (start + index[cols[-1]], have[1])
            text = T(f"{sheet.label((r, odd))} liest {sheet.label((have[3], have[1]))}"
                     + (f", wie schon {_cells([(r, c) for c in twice])}" if twice else "") + ". "
                     f"Die übrigen Jahre gehen die Liste {address(first)}:{address(last)} der Reihe nach durch, hier wäre "
                     f"das {sheet.label(want)}. {address(want)} liest so keine Formel.",
                     f"{sheet.label((r, odd))} reads {sheet.label((have[3], have[1]))}"
                     + (f", as {_cells([(r, c) for c in twice])} already {'does' if len(twice) == 1 else 'do'}"
                        if twice else "") + ". "
                     f"The other years go through the list {address(first)}:{address(last)} in order; here it would "
                     f"be {sheet.label(want)}. This way no formula reads {address(want)}.")
            text += _shows(sheet, (r, odd), formula, "Zeile")
            checked = _checked(sheet, formula)
            found.append(Finding(f"walk:{address((r, odd))}", text, "fehler",
                                 T(f"Korrigiert: {checked}.", f"Corrected: {checked}.") if checked else "", fix=fix))
            break
    return found


def _to_cell(formula: str, target: Pos, new: Pos) -> str:
    """Die Formel mit jedem einzelnen Bezug auf target als Bezug auf new, die $ bleiben."""

    def change(part: str) -> str:
        def one(match: re.Match) -> str:
            before, after = part[:match.start()], part[match.end():]
            if before.endswith(":") or after.startswith(":") or before.endswith("!"):
                return match.group(0)
            if (int(match.group(4)), column_number(match.group(2))) != target:
                return match.group(0)
            return f"{match.group(1)}{column_letters(new[1])}{match.group(3)}{new[0]}"

        return ANY_REF.sub(one, part)

    return outside_literals(formula, change)


def _drift(sheet: Sheet) -> list[Finding]:
    """Ein Bezug ohne $ auf eine einzelne Annahme (=C30/(1+B15)^C31): Nach rechts gezogen liest D32 C15, E32 D15 …,
    und die sind leer. Nur das erste Jahr rechnet so mit dem WACC, die übrigen mit 0."""
    plan = sheet.plan
    periods = _periods(sheet)
    found: list[Finding] = []
    memo = sheet.memo.setdefault("zeilenmodell", {}).setdefault("drift", set())
    for r in _rows(sheet):
        cols = [c for c in _formulas(sheet, r) if c in periods]
        if len(cols) < 3:
            continue
        shapes = {pattern(sheet.cells[(r, c)].formula, (r, c), "Zeile") for c in cols}
        if len(shapes) != 1:
            continue
        refs = {c: _refs(sheet.cells[(r, c)].formula) for c in cols}
        for i, (dc, col, dr, row) in enumerate(refs[cols[0]]):
            if dc or plan.in_table(row, col) and row in plan.names and col in periods:
                continue
            targets = {c: (refs[c][i][3], refs[c][i][1]) for c in cols}
            own = targets[cols[0]]
            rest = [targets[c] for c in cols[1:]]
            if _number(sheet, own) is None or any(sheet.shown.get(t) is not None for t in rest):
                continue
            absolute = f"${column_letters(own[1])}${own[0]}"
            changes = {}
            for c in cols:
                t = targets[c]
                changes[(r, c)] = _to_absolute(sheet.cells[(r, c)].formula, t, absolute)
            if any(not _sensible(sheet, f) for f in changes.values()):
                continue
            fix = einsetzen.formulas(sheet, changes)
            if fix is None:
                continue
            memo.update(changes)
            first, second = (r, cols[0]), (r, cols[1])
            meant = _left_words(sheet, own) or address(own)
            text = T(f"{_span(sheet, r, cols)} ({plan.word(r)}) {'lesen' if len(cols) > 1 else 'liest'} "
                     f"{sheet.label(own)} ohne $: Nach rechts gezogen rutscht der Bezug mit, {address(second)} liest "
                     f"{address(targets[cols[1]])}, {address((r, cols[-1]))} {address(targets[cols[-1]])}, und die sind "
                     f"leer (Excel nimmt dann 0). Nur {address(first)} rechnet so mit {meant}.",
                     f"{_span(sheet, r, cols)} ({plan.word(r)}) {'read' if len(cols) > 1 else 'reads'} "
                     f"{sheet.label(own)} without $: filled to the right, the reference shifts, {address(second)} reads "
                     f"{address(targets[cols[1]])}, {address((r, cols[-1]))} {address(targets[cols[-1]])}, and those "
                     f"are empty (Excel then uses 0). Only {address(first)} actually calculates with {meant}.")
            text += _shows(sheet, second, changes[second], "Zeile")
            checked = _checked(sheet, changes[second])
            found.append(Finding(f"drift:{address(first)}", text, "fehler",
                                 T(f"Korrigiert: jedes Jahr mit {absolute}, {address(second)} {checked}.",
                                   f"Corrected: each year with {absolute}, {address(second)} {checked}.") if checked else "",
                                 fix=fix))
            break
    return found


def _to_absolute(formula: str, target: Pos, absolute: str) -> str:
    def change(part: str) -> str:
        def one(match: re.Match) -> str:
            before, after = part[:match.start()], part[match.end():]
            if before.endswith(":") or after.startswith(":") or before.endswith("!"):
                return match.group(0)
            if (int(match.group(4)), column_number(match.group(2))) != target:
                return match.group(0)
            return absolute

        return ANY_REF.sub(one, part)

    return outside_literals(formula, change)


# --- NBW über Jahr 0 ---


def _call_text(formula: str, name: str) -> tuple[int, int] | None:
    """Anfang und Ende (hinter der Klammer) des ersten Aufrufs von name in der Formel, außerhalb von Texten."""
    plain = re.sub(r'"(?:[^"]|"")*"', lambda m: " " * len(m.group(0)), formula)
    match = re.search(rf"(?<![A-Za-z0-9_.]){name}\(", plain, re.IGNORECASE)
    if not match:
        return None
    depth = 0
    for i in range(match.end() - 1, len(plain)):
        if plain[i] == "(":
            depth += 1
        elif plain[i] == ")":
            depth -= 1
            if depth == 0:
                return match.start(), i + 1
    return None


def _first_period(sheet: Sheet, col: int) -> bool:
    """Ist die Spalte das Jahr 0 (heute)? Am Kopf („Jahr 0“, „0“) oder an einer Zeile, die die Perioden zählt."""
    plan = sheet.plan
    head = plan.headers.get(col, "")
    if re.search(r"(?<![\d.,])0$", head.strip()):
        return True
    for r in _lines(sheet):
        here, after = _number(sheet, (r, col)), _number(sheet, (r, col + 1))
        if here == 0 and after == 1 and (r, col) not in sheet.nodes:
            return True
    return False


def _npv_today(sheet: Sheet) -> list[Finding]:
    """NBW(Zins; B18:G18) über Jahr 0: NBW zinst schon den ersten Wert um ein Jahr ab. Die Investition heute gehört
    vor das NBW (=B18+NBW(Zins; C18:G18)). Der bekannteste Fehler in Investitionsrechnungen."""
    periods = _periods(sheet)
    if len(periods) < 3:
        return []
    rows = _lines(sheet)
    found: list[Finding] = []
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown or (pos[0] in rows and pos[1] in periods):
            continue
        calls = [c for c in formel.calls(node) if c.value == "NPV" and len(c.args) == 2 and c.args[1].kind == "range"]
        if len(calls) != 1:
            continue
        r, c1, r2, c2 = calls[0].args[1].value
        if r != r2 or r not in rows or c2 - c1 < 2 or c1 not in periods or not _first_period(sheet, c1):
            continue
        formula = sheet.cells[pos].formula
        span = _call_text(formula, "NPV")
        if span is None:
            continue
        # Liest die Formel Jahr 0 noch einmal selbst (=NBW(…;B18:G18)-B18*…), ist das schon bedacht.
        singles, _, _ = formel.references(node, sheet.name)
        if (r, c1) in singles:
            continue
        call = _widened(formula[span[0]:span[1]], (r, c1, r, c2), (r, c1 + 1, r, c2))
        first = address((r, c1))
        whole = span == (0, len(formula))
        new = f"{first}+{call}" if whole else formula[:span[0]] + f"({first}+{call})" + formula[span[1]:]
        if not _sensible(sheet, new) or not isinstance(_value(sheet, new), float):
            continue
        fix = einsetzen.formulas(sheet, {pos: new})
        if fix is None:
            continue
        from .bereiche import _german_formula

        text = T(f"{sheet.label(pos)} zinst mit NBW auch {sheet.label((r, c1))} ab. NBW teilt schon den ersten Wert "
                 f"durch (1 + Zins), {sheet.plan.headers[c1]} fällt aber heute an und gehört vor das NBW: "
                 f"={_german_formula(new)}.",
                 f"{sheet.label(pos)} also discounts {sheet.label((r, c1))} with NPV. NPV already divides the first "
                 f"value by (1 + rate), but {sheet.plan.headers[c1]} falls due today and belongs in front of the NPV: "
                 f"={_german_formula(new)}.")
        text += _shows(sheet, pos, new)
        checked = _checked(sheet, new)
        found.append(Finding(f"npv:{address(pos)}", text, "fehler",
                             T(f"Korrigiert: {checked}.", f"Corrected: {checked}.") if checked else "", fix=fix))
    return found


# --- Eine feste Zahl in einer Spalte, die sonst rechnet (Tilgungsplan, Abschreibungsplan) ---

# Bis zu so vielen Zeilen: ein Plan, in dem jede Zeile eine Periode ist. Größere Listen sind Fälle, dort sagt task.py
# nur, dass eine Zahl steht.
DOWN_ROWS = 120


def _last_above_total(sheet: Sheet, r: int, col: int) -> bool:
    """Steht unter (r, col) nichts mehr, oder die Summe der Spalte, die auch r und die Zeile darüber zählt?"""
    below = (r + 1, col)
    if r == sheet.plan.last_row and below not in sheet.nodes and not (sheet.cells.get(below) and sheet.cells[below].value.strip()):
        return True
    reads = sheet.precedents.get(below, ())
    return below in sheet.nodes and (r, col) in reads and (r - 1, col) in reads


def typed_down(sheet: Sheet) -> list[Finding]:
    """Eine einzelne feste Zahl zwischen Zeilen, die mit derselben Formel rechnen (die Rate in Jahr 5 eines
    Tilgungsplans eingetippt): Sie rechnet nicht mit, wenn sich Zins oder Betrag ändern. Mit der Formel der
    Nachbarzeilen zum Einsetzen."""
    plan = sheet.plan
    if plan.across or plan.header_row is None or plan.last_row - plan.header_row > DOWN_ROWS:
        return []
    found: list[Finding] = []
    for col in sorted(plan.headers):
        if col == plan.name_column:
            continue
        rows = [r for r in range(plan.header_row + 1, plan.last_row + 1)
                if (cell := sheet.cells.get((r, col))) is not None and cell.formula and (r, col) in sheet.nodes
                and (r, col) not in sheet.unknown]
        if len(rows) < 3:
            continue
        typed = [r for r in range(rows[0] + 1, plan.last_row + 1) if (r, col) in sheet.cells
                 and not sheet.cells[(r, col)].formula and _number(sheet, (r, col)) is not None]
        # Sonst die erste Zeile über den Formeln: Jahr 1 eingetippt, ab Jahr 2 rechnet die Spalte. Ein Startwert
        # (km-Stand am Anfang) bleibt dabei ein Startwert: Die Formel darunter zeigte dort auf die Kopfzeile.
        start = rows[0] - 1
        if not typed and start > plan.header_row and (start, col) in sheet.cells and not sheet.cells[(start, col)].formula \
                and _number(sheet, (start, col)) is not None:
            typed = [start]
        if len(typed) != 1:
            continue
        r = typed[0]
        shapes = {q: pattern(sheet.cells[(q, col)].formula, (q, col), "Spalte") for q in rows}
        near = [q for q in (r - 1, r + 1) if q in shapes]
        source = r - 1
        if len(near) == 2 and shapes[near[0]] == shapes[near[1]]:
            common = shapes[near[0]]
        elif r - 2 in shapes and r - 1 in shapes and shapes[r - 2] == shapes[r - 1] and _last_above_total(sheet, r, col):
            # Die letzte Zeile vor der Summe (oder am Ende der Tabelle) eingetippt: Darüber rechnet dieselbe Formel.
            common = shapes[r - 1]
        elif r + 2 in shapes and shapes[r + 1] == shapes[r + 2] and (r - 1 not in shapes or shapes[r - 1] != shapes[r + 1]):
            # Die erste Zeile eingetippt (oder die erste nach Jahr 0): Darunter rechnet dieselbe Formel. Ein
            # Startwert, der so nicht rechnen kann (die Formel zeigte auf die Kopfzeile), fällt unten bei _sensible
            # heraus.
            common, source = shapes[r + 1], r + 1
        else:
            continue
        if sum(1 for v in shapes.values() if v == common) < max(3, PATTERN_SHARE * len(rows)):
            continue
        expected = shift(sheet.cells[(source, col)].formula, r - source, 0)
        if not _sensible(sheet, expected):
            continue
        want, have = _value(sheet, expected), _number(sheet, (r, col))
        if not isinstance(want, float):
            continue
        fix = einsetzen.formulas(sheet, {(r, col): expected})
        if fix is None:
            continue
        pos = (r, col)
        same = abs(want - have) <= 1e-9 * max(1.0, abs(want))
        rounded = _rounded(have, want)
        # Ein Datum (Ende eines Projektschritts) steht als Datum da, nicht als Zahl.
        like = sheet.cells[pos].value.strip()
        dated = einsetzen._DATE_TEXT.fullmatch(like) is not None
        shown, readable = like if dated else _num(have), sheet.readable(pos, expected)
        result = einsetzen._date_text(want, False) if dated else _num(want)
        text = T(f"{sheet.label(pos)} ist {'ein festes Datum' if dated else 'eine feste Zahl'}, "
                 f"{shown}. Die Zeilen {'darunter' if source > r else 'darüber und darunter'} rechnen mit "
                 f"einer Formel, hier wäre das {readable}"
                 + (", mit demselben Ergebnis." if same else
                    f" mit {result}."),
                 f"{sheet.label(pos)} is {'a fixed date' if dated else 'a fixed number'}, {shown}. The rows "
                 f"{'below' if source > r else 'above and below'} calculate with a formula; here it would be {readable}"
                 + (", with the same result." if same else f", with {result}."))
        text += (T(" Sie rechnet aber nicht mit, wenn sich etwas davor ändert.",
                   " But it does not update when something before it changes.") if same else
                 T(" Die feste Zahl ist nur gerundet und rechnet nicht mit, wenn sich etwas davor ändert.",
                   " The fixed number is only rounded and does not update when something before it changes.") if rounded
                 else T(f" {'Das feste Datum' if dated else 'Die feste Zahl'} rechnet nicht mit, wenn sich etwas davor ändert.",
                        f" {'The fixed date' if dated else 'The fixed number'} does not update when something before it "
                        f"changes."))
        checked = _checked(sheet, expected, like)
        # Ergibt die Formel dieselbe Zahl, ist heute nichts falsch: nachfragen, mit der Formel zum Einsetzen.
        found.append(Finding(f"typedcell:{address(pos)}", text, "pruefen" if same else "fehler",
                             T(f"Korrigiert: {checked}.", f"Corrected: {checked}.") if checked else "", fix=fix))
    return found[:MAX_FOUND]
