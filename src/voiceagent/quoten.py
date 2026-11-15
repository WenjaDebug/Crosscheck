"""Quoten, Mittelwerte und Wachstum, die falsch gewichtet oder auf die falsche Basis gerechnet sind.

Der Mittelwert der Quoten je Zeile (Komplikationen/Fälle je Klinik) ist nicht die Quote über alle Zeilen: Eine
Klinik mit 12 Fällen zählt dann so viel wie eine mit 900. Ebenso der Mittelwert von Gruppen-Mittelwerten, wenn
die Gruppen verschieden groß sind. Und Wachstum gehört durch den alten Wert geteilt, nicht durch den neuen.

Alles ohne KI: Der Prüfer liest die Formeln und Spaltenköpfe (Jahre, Monate, Vorjahr) und rechnet beides nach.
Ob die ungewichtete Zahl gemeint ist, weiß nur, wer die Tabelle gemacht hat, darum heißt es dort „prüfen“.
"""

from __future__ import annotations

import math
import re

from . import formel, woerter
from .analyse import Finding, Sheet, _fmt, address
from .bereiche import _FILLER, _MONTHS, _QUARTERS, _german_formula, _reach, _use, _words, row_rates, verified
from .kopf import _fix_from_tip, _tip_pair, _tip_plus
from .sprache import T
from .workbook import _left_label, pattern
from .xlsx import column_letters, number_text, shift

Pos = tuple[int, int]

MAX_EACH = 2
_GROWTH = ("wachstum", "veränderung", "änderung", "zuwachs", "zunahme", "steigerung", "growth", "change", "delta",
           "δ", "ggü", "gegenüber", "vs", "vj", "yoy")
_DECLINE = ("rückgang", "abnahme", "minus", "verlust", "decline", "drop")
_OLD = {"vorjahr": -1, "vj": -1, "vormonat": -1, "vorperiode": -1, "alt": -1, "previous": -1, "last": -1}
_NEW = {"aktuell": 0, "neu": 0, "laufend": 0, "current": 0, "this": 0}


def checks(sheet: Sheet, rows: list[int], columns: dict) -> list[Finding]:
    if len(rows) < 3:
        return []
    return (_ratio_means(sheet, rows, columns) + _mean_of_means(sheet, rows, columns) + _growth(sheet, rows)
            + _named_base(sheet, rows))


def _shown(sheet: Sheet, pos: Pos, value) -> str:
    """Wie die Zelle es zeigt: als Prozent, wenn sie Prozent zeigt."""
    cell = sheet.cells.get(pos)
    if isinstance(value, float) and cell is not None and isinstance(cell.value, str) and cell.value.strip().endswith("%"):
        return f"{number_text(round(value * 100, 2))} %"
    return _fmt(value)


def _tipped(sheet: Sheet, item: Finding, german: str) -> Finding:
    _fix_from_tip(sheet, item, german)
    return item


def _label(sheet: Sheet, pos: Pos, rows: set[int]) -> str:
    """Wie eine Zelle heißt: in der Tabelle nach ihrem Spaltenkopf, daneben nach dem Wort links."""
    if pos[0] in rows and pos[1] in sheet.plan.headers:
        return sheet.plan.headers[pos[1]]
    return _left_label(sheet.cells, *pos)


def _number(sheet: Sheet, pos: Pos) -> float | None:
    v = sheet.inputs.get(pos)
    return v if isinstance(v, float) and not isinstance(v, bool) else None


# --- Mittelwert von Quoten ---


def _unguard(node):
    def zero(n) -> bool:
        return (n.kind == "num" and n.value == 0) or (n.kind == "str" and n.value == "")
    if node.kind == "call" and node.value == "IFERROR" and len(node.args) == 2 and zero(node.args[1]):
        return node.args[0]
    if node.kind == "call" and node.value == "IF" and len(node.args) == 3:
        cond, then, other = node.args
        if cond.kind == "op" and len(cond.args) == 2 and cond.args[1].kind == "num" and cond.args[1].value == 0:
            if cond.value == "=" and zero(then):
                return other
            if cond.value == "<>" and zero(other):
                return then
    return node


def _quotient(node, row: int) -> tuple[float, int, int] | None:
    """X/Y in derselben Zeile, auch mal 100: (Faktor, Spalte X, Spalte Y). Ein Schutz vor dem Teilen durch 0
    (=WENN(H18=0;0;K18/H18), =WENNFEHLER(K18/H18;0)) ändert an der Quote nichts."""
    node = _unguard(node)
    factor = 1.0
    if node.kind == "op" and node.value == "*" and len(node.args) == 2:
        num, other = (node.args[0], node.args[1]) if node.args[0].kind == "num" else (node.args[1], node.args[0])
        if num.kind != "num":
            return None
        factor, node = num.value, other
    if node.kind != "op" or node.value != "/" or len(node.args) != 2:
        return None
    top, bottom = node.args
    if top.kind == "op" and top.value == "*" and len(top.args) == 2 and top.args[1].kind == "num" and factor == 1.0:
        factor, top = top.args[1].value, top.args[0]
    if top.kind != "ref" or bottom.kind != "ref" or top.value[0] != row or bottom.value[0] != row:
        return None
    if top.value[1] == bottom.value[1]:
        return None
    return factor, top.value[1], bottom.value[1]


def _ratio_means(sheet: Sheet, rows: list[int], columns: dict) -> list[Finding]:
    """MITTELWERT über eine Spalte mit Quoten je Zeile (C2/D2), wo die Nenner verschieden groß sind."""
    plan = sheet.plan
    calc, _ = sheet.evaluator()
    found = []
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown or node.kind != "call" or node.value != "AVERAGE" or len(node.args) != 1:
            continue
        arg = node.args[0]
        if arg.kind != "range" or arg.value[1] != arg.value[3] or arg.value[1] not in plan.headers:
            continue
        label = sheet.figure_label(pos)
        if re.search(r"ungewichtet|unweighted|gleich gewichtet", label.casefold()):
            continue
        r1, col, r2, _ = arg.value
        inside = [r for r in rows if r1 <= r <= r2]
        cells = [(r, col) for r in inside if (r, col) in sheet.nodes and sheet.cells[(r, col)].formula]
        if len(cells) < 3 or len(cells) < len(inside) * 0.8:
            continue
        first = cells[0]
        shape = _quotient(sheet.nodes[first], first[0])
        if shape is None:
            rated = _rates_at_total(sheet, calc, pos, node, label, col, cells, rows)
            if rated is not None:
                found.append(rated)
                if len(found) >= MAX_EACH:
                    break
            continue
        same = pattern(sheet.cells[first].formula, first)
        if any(pattern(sheet.cells[p].formula, p) != same for p in cells):
            continue
        factor, cx, cy = shape
        used = []
        for r, _ in cells:
            x, y, q = _number(sheet, (r, cx)), _number(sheet, (r, cy)), _number(sheet, (r, col))
            if x is None or y is None or q is None or y == 0:
                continue
            used.append((r, x, y))
        if len(used) < 3:
            continue
        # Nur Anteile (Treffer von Fällen, nie mehr als der Nenner) oder eine Marge auf den Umsatz der Zeile (das
        # Ergebnis darf negativ sein). Sind alle Nenner gleich, ist es dasselbe.
        margin = woerter.rate_word(_words(plan.headers[col])) or "%" in plan.headers[col]
        if any(not (0 <= x <= y or (margin and y > 0 and x <= y)) for _, x, y in used) or len({y for _, _, y in used}) == 1:
            continue
        try:
            value = calc.run(node)
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
            continue
        if not isinstance(value, float):
            continue
        total = factor * math.fsum(x for _, x, _ in used) / math.fsum(y for _, _, y in used)
        if abs(value - total) <= 0.005 * abs(total) or abs(value - total) < 1e-9:
            continue
        top, bottom = used[0][0], used[-1][0]
        english = f"SUM({column_letters(cx)}{top}:{column_letters(cx)}{bottom})/SUM({column_letters(cy)}{top}:{column_letters(cy)}{bottom})"
        if factor != 1.0:
            english = f"{english}*{repr(factor).removesuffix('.0')}"
        formula = verified(sheet, [english], total, pos)
        hx, hy, hq = plan.headers[cx], plan.headers[cy], plan.headers[col]
        small = min(used, key=lambda u: u[2])
        big = max(used, key=lambda u: u[2])
        span = f"{address((r1, col))}:{address((min(r2, plan.last_row), col))}"
        right_now, shown_now = _shown(sheet, first, total), _shown(sheet, first, value)
        text = T(f"{label} mittelt die Quoten {span} ({hq}) "
                 f"ungewichtet: Jede Zeile zählt gleich viel, ob {hy} {_fmt(small[2])} ist (Zeile {small[0]}) oder "
                 f"{_fmt(big[2])} (Zeile {big[0]}). Über alle Zeilen zusammen, Summe {hx} durch Summe {hy}, ist die Quote "
                 f"{right_now} statt {shown_now}.",
                 f"{label} averages the rates {span} ({hq}) unweighted: every row counts the same, whether {hy} is "
                 f"{_fmt(small[2])} (row {small[0]}) or {_fmt(big[2])} (row {big[0]}). Across all rows together, "
                 f"sum of {hx} divided by sum of {hy}, the rate is {right_now} instead of {shown_now}.")
        tip = _tip_pair("Gemeint ist meist die Quote über alle Zeilen"
                + (f": ={formula}, nachgerechnet {right_now}." if formula else ".")
                + " Soll jede Zeile gleich viel zählen (etwa jede Klinik gleich), ist der Mittelwert richtig; "
                "dann schreib „ungewichtet“ dazu.",
                "What is usually meant is the rate across all rows"
                + (f": ={formula}, recalculated {right_now}." if formula else ".")
                + " If every row should count the same (say, every clinic equally), the mean is right; "
                "then add “unweighted” to the label.")
        # „Jahresmarge“, „Marge gesamt“: Der Name verlangt die Quote über alles, nicht das Mittel der Monate.
        whole = re.search(r"jahres|gesamt|insgesamt|total|overall|annual|\bjahr\b|ganze", label.casefold())
        fix = None
        if formula:
            from . import einsetzen
            fix = einsetzen.formulas(sheet, {pos: english})
        found.append(_tipped(sheet, Finding(f"ratio:{address(pos)}", text, "fehler" if whole and fix else "pruefen",
                                            tip[0], fix=fix), tip[1]))
        if len(found) >= MAX_EACH:
            break
    return found


_TOTAL_ROW = re.compile(r"summe|gesamt|total|insgesamt|jahr|year|alle", re.IGNORECASE)


def _rates_at_total(sheet: Sheet, calc, pos: Pos, node, label: str, col: int, cells: list[Pos],
                    rows: list[int]) -> Finding | None:
    """„Summe / Gesamt“ =MITTELWERT(E16:E27) über Auslastung = Pflegetage / (Betten * Tage): Jeder Monat zählt gleich,
    ob er 28 oder 31 Tage hat. Die Quote über alle Zeilen ist dieselbe Formel wie in jeder Zeile, nur mit den Summen
    der Summenzeile (=D28/($B$5*B28)). Gemeldet nur, wenn die Summenzeile die Summen wirklich hat."""
    if set(cells) - set(row_rates(sheet, col, rows)):
        return None
    first = cells[0]
    text = sheet.cells[first].formula or ""
    singles, _, _ = formel.references(sheet.nodes[first], sheet.name)
    own = sorted({c for (r, c) in singles if r == first[0]})
    if not own or col in own:
        return None
    # In der Summenzeile steht unter jeder dieser Spalten die Summe der Zeilen darüber.
    inside = [r for r, _ in cells]
    for c in own:
        total = _number(sheet, (pos[0], c))
        if total is None:
            try:
                total = calc.run(sheet.nodes[(pos[0], c)]) if (pos[0], c) in sheet.nodes else None
            except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
                total = None
        parts = [_number(sheet, (r, c)) for r in inside]
        if not isinstance(total, float) or any(v is None for v in parts) or \
                abs(total - math.fsum(parts)) > 1e-6 * max(1.0, abs(total)):
            return None
    english = shift(text, pos[0] - first[0], 0)
    try:
        value = calc.run(node)
        expected = calc.run(formel.parse(english))
    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
        return None
    if not isinstance(value, float) or not isinstance(expected, float) or abs(value - expected) <= 1e-6 * max(1.0, abs(expected)):
        return None
    if not verified(sheet, [english], expected, pos):
        return None
    from . import einsetzen

    plan = sheet.plan
    head = plan.headers.get(col, column_letters(col))
    shown_now, shown_right = _shown(sheet, first, value), _shown(sheet, first, expected)
    where = f"{address(cells[0])}:{address(cells[-1])}"
    name_cell = sheet.cells.get((pos[0], plan.name_column)) if plan.name_column else None
    row_name = (name_cell.value.strip() if name_cell is not None and name_cell.formula is None
                and isinstance(name_cell.value, str) else "") or _left_label(sheet.cells, *pos)
    named = T(f"{address(pos)} („{row_name}“ unter {head})", f"{address(pos)} (“{row_name}” under {head})") \
        if row_name else T(f"„{label}“ in {address(pos)}", f"“{label}” in {address(pos)}")
    text_out = T(f"{named} mittelt die Quoten {where}: Jede Zeile zählt gleich viel, egal "
                 f"wie groß sie ist. Die Quote über alle Zeilen rechnet wie jede Zeile, aber mit den Summen aus Zeile "
                 f"{pos[0]}: ={_german_formula(english)}. Richtig wäre {shown_right} statt {shown_now}.",
                 f"{named} averages the rates {where}: every row counts the same, no matter how large it is. The "
                 f"rate across all rows is calculated like each row, but with the totals from row "
                 f"{pos[0]}: ={_german_formula(english)}. It should be {shown_right} instead of {shown_now}.")
    totalish = bool(_TOTAL_ROW.search(label or "")) or bool(_TOTAL_ROW.search(_left_label(sheet.cells, *pos) or ""))
    fix = einsetzen.formulas(sheet, {pos: english}) if totalish else None
    tip = _tip_pair(f"Schreib in {address(pos)} ={_german_formula(english)}.",
            f"In {address(pos)}, enter ={_german_formula(english)}.")
    if not totalish:
        tip = _tip_plus(tip, _tip_pair(
            " Soll jede Zeile gleich viel zählen, ist der Mittelwert richtig; dann schreib „ungewichtet“ dazu.",
            " If every row should count the same, the mean is right; then add “unweighted” to the label."))
    return _tipped(sheet, Finding(f"ratio:{address(pos)}", text_out, "fehler" if fix else "pruefen", tip[0], fix=fix),
                   tip[1])


# --- Mittelwert von Mittelwerten ---


def _mean_of_means(sheet: Sheet, rows: list[int], columns: dict) -> list[Finding]:
    """MITTELWERT(H3;H5) oder (H3+H5)/2, wenn H3 und H5 Mittelwerte verschieden großer Gruppen sind."""
    data = set(rows)
    calc, _ = sheet.evaluator()
    found = []
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown:
            continue
        refs = _averaged_refs(node)
        if refs is None or len(refs) < 2 or len(set(refs)) != len(refs):
            continue
        if re.search(r"ungewichtet|unweighted|gleich gewichtet", sheet.figure_label(pos).casefold()):
            continue
        targets = set.intersection(*(_reach(sheet, ref, data, columns) for ref in refs))
        if len(targets) != 1:
            continue
        target = targets.pop()
        groups = []
        for ref in refs:
            use = _use(sheet, sheet.nodes.get(ref), target, data) if ref in sheet.nodes else None
            if use is None or use.kind != "mean" or not use.rows:
                break
            groups.append(set(use.rows))
        else:
            union = set().union(*groups)
            if sum(len(g) for g in groups) != len(union) or len({len(g) for g in groups}) == 1:
                continue
            numbers = columns[target].numbers
            if any(r not in numbers for r in union):
                continue
            weighted = math.fsum(numbers[r] for r in union) / len(union)
            try:
                value = calc.run(node)
            except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
                continue
            if not isinstance(value, float) or abs(value - weighted) <= 1e-9 * max(1.0, abs(weighted)):
                continue
            parts = [T(f"{address(ref)} ({_left_label(sheet.cells, *ref) or 'Mittelwert'}, {len(g)} Zeilen)",
                       f"{address(ref)} ({_left_label(sheet.cells, *ref) or 'mean'}, {len(g)} "
                       f"{'row' if len(g) == 1 else 'rows'})")
                     for ref, g in zip(refs, groups)]
            listed = ", ".join(parts[:-1]) + T(" und ", " and ") + parts[-1]
            letter = column_letters(target)
            candidates = []
            if union == {r for r in rows if r in numbers}:
                candidates.append(f"AVERAGE({letter}{min(union)}:{letter}{max(union)})")
            candidates += _pooled(sheet, refs)
            formula = verified(sheet, candidates, weighted, pos)
            text = T(f"{sheet.figure_label(pos)} mittelt die Mittelwerte {listed} zu gleichen Teilen. Über alle "
                     f"{len(union)} Zeilen ist der Mittelwert {_fmt(weighted)} statt {_fmt(value)}: So zählt jede Gruppe "
                     "gleich viel, egal wie viele Zeilen sie hat.",
                     f"{sheet.figure_label(pos)} averages the means {listed} in equal parts. Across all "
                     f"{len(union)} rows, the mean is {_fmt(weighted)} instead of {_fmt(value)}: this way every group "
                     "counts the same, no matter how many rows it has.")
            tip = _tip_pair((f"Über alle Zeilen: ={formula}, nachgerechnet {_fmt(weighted)}." if formula
                     else "Rechne den Mittelwert direkt über alle Zeilen der Gruppen, nicht über ihre Mittelwerte.")
                    + " Soll jede Gruppe gleich zählen, ist es richtig; dann schreib „ungewichtet“ dazu.",
                    (f"Across all rows: ={formula}, recalculated {_fmt(weighted)}." if formula
                     else "Calculate the mean directly across all rows of the groups, not across their means.")
                    + " If every group should count the same, it is right; then add “unweighted” to the label.")
            found.append(_tipped(sheet, Finding(f"means:{address(pos)}", text, "pruefen", tip[0]), tip[1]))
            if len(found) >= MAX_EACH:
                break
    return found


def _averaged_refs(node) -> list[Pos] | None:
    """MITTELWERT(H3;H5) → [H3, H5], (H3+H5)/2 → [H3, H5]."""
    if node.kind == "call" and node.value == "AVERAGE" and all(a.kind == "ref" for a in node.args):
        return [a.value for a in node.args]
    if node.kind == "op" and node.value == "/" and len(node.args) == 2 and node.args[1].kind == "num":
        terms = _plus_refs(node.args[0])
        if terms and len(terms) == node.args[1].value:
            return terms
    return None


def _plus_refs(node) -> list[Pos] | None:
    if node.kind == "ref":
        return [node.value]
    if node.kind == "op" and node.value == "+" and len(node.args) == 2:
        left, right = _plus_refs(node.args[0]), _plus_refs(node.args[1])
        return left + right if left is not None and right is not None else None
    if node.kind == "call" and node.value == "SUM" and all(a.kind == "ref" for a in node.args):
        return [a.value for a in node.args]
    return None


def _pooled(sheet: Sheet, refs: list[Pos]) -> list[str]:
    """Aus MITTELWERTWENN(B:B;"Nord";D:D) je Gruppe: (SUMMEWENN(…)+…)/(ZÄHLENWENNS(…;D:D;"<>")+…)."""
    sums, counts = [], []
    for ref in refs:
        text = (sheet.cells[ref].formula or "").replace("_xlfn.", "").replace(" ", "")
        m = re.fullmatch(r"AVERAGEIF\(([^,()]+),([^,()]+),([^,()]+)\)", text)
        if m:
            area, crit, avg = m.groups()
            sums.append(f"SUMIF({area},{crit},{avg})")
            counts.append(f"COUNTIFS({area},{crit},{avg},\"<>\")")
            continue
        m = re.fullmatch(r"AVERAGEIFS\(([^,()]+),(.+)\)", text)
        if m:
            avg, rest = m.groups()
            sums.append(f"SUMIFS({avg},{rest})")
            counts.append(f"COUNTIFS({rest},{avg},\"<>\")")
            continue
        return []
    return [f"({'+'.join(sums)})/({'+'.join(counts)})"]


# --- Wachstum auf der falschen Basis ---


def _period(text: str) -> tuple | None:
    """„Umsatz 2025“ → (2025, None, None), „März“ → (None, 3, None), „Vorjahr“ → (None, None, -1)."""
    years = re.findall(r"(?<!\d)((?:19|20)\d\d)(?!\d)", text)
    words = _words(text)
    months = [_MONTHS[w] for w in words if w in _MONTHS]
    quarters = [min(_QUARTERS[w]) for w in words if w in _QUARTERS]
    rel = [_OLD.get(w, _NEW.get(w)) for w in words if w in _OLD or w in _NEW]
    if len(set(years)) > 1 or len(set(months + quarters)) > 1 or len(set(rel)) > 1:
        return None
    period = (int(years[0]) if years else None, (months + quarters or [None])[0], rel[0] if rel else None)
    return period if any(p is not None for p in period) else None


def _older(a: tuple, b: tuple) -> int | None:
    """-1, wenn a vor b liegt, 1 wenn danach, None wenn es sich nicht sagen lässt."""
    if a[0] is not None and b[0] is not None and a[0] != b[0]:
        return -1 if a[0] < b[0] else 1
    if (a[0] == b[0]) and a[1] is not None and b[1] is not None and a[1] != b[1]:
        return -1 if a[1] < b[1] else 1
    if a[2] is not None and b[2] is not None and a[2] != b[2]:
        return -1 if a[2] < b[2] else 1
    return None


def _change(node) -> tuple[float, Pos, Pos, Pos] | None:
    """(A−B)/C, auch mal 100: (Faktor, A, B, C)."""
    factor = 1.0
    if node.kind == "op" and node.value == "*" and len(node.args) == 2:
        num, other = (node.args[0], node.args[1]) if node.args[0].kind == "num" else (node.args[1], node.args[0])
        if num.kind != "num":
            return None
        factor, node = num.value, other
    if node.kind != "op" or node.value != "/" or len(node.args) != 2:
        return None
    top, base = node.args
    if top.kind != "op" or top.value != "-" or len(top.args) != 2 or base.kind != "ref":
        return None
    a, b = top.args
    if a.kind != "ref" or b.kind != "ref":
        return None
    return factor, a.value, b.value, base.value


def _growth(sheet: Sheet, rows: list[int]) -> list[Finding]:
    """Wachstum (neu − alt) / neu statt / alt, oder alt − neu: Vorzeichen umgedreht."""
    data = set(rows)
    calc, now = sheet.evaluator()
    found, seen = [], set()
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown or pos in seen:
            continue
        shape = _change(node)
        if shape is None:
            continue
        factor, a, b, base = shape
        if base not in (a, b):
            continue
        label = _label(sheet, pos, data) or sheet.figure_label(pos)
        words = label.casefold()
        if not any(w in _words(words) or w in words for w in _GROWTH) or any(w in words for w in _DECLINE):
            continue
        pa, pb = _period(_label(sheet, a, data)), _period(_label(sheet, b, data))
        if pa is None or pb is None:
            continue
        order = _older(pa, pb)
        if order is None:
            continue
        new, old = (a, b) if order == 1 else (b, a)
        wrong_base, flipped = base == new, a == old
        if not (wrong_base or flipped):
            continue
        try:
            value = calc.run(node)
            right = factor * (now(new) - now(old)) / now(old)
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
            continue
        if not isinstance(value, float) or not isinstance(right, float) or abs(value - right) <= 1e-9 * max(1.0, abs(right)):
            continue
        # Heruntergezogen: ein Hinweis für die ganze Spalte.
        run = [pos]
        while (nxt := (run[-1][0] + 1, pos[1])) in sheet.nodes and sheet.cells.get(nxt) is not None and \
                sheet.cells[nxt].formula and pattern(sheet.cells[nxt].formula, nxt) == pattern(sheet.cells[pos].formula, pos):
            run.append(nxt)
        seen |= set(run)
        english = f"({address(new)}-{address(old)})/{address(old)}"
        if factor != 1.0:
            english = f"{english}*{repr(factor).removesuffix('.0')}"
        formula = verified(sheet, [english], right, pos)
        what = []
        if flipped:
            what.append(T(f"rechnet {address(old)} − {address(new)}, also alt minus neu: Das Vorzeichen ist umgedreht",
                          f"calculates {address(old)} − {address(new)}, i.e. old minus new: the sign is reversed"))
        if wrong_base:
            what.append(T(f"teilt durch {address(new)} ({_label(sheet, new, data)}), den neuen Wert. Wachstum misst man "
                          f"gegenüber dem alten Wert {address(old)} ({_label(sheet, old, data)})",
                          f"divides by {address(new)} ({_label(sheet, new, data)}), the new value. Growth is measured "
                          f"against the old value {address(old)} ({_label(sheet, old, data)})"))
        span = f"{address(run[0])}:{address(run[-1])}"
        where = T(f" Das gilt für alle {len(run)} Zellen {span}.", f" This applies to all {len(run)} cells {span}.") \
            if len(run) > 1 else ""
        right_now, shown_now = _shown(sheet, pos, right), _shown(sheet, pos, value)
        text = T(f"{address(pos)} ({label}) {' und '.join(what)}. Richtig wäre {right_now} statt "
                 f"{shown_now}.{where}",
                 f"{address(pos)} ({label}) {' and '.join(what)}. It should be {right_now} instead of "
                 f"{shown_now}.{where}")
        if not formula:
            tip = _tip_pair(f"Teil die Veränderung {address(new)} − {address(old)} durch den alten Wert {address(old)}.",
                    f"Divide the change {address(new)} − {address(old)} by the old value {address(old)}.")
        elif len(run) > 1:
            tip = _tip_pair(f"Schreib in {address(pos)} ={formula} und zieh es herunter, nachgerechnet {right_now}.",
                    f"In {address(pos)}, enter ={formula} and fill it down, recalculated {right_now}.")
        else:
            tip = _tip_pair(f"Nimm ={formula}, nachgerechnet {right_now}.", f"Use ={formula}, recalculated {right_now}.")
        found.append(_tipped(sheet, Finding(f"growth:{address(pos)}", text, "fehler", tip[0]), tip[1]))
        if len(found) >= MAX_EACH:
            break
    return found


# --- Durch eine andere Spalte geteilt, als der Name sagt ---

# „Marge (% vom Umsatz)“, „Anteil an Fällen“, „Kosten je Fall“, „in % des Budgets“, „share of revenue“: Was danach
# steht, ist die Basis, durch die geteilt wird.
_BASE = re.compile(r"(?:%|prozent|anteil|quote|rate|share|percent|in %)\s*(?:vom|von|der|des|am|an|of|on|from)\s+([^()]+)"
                   r"|\b(?:je|pro|per)\s+([^()]+)", re.IGNORECASE)
_UNIT_WORDS = {"t", "tsd", "teur", "k", "keur", "mio", "mrd", "chf", "eur", "euro", "usd", "fr", "%", "prozent"}


def _key_words(text: str) -> set[str]:
    return {w for w in _words(re.sub(r"\([^)]*\)", " ", text)) if w not in _FILLER and w not in _UNIT_WORDS and len(w) > 1}


def _named_base(sheet: Sheet, rows: list[int]) -> list[Finding]:
    """E5 =D5/C5 heruntergezogen in „Marge (% vom Umsatz)“, aber C heißt „Kosten“ und B „Umsatz“: Der Name sagt,
    durch welche Spalte geteilt wird. Gibt es genau eine Spalte mit diesem Namen, gehört sie in den Nenner."""
    from . import einsetzen

    plan = sheet.plan
    data = set(rows)
    calc, _ = sheet.evaluator()
    found, seen = [], set()
    for pos, node in sorted(sheet.nodes.items(), key=lambda item: (item[0][1], item[0][0])):
        if node is None or pos in sheet.unknown or pos in seen or pos[0] not in data or pos[1] not in plan.headers:
            continue
        shape = _quotient(node, pos[0])
        if shape is None:
            continue
        _, top, bottom = shape
        header = plan.headers[pos[1]]
        match = _BASE.search(header)
        if not match:
            continue
        base = _key_words(match.group(1) or match.group(2) or "")
        if not base:
            continue
        # Die Spalte, die genau so heißt („Umsatz (T€)“), sonst die einzige, deren Name die Wörter enthält.
        named = [c for c, h in plan.headers.items() if c != pos[1] and base == _key_words(h)] or \
            [c for c, h in plan.headers.items() if c != pos[1] and base <= _key_words(h)]
        if len(named) != 1 or named[0] in (top, bottom):
            continue
        right_col = named[0]
        if base & _key_words(plan.headers[bottom]) or right_col == pos[1]:
            continue
        # Heruntergezogen: dieselbe Formel in der ganzen Spalte, ein Hinweis für alle.
        shape_text = pattern(sheet.cells[pos].formula, pos)
        run = [p for p in sorted(sheet.nodes) if p[1] == pos[1] and p[0] in data and sheet.cells.get(p) is not None
               and sheet.cells[p].formula and pattern(sheet.cells[p].formula, p) == shape_text]
        seen |= set(run)
        changes = {}
        for p in run:
            text = sheet.cells[p].formula or ""
            wrong = re.compile(rf"(?<![A-Za-z0-9_.$!])(\$?){column_letters(bottom)}(\$?){p[0]}(?![0-9A-Za-z(!])")
            hits = wrong.findall(text)
            if len(hits) != 1:
                changes = {}
                break
            changes[p] = wrong.sub(lambda m: f"{m.group(1)}{column_letters(right_col)}{m.group(2)}{p[0]}", text)
        if not changes:
            continue
        try:
            now = calc.run(node)
            right = calc.run(formel.parse(changes[pos]))
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
            continue
        if not isinstance(now, float) or not isinstance(right, float) or abs(now - right) <= 1e-9 * max(1.0, abs(right)):
            continue
        fix = einsetzen.formulas(sheet, changes)
        where = f"{address(run[0])}:{address(run[-1])}" if len(run) > 1 else address(pos)
        text = T(f"{where} ({header}) {'teilen' if len(run) > 1 else 'teilt'} durch die Spalte {column_letters(bottom)} "
                 f"„{plan.headers[bottom]}“, der Name sagt aber „{match.group(0).strip()}“, also durch die Spalte "
                 f"{column_letters(right_col)} „{plan.headers[right_col]}“. Richtig wäre in {address(pos)} "
                 f"{_shown(sheet, pos, right)} statt {_shown(sheet, pos, now)}.",
                 f"{where} ({header}) {'divide' if len(run) > 1 else 'divides'} by column {column_letters(bottom)} "
                 f"“{plan.headers[bottom]}”, but the name says “{match.group(0).strip()}”, so it should be column "
                 f"{column_letters(right_col)} “{plan.headers[right_col]}”. In {address(pos)}, it should be "
                 f"{_shown(sheet, pos, right)} instead of {_shown(sheet, pos, now)}.")
        tip = _tip_pair(f"Schreib in {address(run[0])} ={_german_formula(changes[run[0]])}"
                + (f" und zieh es herunter bis {address(run[-1])}." if len(run) > 1 else "."),
                f"In {address(run[0])}, enter ={_german_formula(changes[run[0]])}"
                + (f" and fill it down to {address(run[-1])}." if len(run) > 1 else "."))
        found.append(_tipped(sheet, Finding(f"nenner:{address(run[0])}", text, "fehler", tip[0], fix=fix), tip[1]))
        if len(found) >= MAX_EACH:
            break
    return found
