"""Sensitivitätstabellen als Gegenprobe des Modells.

Eine Sensitivität rechnet das Ergebnis (den Wert je Aktie) noch einmal, für jede Kombination zweier Annahmen (WACC in
den Zeilen, Wachstum in den Spalten). Die Zeile und die Spalte, deren Annahmen genau die des Modells sind, muss
dasselbe zeigen wie das Modell selbst. Tut sie das nicht, stimmt eine der beiden Rechnungen nicht. Bringen die
Korrekturen der übrigen Hinweise beide zusammen, kommt der Unterschied von diesen Fehlern.

Eine Tabelle ist ein Rechteck aus Formeln, die gleich rechnen und je ihre Zeile in einer Kopfspalte ($A51) und ihre
Spalte in einer Kopfzeile (B$50) lesen. Die Köpfe sind die Annahme des Modells plus Stufen (=$B$15-0,01) oder
eingetippte Werte, von denen einer der Annahme gleicht. Eine Datentabelle von Excel ({=MEHRFACHOPERATION(…)}) rechnet
Excel selbst, die bleibt außen vor.
"""

from __future__ import annotations

import math
import re
from collections import Counter
from dataclasses import dataclass

from . import formel
from .sprache import T
from .workbook import pattern

Pos = tuple[int, int]

# Größer rechnet die Gegenprobe mit den Korrekturen nicht alles neu.
MAX_FORMULAS = 5000


@dataclass(frozen=True)
class Table:
    rows: tuple[int, ...]
    cols: tuple[int, ...]
    # Die Spalte mit den Köpfen der Zeilen (A) und die Zeile mit den Köpfen der Spalten (50).
    head_col: int
    head_row: int
    # Zellen im Rechteck, die anders rechnen (ein Kopierfehler): Die meldet die Prüfung der Muster.
    odd: tuple = ()

    def heads(self) -> set[Pos]:
        return {(r, self.head_col) for r in self.rows} | {(self.head_row, c) for c in self.cols}

    def body(self) -> set[Pos]:
        return {(r, c) for r in self.rows for c in self.cols}


def tables(cells) -> list[Table]:
    """Die Sensitivitätstabellen auf einem Blatt (Zellen wie gelesen)."""
    keyed: dict[Pos, set[Pos]] = {}
    nodes: dict[Pos, formel.Node] = {}
    for pos, cell in cells.items():
        # Ohne $ lässt sich eine Formel nicht in beide Richtungen ziehen.
        if not cell.formula or "$" not in cell.formula:
            continue
        try:
            node = formel.parse(cell.formula)
        except formel.Unknown:
            continue
        nodes[pos] = node
        singles, _, unknown = formel.references(node)
        if unknown:
            continue
        r, c = pos
        lefts = {p[1] for p in singles if p[0] == r and p[1] < c}
        tops = {p[0] for p in singles if p[1] == c and p[0] < r}
        for x in lefts:
            for y in tops:
                keyed.setdefault((x, y), set()).add(pos)
    found: list[Table] = []
    taken: set[Pos] = set()
    for (x, y), body in sorted(keyed.items(), key=lambda item: -len(item[1])):
        if len(body) < 4 or body & taken:
            continue
        rows = sorted({p[0] for p in body})
        cols = sorted({p[1] for p in body})
        if len(rows) < 2 or len(cols) < 2:
            continue
        if rows != list(range(rows[0], rows[-1] + 1)) or cols != list(range(cols[0], cols[-1] + 1)):
            continue
        if not (x < cols[0] and y < rows[0]):
            continue
        # Eine Zelle, die mit Kopierfehler eine andere Zeile liest, gehört trotzdem dazu (sie meldet das Muster).
        missing = [(r, c) for r in rows for c in cols if (r, c) not in body]
        if len(missing) > 0.2 * len(rows) * len(cols) or any(not (cells.get(p) and cells[p].formula) for p in missing):
            continue
        shapes = {p: pattern(cells[p].formula, p) for p in body}
        common, count = Counter(shapes.values()).most_common(1)[0]
        if count < 0.8 * len(rows) * len(cols):
            continue
        odd = sorted(missing + [p for p, shape in shapes.items() if shape != common])
        table = Table(tuple(rows), tuple(cols), x, y, tuple(odd))
        if any(_number(cells.get(p)) is None for p in table.heads()):
            continue
        found.append(table)
        taken |= body
    return found


def axes(cells) -> set[Pos]:
    """Die Köpfe aller Sensitivitätstabellen: Die Stufen darin (=$B$15-0,01) gehören in keine eigene Zelle."""
    return {p for t in tables(cells) for p in t.heads()}


def _number(cell) -> float | None:
    if cell is None:
        return None
    from .workbook import _as_number

    return _as_number(cell.value)


def _words(text: str) -> set[str]:
    from . import woerter

    return {w for w in re.findall(r"[^\W_]+", woerter.spaced(text).casefold()) if len(w) > 1 or w.isalpha()}


def checks(sheet, findings: list) -> list:
    """Der Hinweis, wenn die Sensitivität dort, wo sie mit den Annahmen des Modells rechnet, etwas anderes zeigt als
    das Ergebnis des Modells."""
    from .analyse import Finding

    result: list[Finding] = []
    for table in tables(sheet.cells):
        axis = _odd_heads(sheet, table)
        if axis:
            result += axis
            continue
        item = _check(sheet, table, findings)
        if item is not None:
            result.append(Finding(*item[:4], fix=item[4] if len(item) > 4 else None))
    return result


def _odd_heads(sheet, table: Table) -> list:
    """Ein Kopf, der eine andere Annahme liest als die übrigen derselben Achse (=$B$14 zwischen lauter $B$15 ± Stufe):
    Dann rechnet die ganze Zeile der Sensitivität mit der falschen Größe."""
    from . import einsetzen
    from .analyse import Finding, _left_label, address
    from .zeilenmodell import _to_cell

    found = []
    inside = table.body() | table.heads()
    for heads, word in (([(r, table.head_col) for r in table.rows], "Zeilen"),
                        ([(table.head_row, c) for c in table.cols], "Spalten")):
        refs = {}
        for pos in heads:
            node = sheet.nodes.get(pos)
            if node is not None:
                refs[pos] = {p for p in formel.references(node, sheet.name)[0] if p not in inside}
        if len(refs) < 3:
            continue
        counts = Counter(p for v in refs.values() for p in v)
        common, count = counts.most_common(1)[0]
        odd = [pos for pos, v in refs.items() if common not in v]
        if count < 2 or len(odd) != 1 or len(refs[odd[0]]) != 1:
            continue
        pos = odd[0]
        # Der Kopf, von dem die anderen in Stufen ausgehen (=$C$131 in der Mitte, darüber =B148-$C$43), ist der Anker:
        # Er liest die Annahme selbst, nicht die Stufe.
        if any(pos in formel.references(sheet.nodes[p], sheet.name)[0] for p in refs if p != pos):
            continue
        wrong = next(iter(refs[pos]))
        formula = _to_cell(sheet.cells[pos].formula, wrong, common)
        fix = einsetzen.formulas(sheet, {pos: formula})
        if fix is None:
            continue
        name = _left_label(sheet.cells, *common) or address(common)
        other = _left_label(sheet.cells, *wrong) or address(wrong)
        line = "row" if word == "Zeilen" else "column"
        text = T(f"{address(pos)} ist ein Kopf der Sensitivität und liest {address(wrong)} ({other}). Die übrigen Köpfe "
                 f"der {word} lesen {address(common)} ({name}). So rechnet die ganze "
                 f"{'Zeile' if word == 'Zeilen' else 'Spalte'} von {address(pos)} mit {other} statt mit {name}.",
                 f"{address(pos)} is a heading of the sensitivity table and reads {address(wrong)} ({other}). The other "
                 f"{line} headings read {address(common)} ({name}). So the whole {line} of {address(pos)} calculates "
                 f"with {other} instead of {name}.")
        found.append(Finding(f"sensaxis:{address(pos)}", text, "fehler",
                             T(f"Korrigiert: ={formula}.", f"Corrected: ={formula}."), fix=fix))
    return found


def _text_near(sheet, table: Table) -> str:
    """Titel und Ecke der Tabelle: die Texte in den zwei Zeilen über den Köpfen und in der Ecke."""
    parts = []
    for r in range(table.head_row - 2, table.head_row + 1):
        # Der Titel steht oft ganz links über der Tabelle („8. SENSITIVITÄT – Wert je Aktie …“).
        for c in range(1 if r < table.head_row else table.head_col, table.cols[-1] + 1):
            cell = sheet.cells.get((r, c))
            if cell is not None and not cell.formula and cell.value and _number(cell) is None:
                parts.append(cell.value)
    return " ".join(parts)


def _input(sheet, heads: list[Pos], title: set[str], inside: set[Pos]) -> tuple[Pos, Pos] | None:
    """Die Annahme des Modells, um die eine Achse variiert, und der Kopf, der genau sie ist."""
    from .analyse import _left_label

    # Köpfe in Stufen (=B148-$C$43, in der Mitte =$C$131): Die Annahme liest der Anker, von dem die anderen ausgehen.
    own = set(heads)
    chained = {pos: {p for p in formel.references(sheet.nodes[pos], sheet.name)[0] if p in own}
               for pos in heads if sheet.nodes.get(pos) is not None}
    anchors = [pos for pos, v in chained.items() if not v]
    if len(chained) == len(heads) and len(anchors) == 1 and sum(1 for v in chained.values() if v) == len(heads) - 1:
        anchor = anchors[0]
        refs = {p for p in formel.references(sheet.nodes[anchor], sheet.name)[0] if p not in inside}
        want = sheet.inputs.get(refs and next(iter(refs)))
        if len(refs) == 1 and isinstance(want, float) and not isinstance(want, bool):
            return next(iter(refs)), anchor
    singles: set[Pos] | None = None
    for pos in heads:
        node = sheet.nodes.get(pos)
        if node is None:
            continue
        refs, areas, _ = formel.references(node, sheet.name)
        refs = {p for p in refs if p not in inside}
        singles = refs if singles is None else singles & refs
    values = {p: sheet.inputs.get(p) for p in heads}
    if singles:
        if len(singles) != 1:
            return None
        source = next(iter(singles))
    else:
        # Eingetippte Köpfe: die Annahme mit einem dieser Werte, deren Name im Titel steht (WACC, Wachstum).
        found = []
        for pos, value in sheet.inputs.items():
            if pos in sheet.nodes or pos in inside or not isinstance(value, float) or isinstance(value, bool):
                continue
            if not any(isinstance(v, float) and v == value for v in values.values()):
                continue
            if not sheet.dependents.get(pos) or not (_words(_left_label(sheet.cells, *pos) or "") & title):
                continue
            found.append(pos)
        if len(found) != 1:
            return None
        source = found[0]
    want = sheet.inputs.get(source)
    if not isinstance(want, float) or isinstance(want, bool):
        return None
    base = [p for p in heads if isinstance(values[p], float) and abs(values[p] - want) <= 1e-12 * max(1.0, abs(want))]
    if len(base) != 1:
        return None
    return source, base[0]


def _downstream(sheet, start: Pos) -> set[Pos]:
    seen, todo = set(), [start]
    while todo:
        pos = todo.pop()
        for after in sheet.dependents.get(pos, ()):
            if after not in seen:
                seen.add(after)
                todo.append(after)
    return seen


def _check(sheet, table: Table, findings: list) -> tuple | None:
    from .analyse import _left_label, address
    from .bereiche import _num

    inside = table.body() | table.heads()
    if any(p in sheet.unknown for p in table.body()):
        return None
    near = _text_near(sheet, table)
    title = _words(near)
    across = _input(sheet, [(r, table.head_col) for r in table.rows], title, inside)
    down = _input(sheet, [(table.head_row, c) for c in table.cols], title, inside)
    if across is None or down is None or across[0] == down[0]:
        return None
    (row_input, row_head), (col_input, col_head) = across, down
    center = (row_head[0], col_head[1])
    if center in table.odd:
        return None
    # Das Ergebnis des Modells, das die Tabelle nachrechnet: hängt an beiden Annahmen, und sein Name steht im Titel.
    both = (_downstream(sheet, row_input) & _downstream(sheet, col_input)) - inside
    scored = []
    for pos in both:
        words = _words(_left_label(sheet.cells, *pos) or "")
        if words and words <= title and isinstance(sheet.inputs.get(pos), float):
            scored.append((len(words), pos))
    if not scored:
        return None
    scored.sort(reverse=True)
    if len(scored) > 1 and scored[1][0] == scored[0][0]:
        return None
    output = scored[0][1]
    now, model = sheet.inputs.get(center), sheet.inputs.get(output)
    if not all(isinstance(v, float) and not isinstance(v, bool) and math.isfinite(v) for v in (now, model)):
        return None
    if _close(now, model):
        return None
    row_name = _left_label(sheet.cells, *row_input) or address(row_input)
    col_name = _left_label(sheet.cells, *col_input) or address(col_input)
    label = sheet.label(output)
    text = T(f"{address(center)} in der Sensitivität rechnet mit {row_name} aus {address(row_head)} und {col_name} aus "
             f"{address(col_head)}, genau den Annahmen in {address(row_input)} und {address(col_input)}. Dort muss die "
             f"Tabelle dasselbe zeigen wie {label}. {address(center)} zeigt aber {_num(now)}, "
             f"{address(output)} {_num(model)}.",
             f"{address(center)} in the sensitivity table calculates with {row_name} from {address(row_head)} and "
             f"{col_name} from {address(col_head)}, exactly the assumptions in {address(row_input)} and "
             f"{address(col_input)}. There the table must show the same as {label}. But {address(center)} shows "
             f"{_num(now)} and {address(output)} shows {_num(model)}.")
    fixed = _with_fixes(sheet, findings, [center, output])
    if fixed is not None:
        names, (after_center, after_model) = fixed
        if names and isinstance(after_center, float) and isinstance(after_model, float) and _close(after_center, after_model):
            one = len(names) == 1
            listed = names[0] if one else ", ".join(names[:-1]) + T(f" und {names[-1]}", f" and {names[-1]}")
            same = _num(after_model)
            text += T(f" Mit {'der Korrektur' if len(names) == 1 else 'den Korrekturen'} in {listed} zeigen beide "
                      f"{same}: Der Unterschied kommt von {'diesem Fehler' if len(names) == 1 else 'diesen Fehlern'}.",
                      f" With the {'fix' if one else 'fixes'} in {listed}, both show {same}: the difference comes from "
                      f"{'this error' if one else 'these errors'}.")
            return (f"sensitivity:{address(center)}", text, "pruefen",
                    T("Setz die Korrekturen ein, dann stimmen Modell und Sensitivität überein.",
                      "Apply the fixes, then the model and the sensitivity table agree."))
    split = _diverge(sheet, table, center, output, {row_head: row_input, col_head: col_input})
    if split is not None:
        where, formula, want, fix = split
        from .bereiche import _german_formula

        cell = sheet.cells[where]
        model_text = f"={_german_formula(cell.formula)}" if cell.formula else _num(sheet.inputs.get(where))
        has, other = _num(sheet.inputs.get(where)), f"={_german_formula(formula)}"
        text += T(f" Die beiden Rechnungen trennen sich in {sheet.label(where)}: Das Modell rechnet dort "
                  f"{model_text} mit {has}, die Sensitivität an dieser Stelle "
                  f"{other} mit {_num(want)}. Eine von beiden ist falsch.",
                  f" The two calculations part ways in {sheet.label(where)}: there the model calculates {model_text} "
                  f"and gets {has}; the sensitivity table calculates {other} at this point and gets "
                  f"{_num(want)}. One of the two is wrong.")
        return (f"sensitivity:{address(center)}", text, "fehler",
                T(f"Stimmt die Rechnung der Sensitivität, gehört in {address(where)} {other}. Dann "
                  f"zeigt {address(output)} auch {_num(now)}.",
                  f"If the sensitivity table's calculation is right, {address(where)} should contain {other}. Then "
                  f"{address(output)} also shows {_num(now)}."), fix)
    text += T(" Eine der beiden Rechnungen stimmt also nicht.", " So one of the two calculations is wrong.")
    return (f"sensitivity:{address(center)}", text, "fehler",
            T(f"Vergleiche Schritt für Schritt, wie {address(center)} und {address(output)} rechnen.",
              f"Compare step by step how {address(center)} and {address(output)} calculate."))


def _close(a: float, b: float) -> bool:
    return abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))


def _with_fixes(sheet, findings: list, targets: list[Pos]) -> tuple[list[str], list] | None:
    """Die Werte von targets, wenn alle Korrekturen der übrigen Hinweise eingesetzt sind, und die Zellen der
    Korrekturen, die dafür zählen (die in der Kette von targets)."""
    from .analyse import _chain

    if len(sheet.nodes) > MAX_FORMULAS:
        return None
    changes: dict[Pos, formel.Node] = {}
    owner: dict[Pos, str] = {}
    for finding in findings:
        fix = getattr(finding, "fix", None) or {}
        for where, formula in (fix.get("cells") or {}).items():
            if not isinstance(formula, str) or not formula.startswith("="):
                continue
            try:
                changes[formel.position(where)] = formel.parse(formula[1:])
            except (formel.Unknown, AttributeError):
                continue
            owner[formel.position(where)] = fix.get("range") or where
    if not changes:
        return None
    needed = set()
    for target in targets:
        needed |= _chain(sheet, target)
    used = sorted(p for p in changes if p in needed)
    if not used:
        return None
    saved = {p: sheet.nodes.get(p, _MISSING) for p in used}
    try:
        for p in used:
            sheet.nodes[p] = changes[p]
        _, value = sheet.evaluator(chain=set(sheet.nodes))
        try:
            values = [value(t) for t in targets]
        except (formel.Unknown, RecursionError, ValueError, TypeError, ZeroDivisionError):
            return None
    finally:
        for p, node in saved.items():
            if node is _MISSING:
                sheet.nodes.pop(p, None)
            else:
                sheet.nodes[p] = node
    return list(dict.fromkeys(owner[p] for p in used)), values


_MISSING = object()


def _inputs(sheet, pos: Pos, cache: dict) -> frozenset:
    """Die Eingaben (Zellen ohne Formel), von denen eine Zelle abhängt."""
    from .analyse import _chain

    if pos not in cache:
        cache[pos] = frozenset(p for p in _chain(sheet, pos) if p not in sheet.nodes and sheet.shown.get(p) is not None) \
            if pos in sheet.nodes else frozenset({pos})
    return cache[pos]


def _subtrees(node: formel.Node):
    yield node
    for child in node.args:
        yield from _subtrees(child)


def _diverge(sheet, table: Table, center: Pos, output: Pos, heads: dict[Pos, Pos]):
    """Wo sich Modell und Sensitivität trennen: Jeder Teil der Formel im Mittelpunkt (die Summe der Barwerte, der
    Terminalwert …) hat seinen Wert. Die erste Zelle des Modells, die zu keinem davon passt und mit dem passenden Wert
    das Ergebnis der Sensitivität ergäbe, ist die Stelle. Gibt (Zelle, Formel der Sensitivität dafür, ihr Wert,
    Korrektur) oder None."""
    from . import einsetzen
    from .analyse import _chain, _run
    from .bereiche import _kept, _text

    node = sheet.nodes.get(center)
    if node is None:
        return None
    calc, value = sheet.evaluator()
    cache: dict = {}
    inside = table.body() | table.heads()

    def ins(n: formel.Node) -> frozenset:
        singles, areas, _ = formel.references(n, sheet.name)
        out: set = set()
        for p in singles:
            out |= {heads[p]} if p in heads else _inputs(sheet, p, cache)
        for r1, c1, r2, c2 in areas:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if (r, c) in sheet.cells:
                        out |= _inputs(sheet, (r, c), cache)
        return frozenset(out)

    parts = []
    for n in _subtrees(node):
        if n.kind not in ("op", "call", "neg", "pct"):
            continue
        try:
            v = calc.run(n)
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError):
            continue
        if isinstance(v, float) and not isinstance(v, bool) and math.isfinite(v):
            parts.append((n, v, ins(n)))
    chain = _chain(sheet, output)
    model = [p for p in chain if p in sheet.nodes and p not in inside and p not in sheet.unknown
             and isinstance(sheet.inputs.get(p), float)]
    direct = formel.references(node, sheet.name)[0]

    def matches(p: Pos) -> bool:
        return p in direct or any(_close(sheet.inputs[p], v) and _inputs(sheet, p, cache) == i for _, v, i in parts)

    open_ = sorted((p for p in model if not matches(p)), key=lambda p: (len(_chain(sheet, p)), p))
    for where in open_:
        for part, want, _ in parts:
            if _close(want, sheet.inputs[where]):
                continue
            after = _run(sheet, output, chain, overrides={where: want})
            if not (isinstance(after, float) and _close(after, sheet.inputs.get(center))):
                continue
            # Ein Teil der Sensitivität, den das Modell schon ausrechnet (der Barwert des Terminalwerts in B38),
            # heißt in der Korrektur wie dort, auch wenn ihn gerade keine Formel bis zum Ergebnis liest.
            everyone = [p for p in sheet.nodes if p not in inside and p not in sheet.unknown
                        and isinstance(sheet.inputs.get(p), float) and not isinstance(sheet.inputs.get(p), bool)]
            later = {p for p in everyone if where in _chain(sheet, p)}
            kept = _kept(sheet.cells[center].formula)
            for head, source in heads.items():
                spelled = _spelled(sheet, head, source)
                kept[("ref", source)] = spelled
            used = [p for p in everyone if p != where and p not in later and sheet.cells[p].formula]

            def back(n: formel.Node) -> formel.Node:
                if n.kind in ("op", "call", "neg", "pct") and n is not part:
                    try:
                        v = calc.run(n)
                    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError):
                        v = None
                    if isinstance(v, float) and not isinstance(v, bool):
                        same = [p for p in used if _close(sheet.inputs[p], v) and _inputs(sheet, p, cache) == ins(n)]
                        if same:
                            return formel.Node("ref", min(same, key=lambda p: (len(_chain(sheet, p)) < 2, p)))
                if n.kind == "ref" and n.value in heads:
                    return formel.Node("ref", heads[n.value])
                if n.args:
                    return formel.Node(n.kind, n.value, tuple(back(a) for a in n.args))
                return n

            whole = [p for p in used if _close(sheet.inputs[p], want) and _inputs(sheet, p, cache) == ins(part)]
            rebuilt = formel.Node("ref", whole[0]) if whole else back(part)
            text = _text(rebuilt, kept)
            if text is None:
                continue
            target = _copy_source(sheet, where, text)
            if target is not None:
                where, text = target
            try:
                check = calc.run(formel.parse(text))
            except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError):
                continue
            if not (isinstance(check, float) and _close(check, want if target is None else check)):
                continue
            fix = einsetzen.formulas(sheet, {where: text})
            return where, text, (check if target is not None else want), fix
    return None


def _spelled(sheet, head: Pos, source: Pos) -> str:
    """Wie der Kopf die Annahme schreibt ($B$15), sonst ihre Adresse."""
    from .analyse import address

    formula = sheet.cells[head].formula or ""
    for match in re.finditer(r"\$?[A-Z]{1,3}\$?\d+", formula):
        if formel.position(match.group(0).replace("$", "")) == source:
            return match.group(0)
    return address(source)


def _copy_source(sheet, where: Pos, text: str):
    """Rechnet die Zelle wie die Sensitivität, nur mit einer Kopie (B42 =B37) statt der richtigen Zelle (B38), gehört
    die Korrektur in die Kopie: B42 =B38."""
    from .analyse import address
    from .zeilenmodell import _refs, _skeleton

    formula = sheet.cells[where].formula
    if not formula:
        return None
    have = {(r, c) for _, c, _, r in _refs(formula)}
    want = {(r, c) for _, c, _, r in _refs(text)}
    old, new = have - want, want - have
    if len(old) != 1 or len(new) != 1 or _skeleton(formula) != _skeleton(text):
        return None
    copy, right = next(iter(old)), next(iter(new))
    node = sheet.nodes.get(copy)
    if node is None or node.kind != "ref":
        return None
    return copy, address(right)
