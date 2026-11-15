"""Formeln nebeneinander, die eine Spalte in gleich großen Blöcken zusammenfassen: Q1 =SUMME(Daten!B2:B4),
Q2 =SUMME(Daten!B5:B7), Q3 =SUMME(Daten!B8:B10). Ihre Bezüge sehen verschieden aus, darum findet sie der Vergleich der
Formelmuster nicht. Liegen alle Blöcke bis auf einen ohne Lücke hintereinander und gleich groß, ist der eine falsch:
Q2 =SUMME(Daten!B5:B6) lässt den Juni in jedem Quartal weg, =SUMME(Daten!B4:B7) zählt den März zweimal."""

from __future__ import annotations

import re

from . import formel
from .analyse import Finding, Sheet, _fmt, address
from .bereiche import _german_formula, _kept, _text
from .sprache import T
from .xlsx import column_letters

Pos = tuple[int, int]
_REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d{1,7})")

# Was einen Block zusammenfasst. Nachschlagen oder Bedingungen gehören nicht dazu.
_BLOCKS = {"SUM", "AVERAGE", "COUNT", "COUNTA", "MIN", "MAX", "MEDIAN"}
MIN_BLOCKS = 3
MAX_FINDINGS = 3


def _block(node) -> tuple[str | None, bool, int, int, int] | None:
    """=SUMME(Daten!B5:B7) → (Blatt, senkrecht, Spalte, von, bis); None für alles andere."""
    if node is None or node.kind != "call" or node.value not in _BLOCKS or len(node.args) != 1:
        return None
    arg = node.args[0]
    if arg.kind == "range":
        where, (r1, c1, r2, c2) = None, arg.value
    elif arg.kind == "xrange":
        where, (r1, c1, r2, c2) = arg.value[0].casefold(), arg.value[1]
    else:
        return None
    if r2 >= formel.MAX_ROW or c2 >= formel.MAX_COL:
        return None
    if c1 == c2 and r1 < r2:
        return where, True, c1, r1, r2
    if r1 == r2 and c1 < c2:
        return where, False, r1, c1, c2
    return None


def _runs(sheet: Sheet) -> list[list[tuple[Pos, tuple]]]:
    """Mindestens drei solche Formeln direkt nebeneinander (oder untereinander), dieselbe Funktion über dieselbe Spalte."""
    blocks = {}
    for pos, node in sheet.nodes.items():
        if pos in sheet.unknown:
            continue
        found = _block(node)
        if found is not None:
            blocks[pos] = (node.value, *found)
    runs = []
    for along in (0, 1):  # 0: in einer Zeile nebeneinander, 1: in einer Spalte untereinander
        lines: dict = {}
        for pos, key in blocks.items():
            lines.setdefault((key[:4], pos[along]), []).append(pos)
        for _, cells in sorted(lines.items(), key=lambda kv: (kv[0][1], str(kv[0][0]))):
            cells.sort()
            run = [cells[0]]
            for pos in cells[1:]:
                if pos[1 - along] == run[-1][1 - along] + 1:
                    run.append(pos)
                    continue
                if len(run) >= MIN_BLOCKS:
                    runs.append([(p, blocks[p]) for p in run])
                run = [pos]
            if len(run) >= MIN_BLOCKS:
                runs.append([(p, blocks[p]) for p in run])
    return runs


def _odd(spans: list[tuple[int, int]]) -> tuple[int, tuple[int, int]] | None:
    """Der eine Block, ohne den alle gleich groß ohne Lücke hintereinander liegen, und wo er liegen sollte."""
    if any(b[0] <= a[0] for a, b in zip(spans, spans[1:])):
        return None
    sizes = [e - s + 1 for s, e in spans]
    n = max(set(sizes), key=sizes.count)
    if n < 2 or sizes.count(n) < len(spans) - 1:
        return None
    hits = []
    for k, (s, e) in enumerate(spans):
        rest = [i for i in range(len(spans)) if i != k]
        if any(sizes[i] != n for i in rest):
            continue
        # Die übrigen liegen ohne Lücke hintereinander (außer um den einen herum).
        if any(spans[i + 1][0] != spans[i][1] + 1 for i in range(len(spans) - 1) if k not in (i, i + 1)):
            continue
        if k == 0:
            want = (spans[1][0] - n, spans[1][0] - 1)
        elif k == len(spans) - 1:
            want = (spans[k - 1][1] + 1, spans[k - 1][1] + n)
        else:
            want = (spans[k - 1][1] + 1, spans[k + 1][0] - 1)
        if want[1] - want[0] + 1 != n or want == (s, e) or want[0] > e or want[1] < s or want[0] < 1:
            continue
        hits.append((k, want))
    return hits[0] if len(hits) == 1 else None


def _cells(down: bool, line: int, span: tuple[int, int]) -> list[Pos]:
    return [(i, line) if down else (line, i) for i in range(span[0], span[1] + 1)]


def _source(sheet: Sheet, where: str | None) -> dict:
    return sheet.cells if where is None else sheet.book.get(where, {})


def _value(sheet: Sheet, where: str | None, pos: Pos):
    try:
        return sheet.inputs.get(pos) if where is None else sheet.other(where, pos)
    except formel.Unknown:
        return None


def _number(value) -> bool:
    return isinstance(value, float) and not isinstance(value, bool)


def _name(sheet: Sheet, where: str | None, pos: Pos, down: bool) -> str:
    """Wie die ausgelassene Zeile heißt: der Text am Anfang der Zeile (oder oben in der Spalte), zum Beispiel „Jun“."""
    cells = _source(sheet, where)
    r, c = pos
    for i in range(1, (c if down else r)):
        cell = cells.get((r, i) if down else (i, c))
        if cell is not None and cell.formula is None and cell.value.strip() and not _number(_value(sheet, where, (r, i) if down else (i, c))):
            return cell.value.strip()
    return ""


def _area(shown: str | None, down: bool, line: int, span: tuple[int, int]) -> str:
    cells = _cells(down, line, span)
    text = f"{address(cells[0])}:{address(cells[-1])}"
    return text if shown is None else f"{shown}!{text}"


def _restyled(text: str, first: Pos, last: Pos) -> str:
    """Daten!$B$5:$B$6 → Daten!$B$5:$B$7: derselbe Bezug mit denselben $, nur die neuen Zellen."""
    prefix, bang, ref = text.rpartition("!")
    targets = iter((first, last))

    def put(match: re.Match) -> str:
        r, c = next(targets)
        return f"{match.group(1)}{column_letters(c)}{match.group(3)}{r}"

    return f"{prefix}{bang}{_REF.sub(put, ref, count=2)}"


def checks(sheet: Sheet) -> list[Finding]:
    found: list[Finding] = []
    calc, _ = sheet.evaluator()
    for run in _runs(sheet):
        _, (function, where, down, line, _, _) = run[0]
        spans = [(key[4], key[5]) for _, key in run]
        odd = _odd(spans)
        if odd is None:
            continue
        k, want = odd
        pos = run[k][0]
        cells = _source(sheet, where)
        now, then = set(_cells(down, line, spans[k])), set(_cells(down, line, want))
        missing, doubled = sorted(then - now), sorted(now - then)
        # Was hinzukommt, sind Zahlen wie die übrigen der Blöcke, keine Summe oder Überschrift.
        if not all(_number(_value(sheet, where, p)) for p in then):
            continue
        typed = [p for _, key in run for p in _cells(down, line, (key[4], key[5]))]
        formulas = {p for p in typed if cells.get(p) is not None and cells[p].formula is not None}
        if any((p in formulas) != (len(formulas) * 2 > len(typed)) for p in missing):
            continue
        node = sheet.nodes[pos]
        old = node.args[0]
        shown = old.value[0] if old.kind == "xrange" else None
        kept = _kept(sheet.cells[pos].formula or "")
        first, last = _cells(down, line, want)[0], _cells(down, line, want)[-1]
        area = (first[0], first[1], last[0], last[1])
        arg = formel.Node("range", area) if shown is None else formel.Node("xrange", (shown, area))
        if (old.kind, old.value) in kept:
            kept[(arg.kind, arg.value)] = _restyled(kept[(old.kind, old.value)], first, last)
        english = _text(formel.Node("call", node.value, (arg,)), kept)
        if not english:
            continue
        try:
            got, right = calc.run(node), calc.run(formel.parse(english))
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
            continue
        if not (_number(got) and _number(right)):
            continue
        n = want[1] - want[0] + 1
        others = [i for i in range(len(run)) if i != k][:2]
        examples = ", ".join(f"{address(run[i][0])} {_area(shown, down, line, spans[i])}" for i in others)
        unit = T("Zeilen", "rows") if down else T("Spalten", "columns")
        what = []
        at = (lambda p: f"{shown}!{address(p)}") if shown is not None else address  # noqa: E731
        if missing:
            names = [f"{at(p)}{f' ({nm}, {_fmt(_value(sheet, where, p))})' if (nm := _name(sheet, where, p, down)) else ''}"
                     for p in missing]
            what.append(T(f"{', '.join(names)} {'zählt' if len(missing) == 1 else 'zählen'} dadurch in keinem Block.",
                          f"{', '.join(names)} therefore {'is' if len(missing) == 1 else 'are'} not counted in any "
                          f"block."))
        if doubled:
            other_blocks = {p for i, (_, key) in enumerate(run) if i != k for p in _cells(down, line, (key[4], key[5]))}
            twice = [p for p in doubled if p in other_blocks]
            if twice:
                cells_twice = ", ".join(at(p) for p in twice)
                what.append(T(f"{cells_twice} {'zählt' if len(twice) == 1 else 'zählen'} dadurch "
                              "zweimal, auch im Block daneben.",
                              f"{cells_twice} therefore {'is' if len(twice) == 1 else 'are'} counted twice, also in the "
                              "block next to it."))
        if not what:
            continue
        taken = _area(shown, down, line, spans[k])
        lead = T(f"{sheet.figure_label(pos)} nimmt {taken}. Die Formeln daneben nehmen je "
                 f"{n} {unit} ohne Lücke hintereinander ({examples}). {' '.join(what)}",
                 f"{sheet.figure_label(pos)} takes {taken}. The formulas next to it each take {n} {unit} in a row "
                 f"without gaps ({examples}). {' '.join(what)}")
        same = abs(got - right) <= 1e-9 * max(1.0, abs(right))
        if same:
            text = T(f"{lead} Heute ändert das am Ergebnis zufällig nichts. Sobald sich die Zahlen ändern, stimmt "
                     f"{address(pos)} nicht mehr.",
                     f"{lead} By chance this doesn't change the result today. As soon as the numbers change, "
                     f"{address(pos)} will be wrong.")
            level = "pruefen"
        else:
            text = T(f"{lead} Richtig wäre {_fmt(right)} statt {_fmt(got)}.",
                     f"{lead} Correct would be {_fmt(right)} instead of {_fmt(got)}.")
            level = "fehler"
        german = f"Korrigiert: ={_german_formula(english)}, nachgerechnet {_fmt(right)}."
        # Die Korrektur zum Einsetzen liest einsetzen.fix aus dem deutschen Tipp; hier schon, damit sie in jeder
        # Sprache dieselbe ist.
        from . import einsetzen

        fix = einsetzen.fix(sheet, Finding(f"blocks:{address(pos)}", text, level, german))
        tip = T(german, f"Corrected: ={_german_formula(english)}, recalculated {_fmt(right)}.")
        found.append(Finding(f"blocks:{address(pos)}", text, level, tip, fix=fix))
        if len(found) >= MAX_FINDINGS:
            break
    return found

