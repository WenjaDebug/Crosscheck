"""Statistik, die über Werte rechnet, die keine sind, oder etwas anderes, als ihr Name sagt.

- Codes für „fehlt“ (999, -99) im Bereich von MITTELWERT, MEDIAN, STABW und Co.
- Nullen aus WENNFEHLER(…;0) oder WENN(A2="";0;…), die MITTELWERT als echte Werte mitzählt.
- SUMME geteilt durch ANZAHL2, ZEILEN oder die Zahl der Zeilen: Leere Zellen und Text zählen dann als 0.
  MITTELWERTA rechnet Text ebenso als 0.
- STABW.N, wo „Stichprobe“ steht (und umgekehrt), die Varianz unter „Standardabweichung“, ein Standardfehler ohne
  Wurzel aus n und ein Variationskoeffizient, der falsch herum teilt. Eine Standardabweichung, die anders rechnet als
  ihre Nachbarn (STABW.N zwischen lauter STABW.S).
- Anteile, die durch die falsche Summe teilen, oder Summe durch Teil statt Teil durch Summe.

Jede Korrektur ist auf dem Blatt nachgerechnet: Sie muss genau ergeben, was herauskommt, wenn die Codes und
Nullen leer wären oder die Funktion die richtige ist. Alles ohne KI.
"""

from __future__ import annotations

import math
import statistics
from bisect import bisect_left, bisect_right

from . import formel
from .analyse import Finding, Sheet, _fmt, address
from .auswertung import CODES, codes_in
from .bereiche import NO_LETTERS, _german_formula, _kept, _paths, _put, _text, _words
from .formel import Node
from .kopf import _fix_from_tip, _tip_pair, _tip_plus
from .sprache import T
from .workbook import FUNCTIONS, pattern
from .xlsx import column_letters, number_text

Pos = tuple[int, int]

MAX_FINDINGS = 6
_PLAIN = {"AVERAGE", "MEDIAN", "SUM", "MIN", "MAX", "COUNT", "STDEV", "STDEV.S", "STDEV.P", "STDEVP", "VAR", "VAR.S",
          "VAR.P", "VARP"}
_SPREAD = {"STDEV": ("sd", "s"), "STDEV.S": ("sd", "s"), "STDEV.P": ("sd", "p"), "STDEVP": ("sd", "p"),
           "VAR": ("var", "s"), "VAR.S": ("var", "s"), "VAR.P": ("var", "p"), "VARP": ("var", "p")}
_NAMED = {("sd", "s"): "STDEV.S", ("sd", "p"): "STDEV.P", ("var", "s"): "VAR.S", ("var", "p"): "VAR.P"}
# Wie Zellen mit Nullen aussehen, die eigentlich „fehlt“ heißen: der Test einer WENN-Formel.
_BLANK_TESTS = {"ISBLANK", "ISERROR", "ISERR", "ISNA", "ISNUMBER", "ISTEXT"}
_SAMPLE = {"stichprobe", "stichproben", "sample", "empirisch", "empirische"}
_POPULATION = {"grundgesamtheit", "population", "gesamtheit", "vollerhebung"}
_VARIANCE = {"varianz", "variance"}
_DEVIATION = {"standardabweichung", "stabw", "sd", "std", "stdev", "stddev"}
_ERROR = {"standardfehler", "sem", "se", "stderr"}
_VARIATION = {"variationskoeffizient", "vk", "cv", "variationskoeff"}
_SHARE = {"anteil", "anteile", "prozent", "%", "share", "percent", "prozentual"}


def _german(name: str) -> str:
    return FUNCTIONS.get(name, name)


def _fn(name: str) -> str:
    """Der Name einer Funktion im Satz: MITTELWERT auf Deutsch, AVERAGE auf Englisch."""
    return T(_german(name), name)


def _times(k: int) -> str:
    return "once" if k == 1 else f"{k} times"


def _number(value) -> bool:
    return isinstance(value, float) and not isinstance(value, bool) and math.isfinite(value)


def _close(a: float, b: float) -> bool:
    return abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))


def _cells(cells: list[Pos]) -> str:
    shown = [address(p) for p in cells[:3]]
    if len(cells) > 3:
        shown.append(T(f"{len(cells) - 3} weitere", f"{len(cells) - 3} more"))
    return shown[0] if len(shown) == 1 else f"{', '.join(shown[:-1])}{T(' und ', ' and ')}{shown[-1]}"


def _label(cell) -> bool:
    """Ein Name für eine Kennzahl: Text mit Buchstaben oder einem %-Zeichen, auch lang („Standardabweichung der …“)."""
    return (cell is not None and cell.formula is None and isinstance(cell.value, str) and len(cell.value) <= 120
            and ("%" in cell.value or not NO_LETTERS.fullmatch(cell.value) and any(ch.isalpha() for ch in cell.value)))


def _above(sheet: Sheet, row: int, c: int) -> tuple[int, str] | None:
    """Der nächste Name über row in Spalte c (Zeile, Text), ohne Lücke dazwischen. Gemerkt, damit eine lange Spalte
    voller Formeln nicht für jede Zelle bis nach oben gelesen wird."""
    memo = sheet.memo.setdefault("statistik.above", {})
    walked = []
    found = None
    while row >= 1:
        if (row, c) in memo:
            found = memo[(row, c)]
            break
        cell = sheet.cells.get((row, c))
        if cell is None or cell.value == "" and cell.formula is None:
            break
        walked.append(row)
        if _label(cell):
            found = (row, cell.value)
            break
        row -= 1
    for row in walked:
        memo[(row, c)] = found
    return found


def _name(sheet: Sheet, pos: Pos) -> frozenset[str]:
    """Die Wörter, die eine Kennzahl beschreiben: links daneben und darüber (Kopf der Spalte), dazu „%“."""
    memo = sheet.memo.setdefault("statistik.name", {})
    found = memo.get(pos)
    if found is None:
        # Eine Spalte voller Formeln hat überall dieselben Wörter: einmal im Speicher.
        same = sheet.memo.setdefault("statistik.names", {})
        found = frozenset(_named(sheet, pos))
        found = memo[pos] = same.setdefault(found, found)
    return found


def _named(sheet: Sheet, pos: Pos) -> set[str]:
    r, c = pos
    texts = []
    for left in range(c - 1, max(0, c - 3), -1):
        cell = sheet.cells.get((r, left))
        if _label(cell):
            texts.append(cell.value)
            break
    found = _above(sheet, r - 1, c)
    if found is not None and r - found[0] <= 2000:
        texts.append(found[1])
    if sheet.plan.in_table(r, c):
        texts.append(sheet.plan.headers.get(c, ""))
    words = {w for t in texts for w in _words(t)}
    if any("%" in t for t in texts):
        words.add("%")
    return words


class _Calc:
    """Der Formelrechner des Blatts und die Werte, die Excel zeigt, mit dem Text der Formel jeder Zelle."""

    def __init__(self, sheet: Sheet):
        self.sheet = sheet
        self.calc, _ = sheet.evaluator()

    def run(self, node: Node):
        try:
            return self.calc.run(node)
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError):
            return None

    def english(self, text: str):
        try:
            return self.calc.run(formel.parse(text))
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError):
            return None

    def after(self, pos: Pos, overrides: dict, chain: set[Pos] | None = None):
        """Was die Formel in pos ergäbe, wenn diese Zellen anders wären (None: leer)."""
        node = self.sheet.nodes.get(pos)
        calc, _ = self.sheet.evaluator(overrides=overrides, chain=chain)
        try:
            return calc.run(node)
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError):
            return None

    def english_after(self, text: str, overrides: dict):
        calc, _ = self.sheet.evaluator(overrides=overrides)
        try:
            return calc.run(formel.parse(text))
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError):
            return None

    def values(self, area: Node) -> dict[Pos, object]:
        """Die Zellen eines Bereichs auf diesem Blatt und ihre Werte."""
        r1, c1, r2, c2 = area.value
        rows, cols = self.sheet.extent(None)
        out = {}
        for r in range(r1, min(r2, rows) + 1):
            for c in range(c1, min(c2, cols) + 1):
                try:
                    out[(r, c)] = self.calc.value((r, c))
                except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError):
                    out[(r, c)] = None
        return out


def _marked(sheet: Sheet, kind: str, area: tuple) -> int:
    """Wie viele Zellen im Bereich einen Code wie 999 zeigen (kind „code“) oder als Formel 0 (kind „zero“). Je
    Spalte einmal gesucht: Hunderte Formeln über fast dieselbe Spalte (laufender Mittelwert) gehen sie sonst jedes
    Mal ganz durch."""
    key = ("statistik", "marked", kind)
    if key not in sheet.memo:
        rows: dict[int, list[int]] = {}
        for (r, c), v in sorted(sheet.inputs.items()):
            if _number(v) and (v in CODES if kind == "code" else v == 0 and (r, c) in sheet.nodes):
                rows.setdefault(c, []).append(r)
        sheet.memo[key] = rows
    r1, c1, r2, c2 = area
    marked = sheet.memo[key]
    return sum(bisect_right(marked[c], r2) - bisect_left(marked[c], r1) for c in marked if c1 <= c <= c2)


def _statistics(node: Node):
    """Die Funktionen MITTELWERT, MEDIAN, STABW … über genau einen Bereich dieses Blatts: (Pfad, Aufruf)."""
    # Die meisten Formeln einer großen Tabelle haben keine (schnell, ohne die Pfade).
    if not any(c.value in _PLAIN for c in formel.calls(node)):
        return
    for path, call in _paths(node):
        if call.kind == "call" and call.value in _PLAIN and len(call.args) == 1 and call.args[0].kind == "range":
            yield path, call


def _replaced(sheet: Sheet, pos: Pos, node: Node, path, english: str) -> str | None:
    """Die Formel der Zelle mit einem neuen Teil (englisch, wie Excel ihn speichert), die Bezüge wie vorher."""
    try:
        new = formel.parse(english)
    except formel.Unknown:
        return None
    return _text(_put(node, path, new), _kept(sheet.cells[pos].formula or ""))


def _finding(kind: str, pos: Pos, text: str, level: str, tip: tuple[str, str], sheet: Sheet) -> Finding:
    """tip: (in der Sprache der Anfrage, auf Deutsch). Die Korrektur liest einsetzen.attach später aus dem
    deutschen Tipp; auf Englisch fände es dort nichts, darum dann gleich hier aus dem deutschen."""
    item = Finding(f"{kind}:{address(pos)}", text, level, tip[0])
    _fix_from_tip(sheet, item, tip[1])
    return item


# --- Codes für „fehlt“ im Bereich ---


def _code_fix(name: str, area: str, codes: list[float], blanks: bool) -> tuple[str, bool]:
    """Dieselbe Rechnung ohne die Codes, wie Excel sie kann. Zweiter Wert: braucht FILTER (Excel 2021, Microsoft 365)."""
    shown = [number_text(c).replace(",", ".") for c in codes]
    if name in ("AVERAGE", "SUM", "MIN", "MAX"):
        conditional = {"AVERAGE": "AVERAGEIFS", "SUM": "SUMIFS", "MIN": "MINIFS", "MAX": "MAXIFS"}[name]
        if len(codes) == 1 and name in ("AVERAGE", "SUM"):
            return f'{name}IF({area},"<>{shown[0]}")', False
        return f"{conditional}({area}," + ",".join(f'{area},"<>{c}"' for c in shown) + ")", False
    if name == "COUNT":
        return f"COUNT({area})" + "".join(f"-COUNTIF({area},{c})" for c in shown), False
    keep = "*".join(f"({area}<>{c})" for c in shown) + (f'*({area}<>"")' if blanks else "")
    return f"{name}(FILTER({area},{keep}))", True


def _extended(sheet: Sheet, area: Node) -> tuple[int, int, int, int] | None:
    """Hört der Bereich ein paar Zeilen vor dem Ende seiner Spalte in der Tabelle auf (wie range: in
    auswertung.py), der Bereich bis zur letzten Zahl."""
    from .auswertung import _data_rows

    plan = sheet.plan
    r1, c1, r2, c2 = area.value
    if plan.header_row is None or c1 != c2 or c1 not in plan.headers or r2 >= formel.MAX_ROW:
        return None
    key = ("statistik", "numbers", c1)
    if key not in sheet.memo:
        sheet.memo[key] = [r for r in _data_rows(sheet) if _number(sheet.inputs.get((r, c1)))]
    data = sheet.memo[key]
    if len(data) < 3 or not r1 <= data[0] + 1 or r2 >= data[-1] or r2 < data[0]:
        return None
    left = len(data) - bisect_right(data, r2)
    if not left or left > max(3, len(data) // 5):
        return None
    return r1, c1, data[-1], c1


def _codes(sheet: Sheet, calc: _Calc, done: set[Pos], covered: set[str]) -> list[Finding]:  # noqa: C901
    from .nachschlagen import _restyled

    found = []
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown:
            continue
        for path, call in _statistics(node):
            kept = _kept(sheet.cells[pos].formula or "")
            area = _text(call.args[0], kept)
            longer = _extended(sheet, call.args[0])
            # Ein Code zählt erst, wenn er mindestens zweimal dasteht.
            if _marked(sheet, "code", longer or call.args[0].value) < 2:
                continue
            if longer is not None:
                area = _restyled(area, longer) or area
            values = calc.values(formel.Node("range", longer) if longer else call.args[0])
            numbers = [v for v in values.values() if _number(v)]
            codes = sorted(codes_in(numbers))
            if not codes:
                continue
            where = [p for p, v in values.items() if _number(v) and v in codes]
            now = calc.run(node)
            # Was herauskäme, wenn die Zellen mit dem Code leer wären (und der Bereich bis zum Ende der Spalte ginge):
            # Das muss die Korrektur ergeben.
            plain = _replaced(sheet, pos, node, path, f"{call.value}({area})")
            meant = calc.english_after(plain, {p: None for p in where}) if plain else None
            if not (_number(now) and _number(meant)) or _close(now, meant):
                continue
            blanks = any(v is None or v == "" for v in values.values())
            english, modern = _code_fix(call.value, area, codes, blanks)
            fixed = _replaced(sheet, pos, node, path, english)
            got = calc.english(fixed) if fixed else None
            if not _number(got) or not _close(got, meant):
                continue
            many = [sum(1 for p in where if values[p] == c) for c in codes]
            counts = T(", ".join(f"{n}-mal {_fmt(c)}" for n, c in zip(many, codes)),
                       ", ".join(f"{_fmt(c)} ({_times(n)})" for n, c in zip(many, codes)))
            shown = _text(call.args[0], kept).replace("$", "")
            text = T(f"{sheet.figure_label(pos)} rechnet {_german(call.value)} über {shown}, und darin steht {counts} "
                     f"({_cells(where)}), wohl {'ein Code' if len(codes) == 1 else 'Codes'} für „fehlt“. "
                     f"{_german(call.value)} rechnet damit wie mit echten Werten.",
                     f"{sheet.figure_label(pos)} calculates {call.value} over {shown}, which contains {counts} in "
                     f"{_cells(where)}, probably {'a code' if len(codes) == 1 else 'codes'} for “missing”. "
                     f"{call.value} treats {'it like a real value' if len(codes) == 1 else 'them like real values'}.")
            if longer is not None:
                text += T(f" Dazu hört der Bereich bei {address(call.args[0].value[2:])} auf, in "
                          f"{sheet.plan.headers[longer[1]]} stehen Zahlen aber bis {address(longer[2:])}.",
                          f" Also, the range stops at {address(call.args[0].value[2:])}, but "
                          f"{sheet.plan.headers[longer[1]]} has numbers down to {address(longer[2:])}.")
                covered.add(f"range:{address(pos)}")
            text += T(f" So ergibt {address(pos)} {_fmt(now)} statt {_fmt(meant)}.",
                      f" As a result, {address(pos)} gives {_fmt(now)} instead of {_fmt(meant)}.")
            tip = _tip_pair(f"Korrigiert: ={_german_formula(fixed)}, nachgerechnet {_fmt(meant)}"
                    f"{' (FILTER gibt es ab Excel 2021 und in Microsoft 365)' if modern else ''}. Noch sicherer: Lass die "
                    f"Zellen mit {' und '.join(_fmt(c) for c in codes)} leer, dann rechnen alle Formeln richtig.",
                    f"Corrected: ={_german_formula(fixed)}, recalculated {_fmt(meant)}"
                    f"{' (FILTER is available from Excel 2021 and in Microsoft 365)' if modern else ''}. Even safer: "
                    f"leave the cells with {' and '.join(_fmt(c) for c in codes)} empty, then all formulas calculate "
                    f"correctly.")
            found.append(_finding("statcode", pos, text, "fehler", tip, sheet))
            break
    return found


# --- Nullen, die „fehlt“ heißen ---


def _fallback(sheet: Sheet, calc: _Calc, pos: Pos):
    """Zeigt die Zelle 0, weil WENNFEHLER einen Fehler oder WENN eine leere Zeile durch 0 ersetzt? (Pfad der 0, Art)"""
    node = sheet.nodes.get(pos)
    if node is None or node.kind != "call":
        return None
    if node.value in ("IFERROR", "IFNA") and len(node.args) == 2 and node.args[1].kind == "num" and node.args[1].value == 0:
        error = calc.run(node.args[0])
        if isinstance(error, formel.Err):
            return (1,), error.code
        return None
    if node.value == "IF" and len(node.args) == 3:
        test = node.args[0]
        names = {c.value for c in formel.calls(test)}
        blank = names & _BLANK_TESTS or any(p.kind == "op" and any(a.kind == "str" and a.value == "" for a in p.args)
                                             for _, p in _paths(test))
        if not blank:
            return None
        truth = calc.run(test)
        if not isinstance(truth, bool):
            return None
        branch = 1 if truth else 2
        zero = node.args[branch]
        other = node.args[3 - branch]
        if zero.kind == "num" and zero.value == 0 and other.kind not in ("num", "str"):
            return (branch,), "leer"
    return None


def _zeros(sheet: Sheet, calc: _Calc, done: set[Pos], covered: set[str]) -> list[Finding]:  # noqa: C901 - Ursache, Korrektur, Text
    found = []
    readers = sheet.dependents
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown or pos in done:
            continue
        for path, call in _statistics(node):
            if call.value == "SUM" or not _marked(sheet, "zero", call.args[0].value):
                continue  # Nullen ändern eine Summe nicht.
            values = calc.values(call.args[0])
            zeros = {p: _fallback(sheet, calc, p) for p, v in values.items() if _number(v) and v == 0 and p in sheet.nodes}
            zeros = {p: z for p, z in zeros.items() if z is not None}
            if not zeros:
                continue
            now = calc.run(node)
            meant = calc.after(pos, {p: "" for p in zeros})
            if not (_number(now) and _number(meant)) or _close(now, meant):
                continue
            first = min(zeros)
            zero_path, why = zeros[first]
            cell_node = sheet.nodes[first]
            # Die Korrektur in der Spalte: dieselbe Formel mit "" statt 0, wo die Nullen herkommen.
            fixed_cell = _text(_put(cell_node, zero_path, formel.Node("str", "", ())), {})
            same = [p for p in values if p in sheet.nodes and sheet.cells[p].formula and
                    pattern(sheet.cells[p].formula, p) == pattern(sheet.cells[first].formula, first)]
            # Andere Formeln, die mit diesen Zellen weiterrechnen, dürften mit "" nicht zu #WERT! werden.
            others = {d for p in zeros for d in readers.get(p, ()) if d != pos}
            safe = fixed_cell is not None and all(
                not isinstance(calc.after(d, {p: "" for p in zeros}), formel.Err) for d in others)
            area = _text(call.args[0], _kept(sheet.cells[pos].formula or ""))
            real_zero = any(_number(v) and v == 0 and p not in zeros for p, v in values.items())
            function = _fn(cell_node.value)
            if why != "leer":
                cause = T(f"dort {why} ausrechnet und {function} das durch 0 ersetzt",
                          f"calculates {why} there and {function} replaces it with 0")
            else:
                cause = T(f"{function} dort eine leere Zeile mit 0 füllt", f"{function} fills an empty row with 0 there")
            one = len(zeros) == 1
            source = f"{address(first)}: ={_german_formula(sheet.cells[first].formula or '')}"
            text = T(f"{sheet.figure_label(pos)} rechnet {_german(call.value)} über {area.replace('$', '')}. "
                     f"{len(zeros)} {'Zelle' if one else 'Zellen'} darin ({_cells(sorted(zeros))}) "
                     f"{'zeigt' if one else 'zeigen'} 0, weil die Formel {cause} "
                     f"({source}). {_german(call.value)} "
                     f"zählt diese Nullen als echte Werte. So ergibt {address(pos)} {_fmt(now)} statt {_fmt(meant)}.",
                     f"{sheet.figure_label(pos)} calculates {call.value} over {area.replace('$', '')}. "
                     f"{len(zeros)} {'cell' if one else 'cells'} in it ({_cells(sorted(zeros))}) "
                     f"{'shows' if one else 'show'} 0 because the formula {cause} ({source}). {call.value} counts "
                     f"these zeros as real values. As a result, {address(pos)} gives {_fmt(now)} instead of "
                     f"{_fmt(meant)}.")
            start, end = (min(same), max(same)) if same else (first, first)
            fixed_start = _text(_put(sheet.nodes[start], zero_path, formel.Node("str", "", ())),
                                _kept(sheet.cells[start].formula or "")) if same else None
            if safe and fixed_start and len(same) == len([p for p in values if p in sheet.nodes]):
                tip = _tip_pair(f"Schreib in {address(start)} ={_german_formula(fixed_start)} (\"\" statt 0) und zieh es bis "
                        f"{address(end)}. Dann ergibt {address(pos)} {_fmt(meant)} (nachgerechnet).",
                        f"In {address(start)}, enter ={_german_formula(fixed_start)} (\"\" instead of 0) and fill it "
                        f"down to {address(end)}. Then {address(pos)} gives {_fmt(meant)} (recalculated).")
            elif not real_zero and call.value in ("AVERAGE", "COUNT", "MIN"):
                english = {"AVERAGE": f'AVERAGEIF({area},"<>0")', "COUNT": f"COUNTIF({area},\"<>0\")-COUNTBLANK({area})",
                           "MIN": f'MINIFS({area},{area},"<>0")'}[call.value]
                fixed = _replaced(sheet, pos, node, path, english)
                got = calc.english(fixed) if fixed else None
                if not _number(got) or not _close(got, meant):
                    continue
                tip = _tip_pair(f"Korrigiert: ={_german_formula(fixed)}, nachgerechnet {_fmt(meant)}.",
                        f"Corrected: ={_german_formula(fixed)}, recalculated {_fmt(meant)}.")
            else:
                continue
            found.append(_finding("statzero", pos, text, "fehler", tip, sheet))
            # Die Nullen aus WENNFEHLER sind keine unmöglichen Werte, sondern fehlende.
            covered |= {f"impossible:{address(p)}" for p in zeros}
            break
    return found


# --- Mittelwert über leere Zellen und Text ---


def _divisor(sheet: Sheet, calc: _Calc, node: Node, area: Node) -> tuple[str, float, Node | None] | None:
    """Wodurch die Summe geteilt wird, wenn das Zeilen zählt statt Zahlen: (wie es heißt, Wert, gezählter Bereich).
    Ohne Bereich zählt es jede Zelle (feste Zahl, ZEILEN)."""
    r1, c1, r2, c2 = area.value
    rows, _ = sheet.extent(None)
    size = (min(r2, rows) - r1 + 1) * (c2 - c1 + 1)
    if node.kind == "num" and node.value == size:
        return T(f"die feste Zahl {_fmt(node.value)}", f"the fixed number {_fmt(node.value)}"), node.value, None
    if node.kind == "call" and len(node.args) == 1 and node.args[0].kind == "range":
        o1, oc1, o2, oc2 = node.args[0].value
        if node.value in ("COUNTA", "ROWS") and (o1, min(o2, rows)) == (r1, min(r2, rows)) and oc1 == oc2:
            value = calc.run(node)
            if _number(value):
                said = f"{_fn(node.value)}({_text(node.args[0], {})})"
                return said, value, node.args[0] if node.value == "COUNTA" else None
    return None


def _counted(calc: _Calc, area: Node, counter: Node | None, values: dict) -> list[Pos]:
    """Die Zellen ohne Zahl, die der Teiler trotzdem mitzählt."""
    others = [p for p, v in values.items() if not _number(v)]
    if counter is None:
        return others
    shown = calc.values(counter)
    rows = {p[0] for p, v in shown.items() if v is not None and v != ""} | \
           {p[0] for p, v in shown.items() if isinstance(v, str)}  # ANZAHL2 zählt auch "" aus Formeln
    return [p for p in others if p[0] in rows]


def _blanks(sheet: Sheet, calc: _Calc, done: set[Pos]) -> list[Finding]:  # noqa: C901 - zwei Formen, ein Hinweis
    found = []
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown or pos in done or \
                not any(c.value in ("SUM", "AVERAGEA") for c in formel.calls(node)):
            continue
        for path, part in _paths(node):
            if part.kind == "op" and part.value == "/" and part.args[0].kind == "call" and part.args[0].value == "SUM" and \
                    len(part.args[0].args) == 1 and part.args[0].args[0].kind == "range":
                area = part.args[0].args[0]
                divisor = _divisor(sheet, calc, part.args[1], area)
                if divisor is None:
                    continue
                how, size, counter = divisor
                values = calc.values(area)
                counted = _counted(calc, area, counter, values)
            elif part.kind == "call" and part.value == "AVERAGEA" and len(part.args) == 1 and part.args[0].kind == "range":
                area, how, size = part.args[0], "", None
                values = calc.values(area)
                counted = [p for p, v in values.items() if isinstance(v, str) or isinstance(v, bool)]
            else:
                continue
            numbers = [p for p, v in values.items() if _number(v)]
            texts = [p for p in counted if isinstance(values[p], str) and values[p].strip()]
            if not numbers or not counted or size is not None and len(numbers) + len(counted) != size:
                continue
            text_area = _text(area, _kept(sheet.cells[pos].formula or ""))
            fixed = _replaced(sheet, pos, node, path, f"AVERAGE({text_area})")
            now, got = calc.run(node), calc.english(fixed) if fixed else None
            mean = math.fsum(values[p] for p in numbers) / len(numbers)
            if not (_number(now) and _number(got)) or _close(now, got) or \
                    not _close(calc.english(f"AVERAGE({text_area})") or math.nan, mean):
                continue
            words = []
            if texts:
                words.append(T(f"Text wie „{values[texts[0]].strip()}“ ({_cells(texts)})",
                               f"text such as “{values[texts[0]].strip()}” ({_cells(texts)})"))
            blank = [p for p in counted if p not in texts]
            if blank:
                words.append(T(f"{'eine leere Zelle' if len(blank) == 1 else 'leere Zellen'} ({_cells(blank)})",
                               f"{'an empty cell' if len(blank) == 1 else 'empty cells'} ({_cells(blank)})"))
            plain = text_area.replace("$", "")
            one = len(words) == 1 and (texts or len(blank) == 1)
            if size is None:
                said = T(f"{sheet.figure_label(pos)} rechnet MITTELWERTA über {plain}. MITTELWERTA zählt ",
                         f"{sheet.figure_label(pos)} calculates AVERAGEA over {plain}. AVERAGEA counts ")
            else:
                # Eine feste Zahl sagt ihren Wert schon selbst, ANZAHL2(…) bekommt ihn dazu.
                by = how if part.args[1].kind == "num" else f"{how} = {_fmt(size)}"
                said = T(f"{sheet.figure_label(pos)} teilt die Summe von {plain} durch "
                         f"{by}, aber nur {len(numbers)} Zellen "
                         f"darin enthalten Zahlen. Damit {'zählt' if one else 'zählen'} ",
                         f"{sheet.figure_label(pos)} divides the sum of {plain} by {by}, but only {len(numbers)} "
                         f"{'cell' if len(numbers) == 1 else 'cells'} in it {'contains' if len(numbers) == 1 else 'contain'} numbers. "
                         f"So it counts ")
            text = T(f"{said}{' und '.join(words)} als 0 mit. So ergibt {address(pos)} {_fmt(now)} statt {_fmt(got)}.",
                     f"{said}{' and '.join(words)} as 0. As a result, {address(pos)} gives {_fmt(now)} instead of "
                     f"{_fmt(got)}.")
            tip = _tip_pair(f"Korrigiert: ={_german_formula(fixed)}, nachgerechnet {_fmt(got)}. MITTELWERT zählt nur Zellen mit Zahlen.",
                    f"Corrected: ={_german_formula(fixed)}, recalculated {_fmt(got)}. AVERAGE only counts cells with "
                    f"numbers.")
            if not texts:
                tip = _tip_plus(tip, _tip_pair(" Heißt leer bei dir wirklich 0, ist die Formel richtig.",
                         " If empty really means 0 for you, the formula is right."))
            found.append(_finding("statblank", pos, text, "fehler" if texts else "pruefen", tip, sheet))
            break
    return found


# --- Welche Streuung? ---


def _spread_call(node: Node) -> Node | None:
    while node.kind == "call" and node.value in ("ROUND", "ROUNDUP", "ROUNDDOWN") and node.args:
        node = node.args[0]
    if node.kind == "call" and node.value in _SPREAD and len(node.args) == 1 and node.args[0].kind == "range":
        return node
    return None


def _exact_spread(values: list[float], kind: str, sample: bool) -> float | None:
    if len(values) < (2 if sample else 1):
        return None
    if kind == "sd":
        return statistics.stdev(values) if sample else statistics.pstdev(values)
    return statistics.variance(values) if sample else statistics.pvariance(values)


def _spreads(sheet: Sheet, calc: _Calc, done: set[Pos]) -> list[Finding]:  # noqa: C901 - je Name eine Regel
    found = []
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown or pos in done or not any(c.value in _SPREAD for c in formel.calls(node)):
            continue
        words = _name(sheet, pos)
        if not words & (_SAMPLE | _POPULATION | _VARIANCE | _DEVIATION | _ERROR | _VARIATION):
            continue
        now = calc.run(node)
        if not _number(now):
            continue
        kept = _kept(sheet.cells[pos].formula or "")
        for path, spread in _paths(node):
            if not (spread.kind == "call" and spread.value in _SPREAD and len(spread.args) == 1 and
                    spread.args[0].kind == "range"):
                continue
            kind, variant = _SPREAD[spread.value]
            area = _text(spread.args[0], kept)
            numbers = [v for v in calc.values(spread.args[0]).values() if _number(v)]
            want_kind, want_variant, why = kind, variant, ""
            if words & _VARIANCE and not words & _DEVIATION and kind == "sd":
                want_kind, why = "var", T("die Standardabweichung, nicht die Varianz (die ist ihr Quadrat)",
                                          "the standard deviation, not the variance (which is its square)")
            elif words & _DEVIATION and not words & (_VARIANCE | _ERROR) and kind == "var":
                want_kind, why = "sd", T("die Varianz, nicht die Standardabweichung (die ist ihre Wurzel)",
                                         "the variance, not the standard deviation (which is its square root)")
            if words & _SAMPLE and not words & _POPULATION and variant == "p":
                want_variant = "s"
                why = (why + T(", und sie teilt durch n wie für die ganze Grundgesamtheit, gemeint ist die Stichprobe (n-1)",
                               ", and it divides by n as for the whole population, but the sample (n-1) is meant")
                       if why else T("die Streuung der Grundgesamtheit (geteilt durch n), gemeint ist die Stichprobe (n-1)",
                                     "the spread of the population (divided by n), but the sample (n-1) is meant"))
            elif words & _POPULATION and not words & _SAMPLE and variant == "s":
                want_variant = "p"
                why = (why + T(", und sie teilt durch n-1 wie für eine Stichprobe, gemeint ist die Grundgesamtheit (n)",
                               ", and it divides by n-1 as for a sample, but the population (n) is meant")
                       if why else T("die Streuung einer Stichprobe (geteilt durch n-1), gemeint ist die Grundgesamtheit (n)",
                                     "the spread of a sample (divided by n-1), but the population (n) is meant"))
            if not why:
                continue
            name = _NAMED[(want_kind, want_variant)]
            exact = _exact_spread(numbers, want_kind, want_variant == "s")
            fixed = _replaced(sheet, pos, node, path, f"{name}({area})")
            check = calc.english(f"{name}({area})")
            got = calc.english(fixed) if fixed else None
            if exact is None or not _number(check) or not _close(check, exact) or not _number(got) or _close(got, now):
                continue
            text = T(f"{sheet.figure_label(pos)} rechnet {_german(spread.value)}, das ist {why}. So zeigt {address(pos)} "
                     f"{_fmt(now)} statt {_fmt(got)}.",
                     f"{sheet.figure_label(pos)} calculates {spread.value}, which is {why}. As a result, {address(pos)} "
                     f"shows {_fmt(now)} instead of {_fmt(got)}.")
            found.append(_finding("spread", pos, text, "fehler", _corrected(fixed, got), sheet))
            break
        else:
            hint = _error_or_variation(sheet, calc, pos, node, words, now)
            if hint is not None:
                found.append(hint)
    return found


def _error_or_variation(sheet: Sheet, calc: _Calc, pos: Pos, node: Node, words: set[str], now: float) -> Finding | None:
    """Standardfehler ohne Wurzel aus n (oder mit der falschen Anzahl), Variationskoeffizient falsch herum."""
    kept = _kept(sheet.cells[pos].formula or "")
    bare = node
    while bare.kind == "call" and bare.value in ("ROUND", "ROUNDUP", "ROUNDDOWN") and bare.args:
        bare = bare.args[0]
    path = next(p for p, n in _paths(node) if n is bare)
    if words & _ERROR and not words & (_VARIANCE | _VARIATION):
        spread = _spread_call(bare)
        divided = bare.kind == "op" and bare.value == "/" and _spread_call(bare.args[0]) is not None
        if spread is None and divided:
            spread = _spread_call(bare.args[0])
        if spread is None or _SPREAD[spread.value][0] != "sd":
            return None
        area = _text(spread.args[0], kept)
        numbers = [v for v in calc.values(spread.args[0]).values() if _number(v)]
        if len(numbers) < 2:
            return None
        sd = _exact_spread(numbers, "sd", _SPREAD[spread.value][1] == "s")
        exact = sd / math.sqrt(len(numbers))
        english = f"{_text(spread, kept)}/SQRT(COUNT({area}))"
        fixed = _replaced(sheet, pos, node, path, english)
        got = calc.english(fixed) if fixed else None
        if not _number(got) or not _close(got, exact) or _close(got, now):
            return None
        if divided:
            n = calc.run(bare.args[1])
            why = T(f"teilt die Standardabweichung durch {_fmt(n)}, in {area.replace('$', '')} stehen aber {len(numbers)} Zahlen, "
                    f"also gehört die Wurzel aus {len(numbers)} darunter",
                    f"divides the standard deviation by {_fmt(n)} although {area.replace('$', '')} contains {len(numbers)} "
                    f"numbers, so it should be divided by the square root of {len(numbers)}")
        else:
            why = T(f"rechnet nur die Standardabweichung. Der Standardfehler ist sie geteilt durch die Wurzel aus der Anzahl "
                    f"({len(numbers)} Werte)",
                    f"only calculates the standard deviation. The standard error is the standard deviation divided by "
                    f"the square root of the count ({len(numbers)} values)")
        text = T(f"{sheet.figure_label(pos)} {why}. So zeigt {address(pos)} {_fmt(now)} statt {_fmt(got)}.",
                 f"{sheet.figure_label(pos)} {why}. As a result, {address(pos)} shows {_fmt(now)} instead of {_fmt(got)}.")
        return _finding("spread", pos, text, "fehler", _corrected(fixed, got), sheet)
    if words & _VARIATION and bare.kind == "op" and bare.value == "/" and bare.args[0].kind == "call" and \
            bare.args[0].value == "AVERAGE" and _spread_call(bare.args[1]) is not None:
        english = f"{_text(bare.args[1], kept)}/{_text(bare.args[0], kept)}"
        fixed = _replaced(sheet, pos, node, path, english)
        got = calc.english(fixed) if fixed else None
        if not _number(got) or not _number(now) or not _close(got * now, 1.0) or _close(got, now):
            return None
        text = T(f"{sheet.figure_label(pos)} teilt den Mittelwert durch die Standardabweichung. Der Variationskoeffizient "
                 f"ist es umgekehrt: Standardabweichung durch Mittelwert. So zeigt {address(pos)} {_fmt(now)} statt {_fmt(got)}.",
                 f"{sheet.figure_label(pos)} divides the mean by the standard deviation. The coefficient of variation "
                 f"is the other way round: standard deviation divided by mean. As a result, {address(pos)} shows "
                 f"{_fmt(now)} instead of {_fmt(got)}.")
        return _finding("spread", pos, text, "fehler", _corrected(fixed, got), sheet)
    return None


def _corrected(fixed: str, got: float) -> tuple[str, str]:
    """„Korrigiert: =STABW.S(B2:B9), nachgerechnet 4,2.“"""
    return _tip_pair(f"Korrigiert: ={_german_formula(fixed)}, nachgerechnet {_fmt(got)}.",
             f"Corrected: ={_german_formula(fixed)}, recalculated {_fmt(got)}.")


def _mixed(sheet: Sheet, calc: _Calc, done: set[Pos]) -> tuple[list[Finding], set[str]]:
    """Eine Standardabweichung, die anders rechnet als ihre Nachbarn: STABW.N zwischen lauter STABW.S."""
    found, explained = [], set()
    spreads = {}
    for pos, node in sheet.nodes.items():
        call = _spread_call(node) if node is not None else None
        if call is not None and node is call:
            spreads[pos] = call
    for along in (0, 1):
        groups: dict[int, list[Pos]] = {}
        for pos in spreads:
            groups.setdefault(pos[along], []).append(pos)
        for members in groups.values():
            if len(members) < 3:
                continue
            kinds = [_SPREAD[spreads[p].value] for p in members]
            if len({k for k, _ in kinds}) != 1:
                continue
            variants = [v for _, v in kinds]
            common = max(set(variants), key=variants.count)
            odd = [p for p, v in zip(members, variants, strict=True) if v != common]
            if not odd or len(odd) * 3 > len(members):
                continue
            for pos in sorted(odd):
                if pos in done or address(pos) in explained:
                    continue
                kind = kinds[0][0]
                name = _NAMED[(kind, common)]
                kept = _kept(sheet.cells[pos].formula or "")
                area = _text(spreads[pos].args[0], kept)
                numbers = [v for v in calc.values(spreads[pos].args[0]).values() if _number(v)]
                exact = _exact_spread(numbers, kind, common == "s")
                got = calc.english(f"{name}({area})")
                now = calc.run(sheet.nodes[pos])
                if exact is None or not _number(got) or not _close(got, exact) or not _number(now):
                    continue
                others = len(members) - len(odd)
                said = T("Stichprobe" if common == "s" else "Grundgesamtheit", "sample" if common == "s" else "population")
                text = T(f"{sheet.figure_label(pos)} rechnet {_german(spreads[pos].value)}, die übrigen {others} in "
                         f"{'dieser Zeile' if along == 0 else 'dieser Spalte'} {_german(name)} ({said}). Bei {len(numbers)} "
                         f"Werten zeigt {address(pos)} so {_fmt(now)} statt {_fmt(got)}.",
                         f"{sheet.figure_label(pos)} calculates {spreads[pos].value}, the other {others} in "
                         f"{'this row' if along == 0 else 'this column'} {name} ({said}). With {len(numbers)} "
                         f"values, {address(pos)} therefore shows {_fmt(now)} instead of {_fmt(got)}.")
                tip = _tip_pair(f"Korrigiert: ={_german_formula(name + '(' + area + ')')}, nachgerechnet {_fmt(got)}, wie die Nachbarn. "
                        f"Ist es Absicht, schreib eine Notiz in die Zelle.",
                        f"Corrected: ={_german_formula(name + '(' + area + ')')}, recalculated {_fmt(got)}, like its "
                        f"neighbors. If it is intentional, add a note to the cell.")
                found.append(_finding("spreadmix", pos, text, "pruefen", tip, sheet))
                explained.add(address(pos))
    return found, explained


# --- Anteile von der falschen Summe ---


def _shares(sheet: Sheet, calc: _Calc, done: set[Pos], covered: set[str]) -> list[Finding]:  # noqa: C901 - Gruppen, Summe, Text
    from .summen import _weights

    def cover(members):
        # Wer den Nenner falsch findet, liest sonst den Namen der Zeile („S1“) als Kennzahl oder sieht ein Muster.
        for p, _ in members:
            covered.update(f"{kind}:{address(p)}" for kind in ("label", "pattern", "ratio", "count"))

    found = []
    groups: dict[tuple, list[tuple[Pos, Pos]]] = {}
    swapped: dict[tuple, list[tuple[Pos, Pos]]] = {}
    for pos, node in sheet.nodes.items():
        if node is None or node.kind != "op" or node.value != "/" or pos in sheet.unknown:
            continue
        a, b = node.args
        if a.kind == "ref" and b.kind == "ref":
            groups.setdefault((pos[1], b.value), []).append((pos, a.value))
            swapped.setdefault((pos[1], a.value), []).append((pos, b.value))
    # Je Gruppe die erste Formel, die genau ihre Teile zusammenzählt. Nur die gesuchten Teile: Laufsummen über
    # Tausende Zeilen zählt es so nicht aus.
    wanted = {frozenset(n for _, n in members) for members in groups.values()}
    wanted = [w for w in wanted | {frozenset(d for _, d in members) for members in swapped.values()} if len(w) >= 2]
    totals: dict[frozenset, Pos] = {}
    for pos, node in sheet.nodes.items() if wanted else ():
        weights = _weights(sheet, node) if node is not None else None
        for parts in wanted if weights is not None else ():
            if parts not in totals and all(weights.get(p) == 1.0 for p in parts) and None not in weights and \
                    frozenset(weights) == parts and all(v == 1.0 for v in weights.values()):
                totals[parts] = pos
    for (col, denominator), members in sorted(groups.items()):
        members.sort()
        if len(members) < 2 or any(p in done for p, _ in members):
            continue
        if not (_name(sheet, members[0][0]) & _SHARE):
            continue
        parts = [n for _, n in members]
        right = totals.get(frozenset(parts))
        if right is None or right == denominator:
            continue
        values = [calc.calc.value(p) for p in parts]
        whole = calc.calc.value(denominator)
        if not all(_number(v) for v in values) or not _number(whole):
            continue
        total = math.fsum(values)
        if _close(total, whole) or total == 0:
            continue
        first, start = members[0]
        kept = _kept(sheet.cells[first].formula or "")
        new = f"{kept.get(('ref', start), address(start))}/{_restyle(kept.get(('ref', denominator), address(denominator)), right)}"
        got = [calc.english(new if i == 0 else _shifted(new, members[0][0], p)) for i, (p, _) in enumerate(members)]
        if not all(_number(g) for g in got) or not _close(math.fsum(got), 1.0):
            continue
        shown = math.fsum(calc.run(sheet.nodes[p]) or 0.0 for p, _ in members)
        text = T(f"Die Anteile in {address(members[0][0])}:{address(members[-1][0])} teilen durch {address(denominator)} "
                 f"({_fmt(whole)}), aber ihre Teile {address(parts[0])}:{address(parts[-1])} ergeben zusammen {_fmt(total)}, "
                 f"und diese Summe steht in {address(right)}. Darum ergeben die Anteile zusammen {_percent(shown)} statt 100 %.",
                 f"The shares in {address(members[0][0])}:{address(members[-1][0])} divide by {address(denominator)} "
                 f"({_fmt(whole)}), but their parts {address(parts[0])}:{address(parts[-1])} add up to {_fmt(total)}, "
                 f"and that total is in {address(right)}. So the shares add up to {_percent(shown)} instead of 100 %.")
        tip = _fill_down(first, new, members[-1][0])
        found.append(_finding("base", first, text, "fehler", tip, sheet))
        cover(members)
    for (col, numerator), members in sorted(swapped.items()):
        members.sort()
        if len(members) < 2 or any(p in done for p, _ in members) or not (_name(sheet, members[0][0]) & _SHARE):
            continue
        parts = [d for _, d in members]
        if totals.get(frozenset(parts)) != numerator:
            continue
        first, part = members[0]
        kept = _kept(sheet.cells[first].formula or "")
        new = f"{kept.get(('ref', part), address(part))}/{kept.get(('ref', numerator), address(numerator))}"
        got = [calc.english(new if i == 0 else _shifted(new, first, p)) for i, (p, _) in enumerate(members)]
        if not all(_number(g) for g in got) or not _close(math.fsum(got), 1.0):
            continue
        now = calc.run(sheet.nodes[first])
        if not _number(now):
            continue
        text = T(f"Die Anteile in {address(first)}:{address(members[-1][0])} teilen die Summe {address(numerator)} durch den "
                 f"Teil, nicht den Teil durch die Summe. So zeigt {address(first)} {_percent(now)} statt {_percent(got[0])}.",
                 f"The shares in {address(first)}:{address(members[-1][0])} divide the total {address(numerator)} by "
                 f"the part, not the part by the total. As a result, {address(first)} shows {_percent(now)} instead of "
                 f"{_percent(got[0])}.")
        tip = _fill_down(first, new, members[-1][0])
        found.append(_finding("base", first, text, "fehler", tip, sheet))
        cover(members)
    return found


def _fill_down(first: Pos, new: str, last: Pos) -> tuple[str, str]:
    return _tip_pair(f"Korrigiert in {address(first)}: ={_german_formula(new)}, dann bis {address(last)} herunterziehen. "
             f"Nachgerechnet ergeben die Anteile dann zusammen 100 %.",
             f"Corrected in {address(first)}: ={_german_formula(new)}, then fill it down to {address(last)}. "
             f"Recalculated, the shares then add up to 100 %.")


def _restyle(like: str, pos: Pos) -> str:
    """Die Adresse pos mit denselben $ wie like (E$9 → E$8)."""
    col = ("$" if like.startswith("$") else "") + column_letters(pos[1])
    row = ("$" if "$" in like.lstrip("$") else "") + str(pos[0])
    return col + row


def _shifted(formula: str, start: Pos, pos: Pos) -> str:
    from .xlsx import shift

    return shift(formula, pos[0] - start[0], pos[1] - start[1])


def _percent(share: float) -> str:
    return f"{number_text(round(share * 100, 1))} %"


def checks(sheet: Sheet) -> tuple[list[Finding], set[str]]:
    """Die Hinweise und die Schlüssel allgemeinerer Hinweise (pattern:H7), die sie genauer sagen."""
    calc = _Calc(sheet)
    found: list[Finding] = []
    covered: set[str] = set()
    for step in (_codes, _zeros, _blanks, _spreads, _mixed, _shares):
        done = {formel.position(f.key.partition(":")[2]) for f in found}
        if step in (_codes, _zeros, _shares):
            more = step(sheet, calc, done, covered)
        elif step is _mixed:
            more, explained = step(sheet, calc, done)
            covered |= {f"pattern:{where}" for where in explained}
        else:
            more = step(sheet, calc, done)
        found += more
    for item in found:
        if item.key.startswith("statblank:"):
            # „geteilt durch eine feste Anzahl“ (bereiche.py) sagt dasselbe ohne die leeren Zellen.
            covered.add(f"count:{item.key.partition(':')[2]}")
    return found[:MAX_FINDINGS], covered
