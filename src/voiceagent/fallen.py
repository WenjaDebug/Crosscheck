"""Typische Fallen der Datenanalyse: Formeln, die rechnen, ohne Fehlerwert, aber etwas anderes, als die Mappe meint.

Aus den Prüflisten für Analysen vor dem Weitergeben: Mittel von Mitteln statt gewichtet, ein gewichtetes Mittel durch
die Anzahl statt durch die Gewichte, ein Quartil mit falscher Nummer, Bonferroni ohne Teilen, die Wachstumsrate mit einem Jahr zu viel, gegeneinander verschobene Bereiche in KORREL,
ANZAHL über Text, eine Bedingung, die in einer Spalte fehlt, ein angefangener Monat im Vergleich, ein Anteil über
100 % und ein gleitendes Mittel, das am Anfang weniger Werte nimmt.

Ohne KI: Der Prüfer liest die Namen (Kopf, Wort links), die Form der Formel und die Zahlen. „Fehler“ heißt, die
Mappe sagt, was gemeint ist, und die Korrektur ist nachgerechnet; „prüfen“, wo nur du entscheiden kannst.
"""

from __future__ import annotations

import logging
import re
from datetime import date, timedelta

from . import formel, woerter
from .analyse import Finding, Sheet, _fmt, address
from .kopf import _Ctx, _changes, _close, _finding, _key, _plain, _REF, _rx, _search, _shown, _wrong
from .sprache import T
from .workbook import _left_label
from .xlsx import column_letters

Pos = tuple[int, int]
log = logging.getLogger(__name__)
MAX_EACH = 2


# --- Formeltext ---

_CALL = re.compile(r"(?<![A-Za-z0-9_.!$])(?:_xlfn\.|_xlws\.)?([A-Z][A-Z0-9.]*)\(")


def _calls(text: str) -> list[tuple[str, int, int, list[tuple[int, int]]]]:
    """Jeder Funktionsaufruf im Formeltext: (Name, Anfang, Ende, [(von, bis) je Argument])."""
    quoted, inside = set(), False
    for i, ch in enumerate(text):
        if ch == '"':
            inside = not inside
        elif inside:
            quoted.add(i)
    out = []
    for m in _CALL.finditer(text):
        if m.start() in quoted:
            continue
        depth, start, args, i, inq = 0, m.end(), [], m.end(), False
        while i < len(text):
            ch = text[i]
            if ch == '"':
                inq = not inq
            elif not inq:
                if ch in "({":
                    depth += 1
                elif ch in ")}":
                    if depth == 0:
                        args.append((start, i))
                        break
                    depth -= 1
                elif ch == "," and depth == 0:
                    args.append((start, i))
                    start = i + 1
            i += 1
        else:
            continue
        out.append((m.group(1), m.start(), i + 1, args))
    return out


def _range_rx(r1: int, c: int, r2: int) -> re.Pattern:
    letters = column_letters(c)
    return re.compile(rf"(?<![A-Za-z0-9_.!])(\$?){letters}(\$?){r1}:(\$?){letters}(\$?){r2}(?![0-9])")


def _ref_text(text: str, pos: Pos) -> str | None:
    """Wie der Bezug auf pos im Formeltext steht, mit oder ohne $."""
    m = re.search(_REF.format(col=column_letters(pos[1]), row=pos[0]), text)
    return m.group(0) if m else None


def _walk(node):
    yield node
    for arg in node.args or ():
        if isinstance(arg, formel.Node):
            yield from _walk(arg)


def _formulas(ctx: _Ctx):
    for pos, node in sorted(ctx.sheet.nodes.items()):
        if node is None or pos in ctx.sheet.unknown:
            continue
        text = ctx.sheet.cells[pos].formula if pos in ctx.sheet.cells else None
        if text:
            yield pos, node, text


def _name(ctx: _Ctx, pos: Pos) -> str:
    """Wie eine einzelne Kennzahl heißt: nach dem Wort links („Q1“, „CAGR 2021–2025“), auch wenn ihre Zeile unter
    der Tabelle mitgezählt wird; eine heruntergezogene Spalte nach ihrem Kopf."""
    above, below = ctx.sheet.cells.get((pos[0] - 1, pos[1])), ctx.sheet.cells.get((pos[0] + 1, pos[1]))
    alone = not any(c is not None and c.formula and pattern_same(ctx, (pos[0] + d, pos[1]), pos)
                    for d, c in ((-1, above), (1, below)))
    left = _left_label(ctx.sheet.cells, *pos)
    if alone and left and re.search(r"[^\W\d_]", left) and pos[1] > min(ctx.plan.headers or {1: ""}):
        return left
    return ctx.name(pos)


def pattern_same(ctx: _Ctx, a: Pos, b: Pos) -> bool:
    from .workbook import pattern

    ca, cb = ctx.sheet.cells.get(a), ctx.sheet.cells.get(b)
    return bool(ca and cb and ca.formula and cb.formula and pattern(ca.formula, a) == pattern(cb.formula, b))


def _number(value) -> bool:
    return isinstance(value, float) and not isinstance(value, bool)


def _numbers_in(ctx: _Ctx, r1: int, c: int, r2: int) -> dict[int, float]:
    return {r: v for r in range(r1, r2 + 1) if _number(v := ctx.value((r, c)))}


def _header(ctx: _Ctx, col: int, top: int) -> str:
    """Kopf einer Spalte: aus der Tabelle, sonst der Text über dem Bereich."""
    if col in ctx.plan.headers:
        return ctx.plan.headers[col]
    cell = ctx.sheet.cells.get((top - 1, col))
    return cell.value.strip() if cell is not None and isinstance(cell.value, str) else column_letters(col)


def _day(serial: float) -> date:
    return date(1899, 12, 30) + timedelta(days=int(serial))


def _date_cell(ctx: _Ctx, pos: Pos) -> date | None:
    cell = ctx.sheet.cells.get(pos)
    value = ctx.value(pos)
    if cell is None or not _number(value) or not 1 <= value < 2958466:
        return None
    if isinstance(cell.value, str) and re.fullmatch(r"\s*\d{1,4}[./-]\d{1,2}[./-]\d{1,4}.*", cell.value):
        return _day(value)
    return None


def _de_date(day: date) -> str:
    return f"{day.day:02d}.{day.month:02d}.{day.year}"


# --- Quartil oder Perzentil mit falscher Nummer ---

_QUANTILE = {"QUARTILE": 4, "QUARTILE.INC": 4, "QUARTILE.EXC": 4, "PERCENTILE": 1, "PERCENTILE.INC": 1,
             "PERCENTILE.EXC": 1}


def _wanted_quantile(name: str) -> float | None:
    t = _plain(name, keep_spaces=True)
    if re.search(r"\b(?:median|zentralwert|mediane|mediana)\b", t):
        return 0.5
    if m := re.search(r"(?<![a-z0-9])q\s*([1-3])(?![0-9])", t):
        return int(m.group(1)) / 4
    if m := re.search(r"(?<![0-9])([1-3])\s*\.?\s*(?:st|nd|rd)?\s*-?\s*quarti", t):
        return int(m.group(1)) / 4
    if re.search(r"(?:unter|erst|lower|first|primo|premier)\w*\s*quarti", t):
        return 0.25
    if re.search(r"(?:ober|dritt|upper|third|terzo|troisieme)\w*\s*quarti", t):
        return 0.75
    if m := re.search(r"(?<![a-z0-9])p\s*(\d{1,2})(?![0-9])", t):
        return int(m.group(1)) / 100 if 1 <= int(m.group(1)) <= 99 else None
    if m := re.search(r"(?<![0-9])(\d{1,2})\s*(?:\.|%|th|st|nd|rd|e|es)?\s*-?\s*(?:perzentil|percentil|centile|quantil)", t):
        return int(m.group(1)) / 100 if 1 <= int(m.group(1)) <= 99 else None
    return None


def _quantile_words(share: float) -> str:
    return T({0.25: "das 1. Quartil (25 %)", 0.5: "den Median (50 %)", 0.75: "das 3. Quartil (75 %)"}.get(
        share, f"das {_fmt(round(share * 100, 2))}. Perzentil"),
        {0.25: "the 1st quartile (25 %)", 0.5: "the median (50 %)", 0.75: "the 3rd quartile (75 %)"}.get(
        share, f"the {_ordinal(round(share * 100, 2))} percentile"))


def _ordinal(n: float) -> str:
    """1st, 2nd, 3rd, 4th, 11th, 22nd; 2,5th."""
    if not float(n).is_integer():
        return f"{_fmt(float(n))}th"
    n = int(n)
    end = "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{end}"


def _quantile(ctx: _Ctx) -> list[Finding]:
    found = []
    for pos, _node, text in _formulas(ctx):
        calls = [c for c in _calls(text) if c[0] in _QUANTILE and len(c[3]) == 2]
        if len(calls) != 1:
            continue
        func, _, _, args = calls[0]
        a, b = args[1]
        try:
            has = float(text[a:b].strip()) / _QUANTILE[func]
        except ValueError:
            continue
        name = _name(ctx, pos)
        want = _wanted_quantile(name)
        if want is None or _close(want, has):
            continue
        if _QUANTILE[func] == 4:
            if not float(want * 4).is_integer():
                continue
            arg = str(int(want * 4))
        else:
            arg = _fmt(want).replace(",", ".")
        new = text[:a] + arg + text[b:]
        bad = _wrong(ctx, {pos: new})
        if not bad:
            continue
        _, now, right = bad
        msg = T(f"{address(pos)} heißt „{name}“, ={ctx.german(text)} rechnet aber {_quantile_words(has)}. Für "
                f"{_quantile_words(want)} gehört {arg.replace('.', ',')} hinein: {_shown(ctx, pos, right)} statt "
                f"{_shown(ctx, pos, now)}.",
                f"{address(pos)} is called “{name}”, but ={ctx.german(text)} calculates {_quantile_words(has)}. For "
                f"{_quantile_words(want)}, it needs {arg.replace('.', ',')}: {_shown(ctx, pos, right)} instead of "
                f"{_shown(ctx, pos, now)}.")
        found.append(_finding("quantil", ctx, [pos], msg, "fehler", {pos: new}))
    return found[:MAX_EACH]


# --- Bonferroni im Namen, die Schwelle aber unkorrigiert ---

_BONFERRONI = _rx("bonferroni")
_ALPHA = re.compile(r"(<=?)\s*(0?\.(?:05|01|1|001|10|050|010))(?![0-9.]|\s*/)")


def _bonferroni(ctx: _Ctx) -> list[Finding]:
    found = []
    for run in ctx.runs:
        first = run[0]
        name = ctx.name(first)
        if len(run) < 3 or not _search(_BONFERRONI, name):
            continue
        text = ctx.sheet.cells[first].formula or ""
        hits = list(_ALPHA.finditer(text))
        if len(hits) != 1:
            continue
        own = {c for r, c in formel.references(ctx.sheet.nodes[first])[0] if r == first[0] and c != first[1]}
        if len(own) != 1:
            continue
        col = own.pop()
        values = _numbers_in(ctx, min(p[0] for p in run), col, max(p[0] for p in run))
        if len(values) < 3 or not all(0 <= v <= 1 for v in values.values()):
            continue
        letters = column_letters(col)
        span = f"${letters}${min(values)}:${letters}${max(values)}"
        alpha = float(hits[0].group(2))

        def make(t, pos):
            m = list(_ALPHA.finditer(t))
            return t[:m[0].end(2)] + f"/COUNT({span})" + t[m[0].end(2):] if len(m) == 1 else None

        changes = _changes(ctx, run, make)
        if not changes:
            continue
        differ = [(p, ctx.run(ctx.sheet.cells[p].formula), ctx.run(t)) for p, t in changes.items()]
        differ = [d for d in differ if d[2] is not None and d[1] != d[2]]
        if not differ:
            continue
        where, now, right = differ[0]
        m = len(values)
        msg = T(f"{ctx.label(run)} soll nach Bonferroni korrigieren, vergleicht jeden p-Wert aber mit {_fmt(alpha)}, als gäbe es "
                f"nur einen Test. Bonferroni teilt die Schwelle durch die Zahl der Tests: {_fmt(alpha)}/{m} = "
                f"{_fmt(round(alpha / m, 6))}. So ändert sich das Ergebnis in {len(differ)} von {len(run)} Zeilen, "
                f"in {address(where)} „{_fmt(right)}“ statt „{_fmt(now)}“.",
                f"{ctx.label(run)} is meant to apply the Bonferroni correction, but compares each p-value with "
                f"{_fmt(alpha)} as if there were only one test. The Bonferroni correction divides the threshold by the "
                f"number of tests: {_fmt(alpha)}/{m} = {_fmt(round(alpha / m, 6))}. This changes the result in "
                f"{len(differ)} of {len(run)} rows, in {address(where)} “{_fmt(right)}” instead of “{_fmt(now)}”.")
        found.append(_finding("bonf", ctx, run, msg, "fehler", changes))
    return found[:MAX_EACH]


# --- Jährliche Wachstumsrate mit einem Jahr zu viel ---

_YEARS = re.compile(r"(?<!\d)((?:19|20)\d\d)(?!\d)")


def _growth_parts(node):
    """(X/Y)^(1/N) − 1 oder POTENZ(X/Y;1/N) − 1: (X, Y, N)."""
    if node.kind != "op" or node.value != "-" or len(node.args) != 2 or node.args[1].kind != "num" or \
            node.args[1].value != 1:
        return None
    power = node.args[0]
    if power.kind == "op" and power.value == "^" and len(power.args) == 2:
        base, exp = power.args
    elif power.kind == "call" and power.value == "POWER" and len(power.args) == 2:
        base, exp = power.args
    else:
        return None
    if base.kind != "op" or base.value != "/" or not all(a.kind == "ref" for a in base.args):
        return None
    if exp.kind != "op" or exp.value != "/" or exp.args[0].kind != "num" or exp.args[0].value != 1 or \
            exp.args[1].kind != "num" or not float(exp.args[1].value).is_integer():
        return None
    return base.args[0].value, base.args[1].value, int(exp.args[1].value)


def _year_of(ctx: _Ctx, row: int, skip: int) -> dict[int, int]:
    """Jahre, die in einer Zeile stehen: Spalte → Jahr (eine Zahl 1900–2100 oder ein Datum)."""
    out = {}
    for (r, c), cell in ctx.sheet.cells.items():
        if r != row or c == skip or cell.formula:
            continue
        day = _date_cell(ctx, (r, c))
        value = ctx.value((r, c))
        if day is not None:
            out[c] = (day.year, day.month, day.day)
        elif _number(value) and value.is_integer() and 1900 <= value <= 2100:
            out[c] = (int(value), 0, 0)
    return out


def _cagr(ctx: _Ctx) -> list[Finding]:
    found = []
    for pos, node, text in _formulas(ctx):
        parts = _growth_parts(node)
        if parts is None:
            continue
        end, start, n = parts
        if end[1] != start[1] or end[0] == start[0]:
            continue
        name = _name(ctx, pos)
        years, said = None, ""
        a, b = _year_of(ctx, start[0], start[1]), _year_of(ctx, end[0], end[1])
        for col in sorted(set(a) & set(b)):
            if a[col][1:] == b[col][1:]:
                years, said = b[col][0] - a[col][0], T(f"von {a[col][0]} bis {b[col][0]}",
                                                      f"from {a[col][0]} to {b[col][0]}")
                break
        if years is None:
            named = [int(y) for y in _YEARS.findall(name)]
            if len(named) == 2:
                years, said = named[1] - named[0], T(f"von {named[0]} bis {named[1]}",
                                                     f"from {named[0]} to {named[1]}")
        if not years or years <= 0 or years == n:
            continue
        hits = list(re.finditer(rf"1\s*/\s*0*{n}(?:\.0+)?(?![0-9.])", text))
        if len(hits) != 1:
            continue
        new = text[:hits[0].start()] + f"1/{years}" + text[hits[0].end():]
        bad = _wrong(ctx, {pos: new})
        if not bad:
            continue
        _, now, right = bad
        why = T(f" Es sind {years + 1} Werte, aber nur {years} Schritte dazwischen.",
                f" There are {years + 1} values, but only {years} steps between them.") if n == years + 1 else ""
        msg = T(f"{address(pos)} ({name}) zieht die {n}. Wurzel aus {address(end)}/{address(start)}, {said} sind es "
                f"aber {years} Jahre.{why} Die jährliche Rate ist {_shown(ctx, pos, right)} statt "
                f"{_shown(ctx, pos, now)}.",
                f"{address(pos)} ({name}) takes the {_ordinal(n)} root of {address(end)}/{address(start)}, but {said} "
                f"it is {years} {'year' if years == 1 else 'years'}.{why} The annual rate is "
                f"{_shown(ctx, pos, right)} instead of {_shown(ctx, pos, now)}.")
        found.append(_finding("cagr", ctx, [pos], msg, "fehler", {pos: new}))
    return found[:MAX_EACH]


# --- Gewichtetes Mittel durch die Anzahl statt durch die Gewichte ---

_WEIGHTED = _rx(r"gewichtet|weighted|ponder|\bgew\b")
_UNWEIGHTED = _rx(r"ungewichtet|unweighted|non ponder")


def _weighted(ctx: _Ctx) -> list[Finding]:
    found = []
    for pos, node, text in _formulas(ctx):
        name = _name(ctx, pos)
        if not _search(_WEIGHTED, name) or _search(_UNWEIGHTED, name):
            continue
        if node.kind != "op" or node.value != "/" or len(node.args) != 2:
            continue
        top, bottom = node.args
        if top.kind != "call" or top.value != "SUMPRODUCT" or len(top.args) != 2 or \
                not all(a.kind == "range" and a.value[1] == a.value[3] for a in top.args):
            continue
        if not (bottom.kind == "num" or (bottom.kind == "call" and bottom.value in ("COUNT", "COUNTA", "ROWS"))):
            continue
        calls = _calls(text)
        if not calls or calls[0][0] != "SUMPRODUCT" or calls[0][1] != 0 or not text[calls[0][2]:].lstrip().startswith("/"):
            continue
        (r1, c1, r2, _), (s1, c2, s2, _) = (a.value for a in top.args)
        if (r1, r2) != (s1, s2):
            continue
        heads = {c1: _header(ctx, c1, r1), c2: _header(ctx, c2, r1)}
        words = _key(name)
        measure = [c for c in heads if words & _key(heads[c])]
        if len(measure) != 1:
            whole = {c: all(v.is_integer() and v >= 0 for v in _numbers_in(ctx, r1, c, r2).values()) for c in heads}
            measure = [c for c in heads if not whole[c]] if sum(whole.values()) == 1 else []
        if len(measure) != 1:
            continue
        weight = c2 if measure[0] == c1 else c1
        args = calls[0][3]
        span = text[slice(*args[0 if weight == c1 else 1])].strip()
        new = text[:calls[0][2]] + f"/SUM({span})"
        bad = _wrong(ctx, {pos: new})
        if not bad:
            continue
        _, now, right = bad
        below = text[calls[0][2]:].lstrip()[1:]
        divisor = ctx.run(below)
        divisor_text = f" = {_fmt(divisor)}" if _number(divisor) else ""
        msg = T(f"{address(pos)} heißt „{name}“, teilt die Summe der Produkte aber durch "
                f"{ctx.german(below)}{divisor_text}, die Zahl der Zeilen. Ein nach "
                f"{heads[weight]} gewichtetes Mittel teilt durch die Summe der Gewichte, SUMME({ctx.german(span)}). "
                f"Richtig wäre {_shown(ctx, pos, right)} statt {_shown(ctx, pos, now)}.",
                f"{address(pos)} is called “{name}”, but divides the sum of products by "
                f"{below.strip()}{divisor_text}, the number of rows. A mean weighted by "
                f"{heads[weight]} divides by the sum of the weights, SUM({span}). "
                f"It should be {_shown(ctx, pos, right)} instead of {_shown(ctx, pos, now)}.")
        found.append(_finding("wmean", ctx, [pos], msg, "fehler", {pos: new}))
    return found[:MAX_EACH]


# --- Mittel von Gruppenmitteln in der Gesamtzeile, obwohl die Gruppen verschieden groß sind ---

_MEANISH = re.compile(r"(?:^|[\s(])(?:ø|⌀)|\bmittel|durchschnitt|\baverage\b|\bmean\b|\bavg\b|\bmoyenne\b|\bmedia\b|"
                      r"\brate\b|rate\b|quote\b|\banteil|\bratio\b|\btaux\b|\btasso\b|marge\b|margin|rendite\b")
# Wonach eine Marge oder Rendite gewichtet ist: nach dem Umsatz der Zeile.
_REVENUE = _rx(r"umsatz|umsatze|erlos|erlose|ertrag|ertrage|revenue|revenues|sales|turnover|chiffre|fatturato|ricavi")
_COUNTISH = _rx(r"fall|falle|\bn\b|anzahl|patient|count|number|menge|stuck|cases|personen|einwohner|betten|behandelt|"
                r"teilnehmer|probanden|befragte|mitarbeiter|kunden|besuche|packungen|nombre|numero|cas\b|casi\b")
_TOTAL = _rx(r"gesamt|summe|total|insgesamt|alle|overall|totale|ensemble|tutti|tous")


def _group_mean(ctx: _Ctx) -> list[Finding]:
    found = []
    for pos, node, text in _formulas(ctx):
        if node.kind != "call" or node.value != "AVERAGE" or len(node.args) != 1 or node.args[0].kind != "range":
            continue
        r1, c, r2, c_end = node.args[0].value
        if c != c_end or r2 - r1 < 2 or not (r2 < pos[0]):
            continue
        if any((r, c) in ctx.sheet.nodes for r in range(r1, r2 + 1)):
            continue  # Quoten je Zeile als Formel: das sagt quoten.py genauer
        head = _header(ctx, c, r1)
        if not _MEANISH.search(_plain(head, keep_spaces=True)) or _search(_UNWEIGHTED, _name(ctx, pos)):
            continue
        weights = []
        for w in sorted(ctx.plan.headers):
            money = _search(_REVENUE, ctx.plan.headers[w]) and woerter.rate_word(_plain(head, keep_spaces=True).split())
            if w == c or not (_search(_COUNTISH, ctx.plan.headers[w]) or money) or \
                    _MEANISH.search(_plain(ctx.plan.headers[w], True)):
                continue
            values = _numbers_in(ctx, r1, w, r2)
            if len(values) != r2 - r1 + 1 or not all(v.is_integer() and v >= 0 for v in values.values()) or \
                    len(set(values.values())) == 1:
                continue
            total = ctx.sheet.nodes.get((pos[0], w))
            sums = total is not None and total.kind == "call" and total.value == "SUM" and \
                [a.value for a in total.args if a.kind == "range"] == [(r1, w, r2, w)]
            weights.append((sums, w))
        label = ctx.sheet.cells.get((pos[0], min(ctx.plan.headers) if ctx.plan.headers else 1))
        totalrow = label is not None and isinstance(label.value, str) and _search(_TOTAL, label.value)
        weights = [w for sums, w in weights if sums] or ([w for _, w in weights] if totalrow else [])
        if len(weights) != 1 or not (totalrow or any(True for _ in weights)):
            continue
        w = weights[0]
        lw, lc = column_letters(w), column_letters(c)
        new = f"SUMPRODUCT({lw}{r1}:{lw}{r2},{lc}{r1}:{lc}{r2})/SUM({lw}{r1}:{lw}{r2})"
        bad = _wrong(ctx, {pos: new})
        if not bad:
            continue
        _, now, right = bad
        counts = _numbers_in(ctx, r1, w, r2).values()
        whead = ctx.plan.headers[w]
        if _search(_REVENUE, whead):
            msg = T(f"{address(pos)} ({_name(ctx, pos)}) mittelt die {r2 - r1 + 1} Werte in {head} gleich gewichtet, "
                    f"obwohl {whead} je Zeile sehr verschieden ist ({_fmt(min(counts))} bis {_fmt(max(counts))}). So zählt "
                    f"ein kleiner Bereich so viel wie ein großer. {head} über alles ist nach {whead} gewichtet: "
                    f"{_shown(ctx, pos, right)} statt {_shown(ctx, pos, now)}.",
                    f"{address(pos)} ({_name(ctx, pos)}) averages the {r2 - r1 + 1} values in {head} with equal weights, "
                    f"although {whead} varies a lot from row to row ({_fmt(min(counts))} to {_fmt(max(counts))}). That "
                    f"way a small segment counts as much as a large one. {head} overall is weighted by {whead}: "
                    f"{_shown(ctx, pos, right)} instead of {_shown(ctx, pos, now)}.")
        else:
            msg = T(f"{address(pos)} ({_name(ctx, pos)}) mittelt die {r2 - r1 + 1} Werte in {head} gleich gewichtet, obwohl "
                    f"hinter ihnen verschieden viele {whead} stehen ({_fmt(min(counts))} bis {_fmt(max(counts))}). So zählt "
                    f"eine kleine Gruppe so viel wie eine große. Das Mittel über alle {_fmt(sum(counts))} {whead} ist nach "
                    f"{whead} gewichtet: {_shown(ctx, pos, right)} statt {_shown(ctx, pos, now)}.",
                    f"{address(pos)} ({_name(ctx, pos)}) averages the {r2 - r1 + 1} values in {head} with equal weights, "
                    f"although they stand for different numbers of {whead} ({_fmt(min(counts))} to "
                    f"{_fmt(max(counts))}). That way a small group counts as much as a large one. The mean over all "
                    f"{_fmt(sum(counts))} {whead} is weighted by {whead}: {_shown(ctx, pos, right)} instead of "
                    f"{_shown(ctx, pos, now)}.")
        found.append(_finding("avgmeans", ctx, [pos], msg, "fehler" if totalrow else "pruefen", {pos: new}))
    return found[:MAX_EACH]


# --- Bereiche in KORREL & Co. gegeneinander verschoben ---

_PAIRED = {"CORREL", "PEARSON", "RSQ", "SLOPE", "INTERCEPT", "STEYX", "COVARIANCE.S", "COVARIANCE.P", "COVAR"}
_LAG = _rx(r"lag|verzoger|versetzt|vorjahr|vormonat|vorwoche|vortag|zeitversetzt|lead|shift|verschoben|vorher|danach")


def _pair_shift(ctx: _Ctx) -> list[Finding]:
    found = []
    for pos, node, text in _formulas(ctx):
        name = _name(ctx, pos)
        if _search(_LAG, name):
            continue
        for call in _walk(node):
            if call.kind != "call" or call.value not in _PAIRED or len(call.args) != 2 or \
                    not all(a.kind == "range" and a.value[1] == a.value[3] for a in call.args):
                continue
            (a1, ca, a2, _), (b1, cb, b2, _) = (a.value for a in call.args)
            if a2 - a1 != b2 - b1 or a1 == b1 or a2 - a1 < 2:
                continue
            fill = {ca: _numbers_in(ctx, a1, ca, a2), cb: _numbers_in(ctx, b1, cb, b2)}
            fits = {ca: len(fill[ca]) == a2 - a1 + 1, cb: len(fill[cb]) == b2 - b1 + 1}
            if fits[ca] == fits[cb]:
                continue
            good, bad_range = ((a1, ca, a2), (b1, cb, b2)) if fits[ca] else ((b1, cb, b2), (a1, ca, a2))
            if len(_numbers_in(ctx, good[0], bad_range[1], good[2])) != good[2] - good[0] + 1:
                continue
            rx = _range_rx(*bad_range)
            if len(rx.findall(text)) != 1:
                continue
            letters = column_letters(bad_range[1])
            new = rx.sub(lambda m: f"{m.group(1)}{letters}{m.group(2)}{good[0]}:{m.group(3)}{letters}{m.group(4)}{good[2]}",
                         text)
            wrong = _wrong(ctx, {pos: new})
            if not wrong:
                continue
            _, now, right = wrong
            shift = bad_range[0] - good[0]
            outside = [r for r in range(bad_range[0], bad_range[2] + 1) if r not in _numbers_in(ctx, *bad_range)]
            gbad = f"{address((bad_range[0], bad_range[1]))}:{address((bad_range[2], bad_range[1]))}"
            ggood = f"{address((good[0], good[1]))}:{address((good[2], good[1]))}"
            empty = ", ".join(address((r, bad_range[1])) for r in outside[:3])
            msg = T(f"{address(pos)} ({name}) paart {ggood} mit {gbad}: Die Bereiche sind um {abs(shift)} "
                    f"Zeile{'n' if abs(shift) > 1 else ''} gegeneinander verschoben, so steht jeder Wert neben dem einer "
                    f"anderen Zeile, und in {empty} steht keine "
                    f"Zahl. Zeile für Zeile gepaart ergibt es {_shown(ctx, pos, right)} statt {_shown(ctx, pos, now)}.",
                    f"{address(pos)} ({name}) pairs {ggood} with {gbad}: the ranges are offset from each other by "
                    f"{abs(shift)} row{'s' if abs(shift) > 1 else ''}, so each value is paired with one from a different "
                    f"row, and there is no number in {empty}. Paired row by row, it gives {_shown(ctx, pos, right)} "
                    f"instead of {_shown(ctx, pos, now)}.")
            found.append(_finding("corrshift", ctx, [pos], msg, "fehler", {pos: new}))
            break
    return found[:MAX_EACH]


# --- ANZAHL über Text ---

def _count_text(ctx: _Ctx) -> list[Finding]:
    found = []
    for pos, node, text in _formulas(ctx):
        for call in _walk(node):
            if call.kind != "call" or call.value != "COUNT" or len(call.args) != 1 or call.args[0].kind != "range":
                continue
            r1, c1, r2, c2 = call.args[0].value
            if c1 != c2:
                continue
            values = [ctx.value((r, c1)) for r in range(r1, r2 + 1)]
            texts = [v for v in values if isinstance(v, str) and v.strip()]
            if len(texts) < 3 or any(_number(v) for v in values):
                continue
            if any(re.fullmatch(r"\s*[-+]?[\d.,' ]+\s*%?\s*", v) for v in texts):
                continue  # Zahlen als Text: das sagt alstext.py
            rx = re.compile(r"(?<![A-Za-z0-9_.])COUNT\((\s*)" + _range_rx(r1, c1, r2).pattern[len("(?<![A-Za-z0-9_.!])"):])
            if len(rx.findall(text)) != 1:
                continue
            new = rx.sub(lambda m: "COUNTA(" + m.group(0)[len("COUNT("):], text)
            wrong = _wrong(ctx, {pos: new})
            if not wrong:
                continue
            _, now, right = wrong
            head = _header(ctx, c1, r1)
            span = f"{address((r1, c1))}:{address((r2, c1))}"
            msg = T(f"{address(pos)} ({_name(ctx, pos)}) zählt mit ANZAHL, und ANZAHL zählt nur Zahlen. In {span} "
                    f"({head}) stehen aber {len(texts)} Texte wie „{texts[0]}“ und keine Zahl, darum ergibt es "
                    f"{_fmt(now)}. Die Einträge zählt ANZAHL2: {_fmt(right)}.",
                    f"{address(pos)} ({_name(ctx, pos)}) counts with COUNT, and COUNT only counts numbers. But {span} "
                    f"({head}) holds {len(texts)} texts such as “{texts[0]}” and no number, so it gives "
                    f"{_fmt(now)}. COUNTA counts the entries: {_fmt(right)}.")
            found.append(_finding("counttext", ctx, [pos], msg, "fehler", {pos: new}))
            break
    return found[:MAX_EACH]


# --- ZÄHLENWENNS ohne die Bedingung der Zeile, die die Nachbarspalte hat ---

_IFS = {"COUNTIFS": 0, "SUMIFS": 1, "AVERAGEIFS": 1, "MAXIFS": 1, "MINIFS": 1}


def _row_pairs(text: str, row: int) -> list[str] | None:
    """Die Paare (Bereich, Bedingung) eines ...WENNS, deren Bedingung eine Zelle der eigenen Zeile liest."""
    calls = _calls(text)
    if not calls or calls[0][1] != 0 or calls[0][2] != len(text) or calls[0][0] not in _IFS:
        return None
    args = [text[a:b] for a, b in calls[0][3]][_IFS[calls[0][0]]:]
    if len(args) % 2:
        return None
    own = re.compile(rf"(?<![A-Za-z0-9_.!$])\$?[A-Z]{{1,3}}{row}(?![0-9])")
    return [f"{args[i]},{args[i + 1]}" for i in range(0, len(args), 2) if own.search(args[i + 1])]


def _no_criterion(ctx: _Ctx) -> list[Finding]:
    found = []
    by_rows: dict[tuple[int, ...], list[list[Pos]]] = {}
    for run in ctx.runs:
        by_rows.setdefault(tuple(p[0] for p in run), []).append(run)
    for rows, runs in by_rows.items():
        if len(rows) < 3:
            continue
        for run in runs:
            texts = {ctx.sheet.cells[p].formula for p in run}
            values = {ctx.value(p) for p in run}
            if len(texts) != 1 or len(values) != 1 or not _number(next(iter(values))):
                continue
            text = texts.pop()
            if _row_pairs(text, run[0][0]) != []:
                continue
            func = _calls(text)[0][0]
            for other in runs:
                if other is run or _calls(ctx.sheet.cells[other[0]].formula or "")[:1] == [] or \
                        len({ctx.value(p) for p in other}) < 2:
                    continue

                def make(t, pos, col=other[0][1]):
                    pairs = _row_pairs(ctx.sheet.cells[(pos[0], col)].formula or "", pos[0])
                    return t[:-1] + "," + ",".join(pairs) + ")" if pairs else None

                if _calls(ctx.sheet.cells[other[0]].formula)[0][0] not in _IFS:
                    continue
                changes = _changes(ctx, run, make)
                bad = changes and _wrong(ctx, changes)
                if not bad or len({ctx.run(t) for t in changes.values()}) < 2:
                    continue
                pos, now, right = bad
                said = _row_pairs(ctx.sheet.cells[(pos[0], other[0][1])].formula, pos[0])
                cells = sorted({m for p in said for m in re.findall(rf"\$?[A-Z]{{1,3}}{pos[0]}(?![0-9])", p.split(",", 1)[1])})
                own = ", ".join(c.replace('$', '') for c in cells)
                msg = T(f"{ctx.label(run)} ergibt in jeder Zeile dasselbe ({_fmt(now)}), weil {func} die Bedingung der "
                        f"Zeile nicht enthält. Die Nachbarspalte {ctx.name(other[0])} grenzt mit "
                        f"{own} auf ihre Zeile ein, {ctx.name(run[0])} zählt "
                        f"dagegen immer über alles. Mit derselben Bedingung ergibt {address(pos)} {_fmt(right)} statt "
                        f"{_fmt(now)}.",
                        f"{ctx.label(run)} gives the same result in every row ({_fmt(now)}) because {func} does not "
                        f"include the row's criterion. The neighboring column {ctx.name(other[0])} narrows down to its "
                        f"row with {own}, while {ctx.name(run[0])} always counts over everything. With the same "
                        f"criterion, {address(pos)} gives {_fmt(right)} instead of {_fmt(now)}.")
                found.append(_finding("nocrit", ctx, run, msg, "fehler", changes))
                break
    return found[:MAX_EACH]


# --- Der letzte Zeitraum ist angefangen und wird mit ganzen verglichen ---

def _months_after(day: date, months: int) -> date:
    y, m = divmod(day.month - 1 + months, 12)
    return date(day.year + y, m + 1, 1) if day.day == 1 else day


def _part_period(ctx: _Ctx) -> list[Finding]:
    found = []
    for run in ctx.runs:
        if len(run) < 3:
            continue
        last = run[-1]
        text = ctx.sheet.cells[last].formula or ""
        calls = _calls(text)
        if not calls or calls[0][0] not in _IFS or calls[0][1] != 0:
            continue
        args = [text[a:b] for a, b in calls[0][3]][_IFS[calls[0][0]]:]
        label_col, span = None, None
        for i in range(0, len(args) - 1, 2):
            m = re.search(rf"(?<![A-Za-z0-9_.!$])\$?([A-Z]{{1,3}}){last[0]}(?![0-9])", args[i + 1])
            if m:
                label_col, span = formel.position(f"{m.group(1)}{last[0]}")[1], args[i]
                break
        if label_col is None:
            continue
        days = [_date_cell(ctx, (p[0], label_col)) for p in run]
        if any(d is None for d in days):
            continue
        monthly = all(d.day == 1 for d in days) and all(
            _months_after(a, 1) == b for a, b in zip(days, days[1:]))
        step = (days[-1] - days[-2]).days
        if not monthly and (step <= 0 or any((b - a).days != step for a, b in zip(days, days[1:]))):
            continue
        end = _months_after(days[-1], 1) if monthly else days[-1] + timedelta(days=step)
        top = ctx.run(f"MAX({span})")
        if not _number(top):
            continue
        newest = _day(top)
        if not days[-1] <= newest < end - timedelta(days=1):
            continue
        have, full = (newest - days[-1]).days + 1, (end - days[-1]).days
        compare = None
        for (r, c), node in ctx.sheet.nodes.items():
            if r == last[0] and c != last[1] and node is not None:
                cells = formel.references(node)[0]
                if last in cells and (last[0] - 1, last[1]) in cells:
                    compare = (r, c)
                    break
        what = "Monat" if monthly else ("Woche" if step == 7 else "Zeitraum")
        what_en = "month" if monthly else ("week" if step == 7 else "period")
        cmp = T(f" {address(compare)} vergleicht ihn trotzdem mit dem ganzen Vor{what.lower()} "
                f"({_shown(ctx, compare, ctx.value(compare))}).",
                f" {address(compare)} still compares it with the whole previous {what_en} "
                f"({_shown(ctx, compare, ctx.value(compare))}).") if compare else ""
        msg = T(f"Der letzte {what} in {address(last)} ({_de_date(days[-1])}) ist angefangen: Die Daten reichen nur bis "
                f"{_de_date(newest)}, das sind {have} von {full} Tagen.{cmp} Vergleiche nur ganze Zeiträume oder rechne "
                f"auf gleich viele Tage um.",
                f"The last {what_en} in {address(last)} ({_de_date(days[-1])}) is incomplete: the data only go up to "
                f"{_de_date(newest)}, which is {have} of {full} days.{cmp} Compare only complete periods, or scale "
                f"them to the same number of days.")
        target = compare or last
        found.append(Finding(f"partmonth:{address(target)}", msg, "pruefen", ""))
    return found[:MAX_EACH]


# --- Ein Anteil über 100 % ---

_SHARE = _rx(r"rate|quote|anteil|prozent|share|ratio|taux|tasso|percent|%")


def _over_hundred(ctx: _Ctx) -> list[Finding]:
    found = []
    for run in ctx.runs:
        first = run[0]
        if len(run) < 4 or first[0] not in ctx.data:
            continue
        node = ctx.sheet.nodes[first]
        if node.kind != "op" or node.value != "/" or not all(a.kind == "ref" and a.value[0] == first[0] for a in node.args):
            continue
        cell = ctx.sheet.cells[first]
        if not (_search(_SHARE, ctx.name(first)) or str(cell.value).strip().endswith("%")):
            continue
        values = {p: ctx.value(p) for p in run if p[0] in ctx.data}
        if len(values) < 4 or not all(_number(v) for v in values.values()):
            continue
        over = [p for p, v in values.items() if v > 1]
        if not over or len(over) > max(1, len(values) // 5) or any(v < 0 for v in values.values()):
            continue
        for pos in over:
            top, bottom = (a.value for a in ctx.sheet.nodes[pos].args)
            if top[0] != pos[0]:
                continue
            msg = T(f"{address(pos)} ({ctx.name(pos)}) zeigt {_shown(ctx, pos, values[pos])}: {ctx.name(top)} in "
                    f"{address(top)} ({_fmt(ctx.value(top))}) ist größer als {ctx.name(bottom)} in {address(bottom)} "
                    f"({_fmt(ctx.value(bottom))}). Ein Anteil kann nicht über 100 % liegen, in allen anderen Zeilen liegt "
                    f"er darunter. Einer der beiden Werte ist wohl falsch eingetippt.",
                    f"{address(pos)} ({ctx.name(pos)}) shows {_shown(ctx, pos, values[pos])}: {ctx.name(top)} in "
                    f"{address(top)} ({_fmt(ctx.value(top))}) is greater than {ctx.name(bottom)} in {address(bottom)} "
                    f"({_fmt(ctx.value(bottom))}). A share cannot be above 100 %, and in all other rows it is below. "
                    f"One of the two values was probably mistyped.")
            found.append(Finding(f"over100:{address(top)}", msg, "pruefen", ""))
    return found[:MAX_EACH]


# --- Gleitendes Mittel, das am Anfang weniger Werte nimmt ---

_WINDOWED = {"AVERAGE", "MEDIAN"}
_MOVING = _rx(r"gleitend|rollierend|moving|rolling|glissant|mobile|ø|⌀|schnitt|mittel|average|mean")


def _window(node, row: int) -> tuple[str, int, int] | None:
    if node is None or node.kind != "call" or node.value not in _WINDOWED or len(node.args) != 1 or \
            node.args[0].kind != "range":
        return None
    r1, c1, r2, c2 = node.args[0].value
    if c1 != c2 or r2 != row:
        return None
    return node.value, c1, r2 - r1 + 1


def _window_start(ctx: _Ctx) -> list[Finding]:
    found = []
    for run in ctx.runs:
        first = run[0]
        if len(run) < 3:
            continue
        shape = _window(ctx.sheet.nodes[first], first[0])
        if shape is None or shape[2] < 2:
            continue
        func, col, size = shape
        name = ctx.name(first)
        if not _search(_MOVING, name) and str(size) not in re.findall(r"\d+", name):
            continue
        start = ctx.sheet.nodes[first].args[0].value[0]
        short = []
        r = first[0] - 1
        while r >= start:
            here = _window(ctx.sheet.nodes.get((r, first[1])), r)
            if here is None or here[:2] != (func, col) or here[2] >= size or \
                    ctx.sheet.nodes[(r, first[1])].args[0].value[0] != start:
                break
            short.insert(0, (r, first[1]))
            r -= 1
        if not short:
            continue
        sizes = [_window(ctx.sheet.nodes[p], p[0])[2] for p in short]
        cells = f"{address(short[0])}" if len(short) == 1 else f"{address(short[0])}:{address(short[-1])}"
        msg = T(f"{cells} ({ctx.name(short[0])}) nehmen nur {', '.join(map(str, sizes))} "
                f"Wert{'e' if sizes != [1] else ''}, ab {address(first)} sind es {size}. Diese ersten Werte sind kein "
                f"{size}er-Mittel und schwanken stärker als die übrigen. Lass sie leer oder kennzeichne sie, sonst "
                f"wirken sie wie Teil der Reihe.",
                f"{cells} ({ctx.name(short[0])}) {'takes' if len(short) == 1 else 'take'} only "
                f"{', '.join(map(str, sizes))} value{'s' if sizes != [1] else ''}; from {address(first)} on, it is "
                f"{size}. These first values are not a {size}-value average and fluctuate more than the rest. Leave "
                f"them empty or mark them, otherwise they look like part of the series.")
        found.append(Finding(f"winstart:{address(short[0])}", msg, "pruefen", ""))
    return found[:MAX_EACH]


CHECKS = (_quantile, _bonferroni, _cagr, _weighted, _group_mean, _pair_shift, _count_text, _no_criterion,
          _part_period, _over_hundred, _window_start)


def checks(sheet: Sheet, rows: list[int]) -> list[Finding]:
    if sheet.plan.header_row is None:
        return []
    ctx = _Ctx(sheet, rows)
    found: list[Finding] = []
    seen: set[str] = set()
    for check in CHECKS:
        try:
            more = check(ctx)
        except Exception:  # noqa: BLE001 - eine Prüfung darf die anderen nicht mitreißen
            log.exception("Prüfung %s", check.__name__)
            continue
        for item in more:
            where = item.key.partition(":")[2]
            if where not in seen:
                seen.add(where)
                found.append(item)
    return found
