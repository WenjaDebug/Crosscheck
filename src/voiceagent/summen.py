"""Teile und Ganzes: passen Summenzeilen, Summenspalten und Gesamtsummen zu ihren Teilen?

„Summe“ unter einer Übersicht, „Total“ rechts neben den Monaten, „Umsatz gesamt“ unter Umsatz Nord, Ost und Süd:
Der Prüfer zählt die Teile selbst zusammen und sagt, welcher Teil fehlt, doppelt zählt, abgezogen wird oder nicht
dazugehört, mit einer Summe, die er nachgerechnet hat. Rechnen die Teile aus einer Tabelle (SUMMEWENN je Region),
prüft er auch, ob sie zusammen alle Zeilen erfassen: eine Gruppe ohne eigene Zeile („West“), ein vertippter Wert
(„Nrod“) oder eine leere Zelle. Alles ohne KI.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from functools import lru_cache

from . import formel, woerter
from .analyse import Finding, Sheet, _fmt, address
from .bereiche import _FILLER, _MEANS, _MONTHS, _QUARTERS, _german_formula, _head, _number_at, _rounding, _wanted, _words
from .formel import Node
from .sprache import T
from .xlsx import column_letters

Pos = tuple[int, int]

MAX_FINDINGS = 6
_TOTALS = {"summe", "summen", "total", "totals", "gesamt", "insgesamt", "gesamtsumme", "zusammen", "sum", "zwischensumme",
           "subtotal", "totalsumme", "endsumme", "gesamttotal"} | woerter.TOTAL
# Eine Zeile „Durchschnitt“ oder „Anteil gesamt“ unter den Teilen ist keine Summe der Teile.
_NOT_TOTALS = {"mittelwert", "durchschnitt", "schnitt", "ø", "median", "anteil", "anteile", "quote", "rate", "prozent",
               "max", "min", "maximum", "minimum", "mean", "average", "avg", "%", "stabw", "std"} | woerter.MEANS | \
              {"part", "quota", "share", "pourcentage", "percentuale", "proportion"}
_ROUNDS = {"ROUND", "ROUNDUP", "ROUNDDOWN"}
# Was eine Summenzeile so rechnet, ist kein Zusammenzählen der Teile (auch gerundet).
_NOT_ADDING = {"AVERAGE", "AVERAGEA", "AVERAGEIF", "AVERAGEIFS", "MEDIAN", "MIN", "MAX", "MINA", "MAXA", "MINIFS", "MAXIFS",
               "STDEV", "STDEV.S", "STDEV.P", "STDEVP", "STDEVA", "STDEVPA", "VAR", "VAR.S", "VAR.P", "VARP", "MODE", "MODE.SNGL",
               "GEOMEAN", "HARMEAN", "PERCENTILE", "PERCENTILE.INC", "PERCENTILE.EXC", "QUARTILE", "QUARTILE.INC",
               "QUARTILE.EXC", "TRIMMEAN", "LARGE", "SMALL", "RANK", "RANK.EQ", "RANK.AVG"}


def _averaging(node: Node | None) -> bool:
    while node is not None and node.kind == "call" and node.value in _ROUNDS and node.args:
        node = node.args[0]
    return node is not None and node.kind == "call" and node.value in _NOT_ADDING


def _from_totals(node: Node | None, line: "_Line") -> bool:
    """=E14/C14 in der Summenzeile: ein Verhältnis aus den Summen daneben (Abweichung in % des Budgets). Das ist keine
    Summe der Prozente darüber, und sie soll es auch nicht sein."""
    refs, divides, stack = [], False, [node] if node is not None else []
    while stack:
        n = stack.pop()
        if n.kind in ("range", "xref", "xrange", "unknown"):
            return False
        if n.kind == "op" and n.value == "/" and len(n.args) == 2 and n.args[1].kind != "num":
            divides = True
        if n.kind == "ref" and isinstance(n.value, tuple):
            refs.append(n.value)
        stack.extend(n.args or [])
    axis = 0 if line.vertical else 1
    # =D28/($B$5*B28): Die Bettenzahl oben bei den Annahmen gehört nicht zu den Teilen, sie ist eine Vorgabe. Nur
    # ein Bezug mitten in die Teile hinein zeigt, dass hier aus den Zeilen darüber gerechnet wird.
    lo, hi = min(p[axis] for p in line.parts), max(p[axis] for p in line.parts)
    return divides and any(p[axis] == line.total[axis] for p in refs) and \
        all(p[axis] == line.total[axis] or not lo <= p[axis] <= hi for p in refs)


# „Jahr“ rechts neben Q1 bis Q4 oder den Monaten: die Summe des Jahres, nur dort. Sonst ist „Jahr“ eine Spalte mit
# Jahreszahlen oder ein Mittelwert (Auslastung im Jahr).
_PERIODS = {"jahr", "year", "fy", "année", "annee", "anno", "ganzjahr", "jahreswert", "annual"}


def _period(text: str) -> bool:
    tokens = _words(text)
    return bool(tokens) and any(t in _PERIODS for t in tokens) and all(t in _PERIODS or _YEARS.fullmatch(t) for t in tokens)


def _is_total(text: str) -> bool:
    tokens = _words(text)
    if "%" in text or any(t in _NOT_TOTALS for t in tokens):
        return False
    return any(t in _TOTALS or (t.startswith("gesamt") and len(t) > 6) or t.endswith("summe") for t in tokens)


def _figure(sheet: Sheet, pos: Pos) -> bool:
    """Eine Zahl oder eine Formel, kein Kopf. Je Zelle einmal gerechnet: Große Tabellen fragen jede Zelle oft."""
    memo = sheet.memo.setdefault("summen.figure", {})
    try:
        return memo[pos]
    except KeyError:
        pass
    cell = sheet.cells.get(pos)
    if cell is None or cell.value.strip() == "" and cell.formula is None or _head(sheet, pos):
        found = False
    else:
        found = cell.formula is not None or _number_at(sheet, pos)
    memo[pos] = found
    return found


@dataclass
class _Line:
    """Eine Summe und ihre Teile: die Zahlen darüber (vertical) oder links daneben."""

    total: Pos
    parts: list[Pos]
    heads: list[str]  # wie jeder Teil heißt (Zeilenkopf oder Spaltenkopf)
    label: str  # wie die Summe heißt
    vertical: bool
    across: str = ""  # in einer Übersicht: der Kopf der Spalte (oder Zeile), in der die Summe steht
    # Woran er sieht, dass das Teile eines Ganzen sind: der Kopf darüber (across), ein gemeinsames Wort (stem),
    # Namen derselben Art wie Regionen einer Spalte oder Monate (family) oder nur die Formel der Summe (formula).
    evidence: str = "across"
    period: bool = False  # „Jahr“ neben den Quartalen: eine Summe nur, wenn die Formel zusammenzählt


def _texts(sheet: Sheet) -> dict[int, set[str]]:
    """Die Texte jeder Spalte: Daran erkennt er „Nord“, „Ost“ und „Süd“ als Werte derselben Spalte."""
    out: dict[int, set[str]] = {}
    for (_, c), cell in sheet.cells.items():
        if cell.formula is None and cell.value.strip():
            out.setdefault(c, set()).add(cell.value.strip().casefold())
    return out


def _kind_of(head: str) -> str:
    tokens = _words(head)
    if len(tokens) == 1 and (tokens[0] in _MONTHS or tokens[0] in _QUARTERS):
        return "monat"
    return "jahr" if re.fullmatch(r"(19|20)\d\d", head.strip()) else ""


def _family(heads: list[str], texts: dict[int, set[str]], own: int | None):
    """Wie weit unten die Namen von derselben Art sind (Monate, Jahre, Werte einer Spalte), oder None."""
    if len(heads) < 2:
        return None
    a, b = heads[-1], heads[-2]
    if _kind_of(a) and _kind_of(a) == _kind_of(b):
        kind = _kind_of(a)
        return lambda h: _kind_of(h) == kind
    for col, values in texts.items():
        if col != own and a.strip().casefold() in values and b.strip().casefold() in values:
            return lambda h, values=values: h.strip().casefold() in values
    return None


class _Runs:
    """Wo eine Reihe von Zahlen und Formeln anfängt, nach oben oder nach links: einmal gerechnet, nicht je Zelle neu."""

    def __init__(self, sheet: Sheet):
        self.sheet = sheet
        self.starts: dict[tuple[Pos, bool], int] = {}

    def start(self, pos: Pos, leftwards: bool) -> int:
        """Die erste Zeile (oder Spalte) der Reihe, in der pos steht; pos selbst, wenn davor nichts ist."""
        r, c = pos
        walked = []
        while True:
            key = ((r, c), leftwards)
            if key in self.starts:
                first = self.starts[key]
                break
            walked.append(key)
            before = (r, c - 1) if leftwards else (r - 1, c)
            if min(before) < 1 or not _figure(self.sheet, before):
                first = c if leftwards else r
                break
            r, c = before
        for key in walked:
            self.starts[key] = first
        return first

    def filled(self, pos: Pos, first: int, leftwards: bool) -> bool:
        """Stehen von first bis pos (in derselben Zeile oder Spalte) nur Zahlen und Formeln?"""
        last = pos[1] if leftwards else pos[0]
        return last < first or _figure(self.sheet, pos) and self.start(pos, leftwards) <= first


_YEARS = re.compile(r"(?<!\d)((?:19|20)\d\d)(?!\d)")


def _counter(sheet: Sheet, col: int, row: int) -> bool:
    """Die Spalte der Zeilennamen zählt nur durch (Jahr 0, 1, 2 …, Nr. 1, 2, 3 …): ein Name, kein Teil einer Summe."""
    plan = sheet.plan
    if col != plan.name_column or plan.header_row is None or plan.last_row is None:
        return False
    if not plan.header_row < row <= plan.last_row:
        return False
    memo = sheet.memo.setdefault("summen.counter", {})
    if col not in memo:
        values = [sheet.inputs.get((r, col)) for r in range(plan.header_row + 1, plan.last_row + 1)]
        steps = [b - a for a, b in zip(values, values[1:]) if isinstance(a, float) and isinstance(b, float)]
        memo[col] = len(values) >= 3 and len(steps) == len(values) - 1 and all(s == 1 for s in steps)
    return memo[col]


def _long_name(sheet: Sheet, pos: Pos) -> str:
    """Ein langer Zeilenname („Klebsiella pneumoniae (Carbapenem-resistent)“): als Name einer Zeile so gut wie ein
    kurzer. Kein Satz (mit Punkt am Ende) und keine Formel."""
    cell = sheet.cells.get(pos)
    if cell is None or cell.formula is not None:
        return ""
    value = (cell.value or "").strip()
    if not 40 < len(value) <= 80 or not any(ch.isalpha() for ch in value) or value.endswith((".", "!", "?")):
        return ""
    return value


@lru_cache(maxsize=4096)
def _total_label(text: str) -> bool:
    return _is_total(text)


def _line(sheet: Sheet, pos: Pos, vertical: bool, texts: dict | None = None,  # noqa: C901 - Kopf, Teile, Name
          runs: _Runs | None = None) -> _Line | None:
    r, c = pos
    runs = runs or _Runs(sheet)
    # Der Kopf der Summe: links in derselben Zeile (Summenzeile) oder oben in derselben Spalte (Summenspalte).
    k = runs.start(pos, vertical) - 1
    if k < 1:
        return None
    head_pos = (r, k) if vertical else (k, c)
    # Viele Zahlen teilen sich denselben Kopf (die ganze Zeile, die ganze Spalte): ihn einmal ansehen.
    heads = sheet.memo.setdefault("summen.total_head", {})
    label = heads.get(head_pos)
    if label is None:
        label = _head(sheet, head_pos)
        label = heads[head_pos] = label if label and (_total_label(label) or _period(label)) else ""
    if not label:
        return None
    period = not _total_label(label)
    if period and pos not in sheet.nodes:
        return None
    parts, heads = [], []
    step = r - 1 if vertical else c - 1
    while step >= 1:
        part = (step, c) if vertical else (r, step)
        if not _figure(sheet, part):
            break
        if not vertical and _counter(sheet, step, r):
            break
        name = _head(sheet, (step, k) if vertical else (k, step)) or (_long_name(sheet, (step, k)) if vertical else "")
        # Zwischen dem Kopf der Teile und der Summe stehen nur Zahlen (die Spalte unter „Maus“ ist voll).
        if not name or not runs.filled((step, c - 1) if vertical else (r - 1, step), k + 1, vertical):
            break
        parts.append(part)
        heads.append(name)
        step -= 1
    parts.reverse()
    heads.reverse()
    if len(parts) < 2 or any(_is_total(h) for h in heads):
        return None
    if period and not all(_kind_of(h) == "monat" for h in heads):
        return None
    # „Summe 2026-2030“ zählt nur diese Jahre, nicht das Ist-Jahr 2025 davor; „Total 2024“ nur die Quartale von 2024.
    years = [int(y) for y in _YEARS.findall(label)]
    if years:
        def inside(head: str) -> bool:
            found = _YEARS.findall(head)
            return bool(found) and all(min(years) <= int(y) <= max(years) for y in found)

        chosen = [i for i, h in enumerate(heads) if inside(h)]
        if chosen and len(chosen) < len(heads):
            parts, heads = [parts[i] for i in chosen], [heads[i] for i in chosen]
            if len(parts) < 2:
                return None
    # Der Kopf über den Teilen (die Spalte „Maus“ einer Übersicht) oder ein gemeinsames Wort in den Namen
    # („Umsatz Nord“, „Umsatz Ost“): Sonst stehen da verschiedene Dinge untereinander, keine Teile eines Ganzen.
    above = (parts[0][0] - 1, c) if vertical else (r, parts[0][1] - 1)
    across = _head(sheet, above) if min(above) >= 1 else ""
    if vertical and across and woerter.rate_word(_words(across)):
        return None  # Bruttomarge je Bereich: Die Margen zusammengezählt sind keine Marge.
    if vertical and sheet.plan.across:
        # In einem Finanzplan steht über einer Spalte ein Jahr: Darunter stehen verschiedene Größen (Erlöse, Kosten),
        # nicht Teile eines Ganzen. Teile sind dort nur, was der Name der Summe oder die Formel sagt.
        across = ""
    evidence = "across"
    if across:
        # „Summe Kosten“ unter Erlöse, Raumkosten, Personal, Sachkosten: Was vor (oder nach) den gezählten Zeilen
        # steht und nicht nach dem Namen der Summe heißt, ist eine andere Größe, kein vergessener Teil.
        measure = _measure(label)
        node = sheet.nodes.get(pos)
        weights = _weights(sheet, node) if measure and node is not None else None
        summed = [i for i, p in enumerate(parts) if weights is not None and weights.get(p)]
        if summed:
            def named(head: str) -> bool:
                return any(m in w for w in _words(head) for m in measure)

            chosen = [i for i in range(len(parts)) if summed[0] <= i <= summed[-1] or named(heads[i])]
            if len(chosen) < len(parts):
                parts, heads = [parts[i] for i in chosen], [heads[i] for i in chosen]
                if len(parts) < 2:
                    return None
    if not across:
        measure = _measure(label)
        stem = measure or (set(_words(heads[-1])) & set(_words(heads[-2]))) - _FILLER
        family = None if stem else _family(heads, texts if texts is not None else _texts(sheet), k if vertical else None)
        if stem:
            fits, evidence = (lambda h: stem <= set(_words(h))), "stem"
            # „Umsatz Nord“, „Umsatz Ost“: Ohne das gemeinsame Wort bleiben Werte derselben Spalte, Teile derselben Art.
            rest = [" ".join(t for t in _words(h) if t not in stem) for h in heads[-2:]]
            if all(rest) and _family(rest, texts if texts is not None else _texts(sheet), None):
                evidence = "family"
        elif family:
            fits, evidence = family, "family"
        else:
            # Ohne gemeinsamen Namen zählen nur die Zellen, die die Formel selbst zusammenzählt, und was dazwischen liegt.
            node = sheet.nodes.get(pos)
            weights = _weights(sheet, node) if node is not None else None
            summed = [i for i, p in enumerate(parts) if weights is not None and weights.get(p)]
            if len(summed) < 2:
                return None
            first = summed[0]
            fits, evidence = (lambda h: True), "formula"
            parts, heads = parts[first:], heads[first:]
        keep = len(heads)
        while keep and fits(heads[keep - 1]):
            keep -= 1
        parts, heads = parts[keep:], heads[keep:]
        if len(parts) < 2:
            return None
    return _Line(pos, parts, heads, label, vertical, across, evidence, period)


def _measure(label: str) -> set[str]:
    """Was eine Summe zählt, nach ihrem Namen: „Summe Kosten“ → kosten. Ein zusammengesetztes Wort zählt mit seinem
    zweiten Teil: „Gesamterlös“ → erlös, „Gesamtkosten“ → kosten. Sonst hielte der Prüfer bei „Gesamterlös =
    KV-Umsatz + Privatumsatz“ auch Fälle, Punkte und Punktwert links davon für vergessene Teile."""
    out = set()
    for t in _words(label):
        if t in _TOTALS or t in _FILLER or t in _PERIODS or _YEARS.fullmatch(t):
            continue
        prefix = next((p for p in _COMPOUND_TOTALS if t.startswith(p)), None)
        if prefix is not None:
            rest = t[len(prefix):]
            if len(rest) >= 3 and rest not in _TOTALS and rest not in _FILLER:
                out.add(rest)
            continue
        out.add(t)
    return out


_COMPOUND_TOTALS = ("gesamt", "total", "summe", "insgesamt")


def _lines(sheet: Sheet) -> list[_Line]:
    found = []
    texts, runs = _texts(sheet), _Runs(sheet)
    for pos in sorted(sheet.cells):
        if not _figure(sheet, pos) or pos in sheet.unknown:
            continue
        for vertical in (True, False):
            line = _line(sheet, pos, vertical, texts, runs)
            if line is not None:
                found.append(line)
    return found


# --- Was die Summe zusammenzählt ---


class _Weights:
    """Mit welchem Faktor jede Zelle in die Summe eingeht (2: doppelt, -1: abgezogen), feste Zahlen unter None.

    Ausgezählt wird erst, wenn eine Prüfung alle Zellen braucht: Eine Laufsumme =SUMME($E$2:E9000) hätte sonst in
    jeder Zeile Tausende Einträge, obwohl meist nur gefragt ist, wie einzelne Zellen zählen."""

    def __init__(self, sheet: Sheet, node: Node, cells: dict, areas: list) -> None:
        self._sheet, self._node, self._cells, self._areas = sheet, node, cells, areas
        self._all: dict | None = None

    def get(self, pos, default=None):
        if pos is None:
            return self.full().get(None, default)
        weight = self._cells.get(pos, 0.0)
        if any(r1 <= pos[0] <= r2 and c1 <= pos[1] <= c2 for (r1, c1, r2, c2), _ in self._areas) and (
                _number_at(self._sheet, pos) or pos in self._sheet.nodes):
            weight += sum(f for (r1, c1, r2, c2), f in self._areas if r1 <= pos[0] <= r2 and c1 <= pos[1] <= c2)
        return weight if weight else default

    def full(self) -> dict:
        if self._all is None:
            self._all = _counted(self._sheet, self._node)
        return self._all

    def __iter__(self):
        return iter(self.full())

    def __len__(self) -> int:
        return len(self.full())

    def __contains__(self, pos) -> bool:
        return pos in self.full()

    def items(self):
        return self.full().items()

    def values(self):
        return self.full().values()


def _weights(sheet: Sheet, node: Node) -> _Weights | None:
    """Wie die Formel ihre Zellen zusammenzählt, None für eine andere Rechnung (Produkt, Mittelwert, Bedingung)."""
    cells: dict = {}
    areas: list = []
    return _Weights(sheet, node, cells, areas) if _terms(sheet, node, 1.0, cells, areas) else None


def _terms(sheet: Sheet, node: Node, factor: float, cells: dict, areas: list, depth: int = 0) -> bool:
    """Die Zellen und Bereiche der Summe mit ihrem Faktor; False, wo _counted None ergibt."""
    if depth > 20:
        return False
    if node.kind == "ref":
        cells[node.value] = cells.get(node.value, 0.0) + factor
        return True
    if node.kind == "range":
        r1, c1, r2, c2 = node.value
        rows, cols = sheet.extent(None)
        areas.append(((r1, c1, min(r2, rows), min(c2, cols)), factor))
        return True
    if node.kind == "num":
        return True
    if node.kind == "neg":
        return _terms(sheet, node.args[0], -factor, cells, areas, depth + 1)
    if node.kind == "call" and node.value == "SUM":
        return all(_terms(sheet, a, factor, cells, areas, depth + 1) for a in node.args)
    if node.kind == "call" and node.value in _ROUNDS and node.args:
        return _terms(sheet, node.args[0], factor, cells, areas, depth + 1)
    if node.kind == "op" and node.value in ("+", "-") and len(node.args) == 2:
        return (_terms(sheet, node.args[0], factor, cells, areas, depth + 1)
                and _terms(sheet, node.args[1], factor if node.value == "+" else -factor, cells, areas, depth + 1))
    return False


def _counted(sheet: Sheet, node: Node, depth: int = 0) -> dict | None:
    """_Weights ausgezählt: jede Zelle mit ihrem Faktor, in der Reihenfolge der Formel."""
    if depth > 20:
        return None
    if node.kind == "ref":
        return {node.value: 1.0}
    if node.kind == "range":
        r1, c1, r2, c2 = node.value
        rows, cols = sheet.extent(None)
        return {p: 1.0 for p in sheet._inside(r1, c1, min(r2, rows), min(c2, cols)) if _number_at(sheet, p) or p in sheet.nodes}  # noqa: SLF001
    if node.kind == "num":
        return {None: node.value}
    if node.kind == "neg":
        inner = _counted(sheet, node.args[0], depth + 1)
        return None if inner is None else {k: -v for k, v in inner.items()}
    if node.kind == "call" and node.value == "SUM":
        return _combine([_counted(sheet, a, depth + 1) for a in node.args], [1.0] * len(node.args))
    if node.kind == "call" and node.value in _ROUNDS and node.args:
        return _counted(sheet, node.args[0], depth + 1)
    if node.kind == "op" and node.value in ("+", "-") and len(node.args) == 2:
        return _combine([_counted(sheet, a, depth + 1) for a in node.args], [1.0, 1.0 if node.value == "+" else -1.0])
    return None


def _combine(parts: list, signs: list[float]) -> dict | None:
    if any(p is None for p in parts):
        return None
    out: dict = {}
    for part, sign in zip(parts, signs, strict=True):
        for k, v in part.items():
            out[k] = out.get(k, 0.0) + sign * v
    return {k: v for k, v in out.items() if v}


# --- Hinweise ---


def _value(sheet: Sheet, calc, pos: Pos):
    node = sheet.nodes.get(pos)
    if node is None:
        return sheet.inputs.get(pos)
    try:
        return calc.run(node)
    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError):
        return None


def _number(value) -> bool:
    return isinstance(value, float) and not isinstance(value, bool) and math.isfinite(value)


def _span(parts: list[Pos]) -> str:
    return f"{address(parts[0])}:{address(parts[-1])}"


def _named(line: _Line, pos: Pos, values: dict) -> str:
    from .workbook import like_cell

    head = line.heads[line.parts.index(pos)]
    # „C9 („O1“, 4509591,62)“: Ein Teil, der wie eine Zelle heißt, steht in Anführungszeichen.
    return f"{address(pos)} ({T(f'„{head}“', f'“{head}”') if like_cell(head) else head}, {_fmt(values[pos])})"


def _listing(items: list[str]) -> str:
    return items[0] if len(items) == 1 else T(f"{', '.join(items[:-1])} und {items[-1]}", f"{', '.join(items[:-1])} and {items[-1]}")


def _title(line: _Line) -> str:
    """„„Umsatz gesamt“ in H6“ oder „Die Summe der Spalte „Maus“ in H8“."""
    where = address(line.total)
    if line.across and _is_total(line.across):
        return T(f"Die Gesamtsumme in {where}", f"The grand total in {where}")
    if line.across and not (set(_words(line.label)) - _TOTALS - _FILLER - (_PERIODS if line.period else set())):
        kind = "Spalte" if line.vertical else "Zeile"
        return T(f"Die Summe der {kind} „{line.across}“ in {where}",
                 f"The total of {'column' if line.vertical else 'row'} “{line.across}” in {where}")
    return T(f"„{line.label}“ in {where}", f"“{line.label}” in {where}")


def _wrongly(line: _Line, values: dict, missing: list[Pos], doubled: list[Pos], minus: list[Pos]) -> list[str]:
    """Was die Summe falsch zählt, als Satzteile nach „zählt aber …“ (Englisch: „but it …“)."""
    what = []
    if missing:
        listed = _listing([_named(line, p, values) for p in missing])
        what.append(T(f"{listed} nicht mit", f"leaves out {listed}"))
    if doubled:
        listed = _listing([_named(line, p, values) for p in doubled])
        what.append(T(f"{listed} doppelt", f"counts {listed} twice"))
    if minus:
        listed = _listing([_named(line, p, values) for p in minus])
        what.append(T(f"{listed} abgezogen statt dazu", f"subtracts {listed} instead of adding {'it' if len(minus) == 1 else 'them'}"))
    return what


def _should(line: _Line, what: list[str]) -> str:
    """„„Summe“ in C7 soll die 4 Teile C3:C6 zusammenzählen, zählt aber …“"""
    return T(f"{_title(line)} soll die {len(line.parts)} Teile {_span(line.parts)} zusammenzählen, zählt aber {_listing(what)}.",
             f"{_title(line)} should add up the {len(line.parts)} parts {_span(line.parts)}, but it {_listing(what)}.")


def _fixed_tip(fixed: str, expected: float) -> str:
    return T(f"Korrigiert: ={fixed}, nachgerechnet {_fmt(expected)}.", f"Fixed: ={fixed}, recalculated this gives {_fmt(expected)}.")


def _tolerance(sheet: Sheet, pos: Pos, expected: float) -> float:
    node = sheet.nodes.get(pos)
    step = _rounding(node) if node is not None else 0.0
    if node is None:
        # Eine getippte Summe darf so gerundet sein, wie sie dasteht (12 000 für 11 998,40 ist eine andere Zahl).
        shown = (sheet.cells[pos].value or "").strip()
        decimals = len(re.split(r"[.,]", shown)[-1]) if re.search(r"[.,]\d+$", shown) else 0
        step = 10 ** -decimals if decimals else 0.0
    return step / 2 + 1e-9 * max(1.0, abs(expected))


def _row_text(sheet: Sheet, pos: Pos) -> str:
    """Die Texte links in der Zeile und über der Spalte: woran man eine Kontrollzeile oder -spalte erkennt."""
    texts = [sheet.cells[(pos[0], c)].value for c in range(1, pos[1]) if (pos[0], c) in sheet.cells
             and sheet.cells[(pos[0], c)].formula is None and isinstance(sheet.cells[(pos[0], c)].value, str)]
    return " ".join(texts)


_CONTROL = re.compile(r"kontrolle|abweichung|differenz|\bcheck|prüfung|abgleich|muss 0|= ?ok|0 = ok|control|reconcil", re.I)


def _footing(sheet: Sheet, line: _Line, calc, siblings: bool) -> tuple[Finding, float] | None:  # noqa: C901
    """Die Summe stimmt nicht mit ihren Teilen: welcher Teil fehlt, doppelt zählt oder nicht dazugehört.
    siblings: Andere Summen derselben Summenzeile (oder -spalte) zählen genau ihre Teile."""
    node = sheet.nodes.get(line.total)
    # „Mittlere Liegedauer“ in der Zeile „Summe“: der Mittelwert über alle, nicht die Summe der Mittelwerte.
    if _averaging(node) or line.across and set(_words(line.across)) & _MEANS or _from_totals(node, line):
        return None
    weights = _weights(sheet, node) if node is not None else None
    # „Kontrolle: Länder + nicht zugeordnet − Gesamt (0 = OK)“: eine Prüfung, die eine Summe abzieht, keine Summe.
    if weights is not None and any(v < 0 for v in weights.values()) and _CONTROL.search(_row_text(sheet, line.total)):
        return None
    # „Jahr“ = M8 (Kontostand am Jahresende) oder =M5-B5 (Veränderung übers Jahr) ist keine Summe der Monate: Nur was
    # mindestens zwei Teile zusammenzählt und nichts abzieht, prüft er gegen alle Monate.
    if line.period and (weights is None or sum(1 for p in set(line.parts) if weights.get(p)) < 2
                        or any(v < 0 for v in weights.values())):
        return None
    # Nur eine Summe ist eine Summe der Teile: „Summe“ über =B2*C2 (Betrag mal Satz) rechnet etwas anderes. Eine
    # feste Zahl oder eine andere Rechnung prüft er in einer Übersicht nur, wenn die Nachbarn zeigen, dass dort
    # Summen stehen.
    if weights is None and (line.across or node is not None) and not siblings:
        return None
    if weights is not None and not siblings and sum(1 for p in set(line.parts) if weights.get(p)) < 2:
        return None
    values = {p: _value(sheet, calc, p) for p in line.parts}
    if not all(_number(v) for v in values.values()):
        return None
    got = _value(sheet, calc, line.total)
    if not _number(got):
        return None
    expected = math.fsum(values.values())
    if abs(got - expected) <= _tolerance(sheet, line.total, expected):
        return _unseen(sheet, line, calc, weights, values, expected, siblings)
    fixed = _fix(sheet, line, calc, expected)
    if not fixed:
        return None
    title, span = _title(line), _span(line.parts)
    if weights is not None:
        missing = [p for p in line.parts if not weights.get(p) and values[p]]
        doubled = [p for p in line.parts if weights.get(p, 0) > 1 and values[p]]
        minus = [p for p in line.parts if weights.get(p, 0) < 0 and values[p]]
        foreign = [p for p in weights if p is not None and p not in line.parts and p != line.total
                   and _number(_value(sheet, calc, p)) and _value(sheet, calc, p)]
        fixed_number = weights.get(None)
        if not (missing or doubled or minus or foreign or fixed_number):
            # Zählt die Summe genau ihre Teile, kommt der Unterschied aus einem Teil: Das meldet dessen Prüfung.
            return None
        if minus and not siblings and line.evidence != "family":
            # =B5-C5 kann gewollt sein (Brutto minus Rabatt); falsch ist es, wenn die Nachbarn alles zusammenzählen
            # oder die Teile von derselben Art sind (Regionen, Monate).
            return None
        what = _wrongly(line, values, missing, doubled, minus)
        if foreign:
            listed = _listing([sheet.label(p) for p in foreign[:3]])
            what.append(T(f"{listed} dazu, {'das' if len(foreign) == 1 else 'die'} "
                          f"nicht zu den Teilen {'gehört' if len(foreign) == 1 else 'gehören'}",
                          f"adds {listed}, which {'is' if len(foreign) == 1 else 'are'} not one of the parts"))
        if fixed_number:
            what.append(T(f"die feste Zahl {_fmt(fixed_number)} dazu", f"adds the fixed number {_fmt(fixed_number)}"))
        text = _should(line, what) + T(f" So ergibt {address(line.total)} {_fmt(got)} statt {_fmt(expected)}.",
                                       f" So {address(line.total)} gives {_fmt(got)} instead of {_fmt(expected)}.")
    else:
        typed = node is None
        text = T(f"{title} {'ist eine feste Zahl' if typed else 'ergibt'} {_fmt(got)}, die {len(line.parts)} Teile {span} "
                 f"ergeben zusammen {_fmt(expected)}.",
                 f"{title} {'is the fixed number' if typed else 'gives'} {_fmt(got)} while the {len(line.parts)} parts "
                 f"{span} add up to {_fmt(expected)}.")
        gap = expected - got
        same = [p for p in line.parts if abs(abs(values[p]) - abs(gap)) <= 1e-9 * max(1.0, abs(gap))]
        if len(same) == 1:
            p = same[0]
            text += T(f" Der Unterschied ist genau {_named(line, p, values)}{', als fehle der Teil' if gap > 0 else ''}.",
                      f" The difference is exactly {_named(line, p, values)}{', as if that part were missing' if gap > 0 else ''}.")
    return Finding(f"parts:{address(line.total)}", text, "fehler", _fixed_tip(fixed, expected)), expected


def _unseen(sheet: Sheet, line: _Line, calc, weights: dict | None, values: dict, expected: float,
            siblings: bool) -> tuple[Finding, float] | None:
    """Die Summe stimmt nur, weil der ausgelassene (doppelte, abgezogene) Teil gerade 0 ist."""
    if weights is None or not (siblings or line.evidence in ("stem", "family")):
        return None
    missing = [p for p in line.parts if not weights.get(p)]
    doubled = [p for p in line.parts if weights.get(p, 0) > 1]
    minus = [p for p in line.parts if weights.get(p, 0) < 0]
    if not (missing or doubled or minus) or weights.get(None) or any(p not in line.parts for p in weights if p is not None):
        return None
    fixed = _fix(sheet, line, calc, expected)
    if not fixed:
        return None
    what = _wrongly(line, values, missing, doubled, minus)
    cells = missing + doubled + minus
    one = len(cells) == 1
    text = _should(line, what) + T(
        f" Weil {'dort' if one else 'in diesen Zellen'} gerade 0 steht, stimmt {address(line.total)} noch. "
        f"Sobald sich {'der Wert' if one else 'die Werte'} ändert, ist die Summe falsch.",
        f" Because {'that cell' if one else 'these cells'} currently {'contains' if one else 'contain'} 0, "
        f"{address(line.total)} is still right. As soon as {'the value changes' if one else 'the values change'}, "
        f"the total is wrong.")
    return Finding(f"parts:{address(line.total)}", text, "pruefen", _fixed_tip(fixed, expected)), expected


def _fix(sheet: Sheet, line: _Line, calc, expected: float) -> str:
    """=SUMME(H2:H5) über alle Teile, wie Excel sie braucht, auf dem Blatt nachgerechnet."""
    english = f"SUM({address(line.parts[0])}:{address(line.parts[-1])})"
    node = sheet.nodes.get(line.total)
    if node is not None:
        # Rundet die Summe, rundet auch die Korrektur.
        for call in formel.calls(node):
            if call.value in _ROUNDS and call is node and len(call.args) == 2 and call.args[1].kind == "num":
                english = f"{call.value}({english},{int(call.args[1].value)})"
    try:
        got = calc.run(formel.parse(english))
    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
        return ""
    if not _number(got) or abs(got - expected) > _tolerance(sheet, line.total, expected):
        return ""
    return _german_formula(english)


# --- Erfassen die Teile alle Zeilen der Tabelle? ---


def _similar(a: str, b: str) -> bool:
    """„Nrod“ und „Nord“, „Nordd“ und „Nord“: vertippt, nicht eine andere Gruppe."""
    a, b = a.strip().casefold(), b.strip().casefold()
    if a == b or min(len(a), len(b)) < 3:
        return a == b
    if sorted(a) == sorted(b):
        return True
    if abs(len(a) - len(b)) > 1:
        return False
    # Höchstens ein Zeichen anders, zu viel oder zu wenig.
    previous = list(range(len(b) + 1))
    for i, x in enumerate(a, start=1):
        row = [i]
        for j, y in enumerate(b, start=1):
            row.append(min(previous[j] + 1, row[j - 1] + 1, previous[j - 1] + (x != y)))
        previous = row
    return previous[-1] <= 1


def _spelled(label: str) -> str:
    """„Gesamtumsatz“ heißt „gesamt umsatz“, „Umsatzsumme“ heißt „umsatz summe“."""
    out = []
    for token in _words(label):
        if token.startswith("gesamt") and len(token) > 8:
            out += ["gesamt", token[6:]]
        elif token.endswith("summe") and len(token) > 8 and token not in _TOTALS:
            out += [token[:-5], "summe"]
        else:
            out.append(token)
    return " ".join(out)


def _groups(sheet: Sheet, line: _Line, calc, rows: list[int], columns: dict) -> Finding | None:  # noqa: C901 - eine Ursache je Abschnitt
    """Die Teile rechnen je Gruppe aus der Tabelle (SUMMEWENN je Region), die Summe heißt „gesamt“: Fehlen Zeilen?"""
    if not line.vertical or not all(p in sheet.nodes for p in line.parts):
        return None
    total_name = _spelled(f"{line.label} {line.across}")
    whole = _wanted(total_name, sheet, columns)
    pieces = [_wanted(f"{h} {line.across}".strip(), sheet, columns) for h in line.heads]
    if whole is None or any(p is None for p in pieces):
        return None
    split = None
    for piece in pieces:
        extra = [col for col in piece.conditions if col not in whole.conditions]
        if len(extra) != 1 or len(piece.conditions[extra[0]]) != 1 or piece.function not in (None, "SUM"):
            return None
        if {c: v for c, v in piece.conditions.items() if c != extra[0]} != whole.conditions:
            return None
        if (piece.months, piece.year) != (whole.months, whole.year) or split not in (None, extra[0]):
            return None
        split = extra[0]
    targets = set.union(*(p.targets for p in pieces)) | whole.targets or _summed(sheet, calc, line.parts)
    numeric = [col for col in targets if columns.get(col) is not None and columns[col].kind == "zahl"]
    if len(numeric) != 1 or whole.function not in (None, "SUM"):
        return None
    target = numeric[0]
    numbers = columns[target].numbers
    values = {p: _value(sheet, calc, p) for p in line.parts}
    got = _value(sheet, calc, line.total)
    if not all(_number(v) for v in values.values()) or not _number(got):
        return None
    # Rechnen die Teile, was ihr Name sagt? Sonst meldet das bereiche.py für den Teil.
    for part, piece in zip(line.parts, pieces, strict=True):
        own = math.fsum(numbers[r] for r in numbers if piece.matches(columns, r))
        if abs(own - values[part]) > 1e-6 * max(1.0, abs(own)):
            return None
    everything = [r for r in numbers if whole.matches(columns, r)]
    expected = math.fsum(numbers[r] for r in everything)
    if abs(got - expected) <= _tolerance(sheet, line.total, expected) or abs(got - math.fsum(values.values())) > \
            _tolerance(sheet, line.total, got):
        return None
    texts = columns[split].texts
    names = {next(iter(p.conditions[split])) for p in pieces}
    shown = {next(iter(p.conditions[split])): next(iter(p.conditions[split].values())) for p in pieces}
    left = [r for r in everything if texts.get(r, "").strip().casefold() not in names]
    doubled = sorted({v for v in names if [next(iter(p.conditions[split])) for p in pieces].count(v) > 1})
    header = sheet.plan.headers.get(split, "")
    by_value: dict[str, list[int]] = {}
    for r in left:
        by_value.setdefault(texts.get(r, "").strip(), []).append(r)
    causes, fixes, overrides = [], [], {}
    typos = 0
    for value, where in sorted(by_value.items(), key=lambda kv: kv[1][0]):
        amount = math.fsum(numbers[r] for r in where)
        cells = _listing([address((r, split)) for r in where[:3]] +
                         ([T(f"{len(where) - 3} weitere", f"{len(where) - 3} more")] if len(where) > 3 else []))
        near = [shown[n] for n in names if value and _similar(value, n)]
        one = len(where) == 1
        if not value:
            causes.append(T(f"in {cells} {'ist' if one else 'sind'} „{header}“ leer ({_fmt(amount)}), und "
                            f"{'diese Zeile zählt' if one else 'diese Zeilen zählen'} in keinem Teil",
                            f"in {cells}, “{header}” is empty ({_fmt(amount)}), and "
                            f"{'this row is' if one else 'these rows are'} not counted in any part") if header else
                          T(f"{cells} {'ist' if one else 'sind'} leer ({_fmt(amount)}), und "
                            f"{'diese Zeile zählt' if one else 'diese Zeilen zählen'} in keinem Teil",
                            f"{cells} {'is' if one else 'are'} empty ({_fmt(amount)}), and "
                            f"{'this row is' if one else 'these rows are'} not counted in any part"))
            fixes.append(T(f"Trag in {cells} {'die passende' if header else 'den passenden Wert'}{f' „{header}“' if header else ''} ein.",
                           f"Enter the right {f'“{header}”' if header else 'value'} in {cells}."))
        elif len(near) == 1:
            causes.append(T(f"in {cells} steht „{value}“, wohl vertippt für „{near[0]}“ ({_fmt(amount)})",
                            f"{cells} {'contains' if one else 'contain'} “{value}”, probably a typo for “{near[0]}” "
                            f"({_fmt(amount)})"))
            typos += 1
            fixes.append(T(f"Schreib in {cells} „{near[0]}“ statt „{value}“.", f"Write “{near[0]}” instead of “{value}” in {cells}."))
            overrides |= {(r, split): near[0] for r in where}
        else:
            causes.append(T(f"{header} „{value}“ hat keine eigene Zeile ({len(where)} {'Zeile' if one else 'Zeilen'}, "
                            f"zusammen {_fmt(amount)})",
                            f"{header} “{value}” has no row of its own ({len(where)} {'row' if one else 'rows'}, "
                            f"together {_fmt(amount)})"))
            fixes.append(_new_part(sheet, line, calc, value, amount))
    if doubled:
        causes.append(T(f"{_listing([f'„{shown[v]}“' for v in doubled])} {'zählt' if len(doubled) == 1 else 'zählen'} zweimal",
                        f"{_listing([f'“{shown[v]}”' for v in doubled])} {'is' if len(doubled) == 1 else 'are'} counted twice"))
    if not causes or any(not f for f in fixes):
        return None
    text = T(f"{_title(line)} zählt die {len(line.parts)} Teile {_span(line.parts)} zusammen, aber die Teile erfassen nicht "
             f"alle Zeilen der Tabelle: {_listing(causes)}. Darum ergibt {address(line.total)} {_fmt(got)} statt "
             f"{_fmt(expected)}.",
             f"{_title(line)} adds up the {len(line.parts)} parts {_span(line.parts)}, but the parts do not cover "
             f"all rows of the table: {_listing(causes)}. So {address(line.total)} gives {_fmt(got)} instead of "
             f"{_fmt(expected)}.")
    tip = " ".join(fixes)
    if overrides and len(fixes) == typos:
        after = _after(sheet, line, overrides)
        if after is None or abs(after - expected) > _tolerance(sheet, line.total, expected):
            return None
        tip += T(f" Dann ergibt {address(line.total)} {_fmt(expected)} (nachgerechnet).",
                 f" Then {address(line.total)} gives {_fmt(expected)} (recalculated).")
    elif not doubled and not whole.conditions and whole.date_column is None:
        direct = _direct(calc, target, rows, expected, _tolerance(sheet, line.total, expected))
        if direct:
            tip += T(f" Oder rechne {address(line.total)} direkt aus der Tabelle: ={direct} (nachgerechnet {_fmt(expected)}).",
                     f" Or calculate {address(line.total)} directly from the table: ={direct} "
                     f"(recalculated this gives {_fmt(expected)}).")
    return Finding(f"groups:{address(line.total)}", text, "fehler", tip.strip())


def _direct(calc, target: int, rows: list[int], expected: float, tolerance: float) -> str:
    """=SUMME(D:D), wenn in der Spalte sonst nichts steht, sonst =SUMME(D2:D41): auf dem Blatt nachgerechnet."""
    column = column_letters(target)
    for english in (f"SUM({column}:{column})", f"SUM({column}{min(rows)}:{column}{max(rows)})"):
        try:
            got = calc.run(formel.parse(english))
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
            continue
        if _number(got) and abs(got - expected) <= tolerance:
            return _german_formula(english)
    return ""


_SUM_RANGE = {"SUMIF": 2, "SUMIFS": 0}


def _summed(sheet: Sheet, calc, parts: list[Pos]) -> set[int]:
    """Heißen die Teile nur „Nord“ und „Ost“: die Spalte, die ihre SUMMEWENN zusammenzählen."""
    found = set()
    for part in parts:
        node = sheet.nodes.get(part)
        if node is None or node.kind != "call" or node.value not in _SUM_RANGE or len(node.args) < 2:
            return set()
        arg = node.args[_SUM_RANGE[node.value]] if node.value == "SUMIFS" or len(node.args) == 3 else node.args[0]
        try:
            where, area = calc.area_of(arg)
        except formel.Unknown:
            return set()
        if where is not None or area[1] != area[3]:
            return set()
        found.add(area[1])
    return found if len(found) == 1 else set()


def _new_part(sheet: Sheet, line: _Line, calc, value: str, amount: float) -> str:
    """Eine Zeile mehr: die Formel des letzten Teils mit dem fehlenden Wert, auf dem Blatt nachgerechnet."""
    last = line.parts[-1]
    formula = sheet.cells[last].formula or ""
    texts = re.findall(r'"((?:[^"]|"")*)"', formula)
    head = line.heads[-1]
    known = [t for t in texts if t.casefold() in head.casefold()]
    if len(known) != 1:
        return ""
    english = formula.replace(f'"{known[0]}"', '"' + value.replace('"', '""') + '"', 1)
    try:
        got = calc.run(formel.parse(english))
    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
        return ""
    if not _number(got) or abs(got - amount) > 1e-6 * max(1.0, abs(amount)):
        return ""
    name = head.replace(known[0], value) if known[0] in head else f"{head} {value}"
    return T(f"Füg unter {address((last[0], last[1] - 1))} eine Zeile „{name}“ mit ={_german_formula(english)} ein "
             f"(nachgerechnet {_fmt(amount)}) und nimm sie in die Summe {address(line.total)} auf.",
             f"Insert a row “{name}” below {address((last[0], last[1] - 1))} with ={_german_formula(english)} "
             f"(recalculated this gives {_fmt(amount)}) and include it in the total {address(line.total)}.")


def _after(sheet: Sheet, line: _Line, overrides: dict) -> float | None:
    chain = set(line.parts) | {line.total}
    _, value = sheet.evaluator(overrides=overrides, chain=chain)
    try:
        got = value(line.total)
    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
        return None
    return got if _number(got) else None


def _exactly(sheet: Sheet, line: _Line) -> bool:
    """Zählt die Formel genau die Teile dieser Linie, jeden einmal?"""
    node = sheet.nodes.get(line.total)
    weights = _weights(sheet, node) if node is not None else None
    return (weights is not None and all(weights.get(p) == 1.0 for p in line.parts) and set(weights) == set(line.parts)
            and all(v == 1.0 for v in weights.values()))


def _matches(sheet: Sheet, line: _Line, calc) -> bool:
    values = [_value(sheet, calc, p) for p in line.parts]
    got = _value(sheet, calc, line.total)
    if not _number(got) or not all(_number(v) for v in values):
        return False
    expected = math.fsum(values)
    return abs(got - expected) <= _tolerance(sheet, line.total, expected)


def checks(sheet: Sheet, rows: list[int] | None = None, columns: dict | None = None) -> tuple[list[Finding], set[str]]:
    """Die Hinweise und die Zellen (Adressen), deren allgemeine Hinweise sie genauer sagen."""
    calc, _ = sheet.evaluator()
    found: list[Finding] = []
    explained: set[str] = set()
    by_total: dict[Pos, list[_Line]] = {}
    every = _lines(sheet)
    for line in every:
        by_total.setdefault(line.total, []).append(line)
    exact = {(ln.total, ln.vertical) for ln in every if _exactly(sheet, ln)}
    exact_at: dict[tuple[bool, int], set[Pos]] = {}
    for total, vertical in exact:
        exact_at.setdefault((vertical, total[0] if vertical else total[1]), set()).add(total)
    for total, lines in by_total.items():
        if len(found) >= MAX_FINDINGS:
            break
        node = sheet.nodes.get(total)
        if len(lines) == 2:
            # Die Gesamtsumme in der Ecke darf die Zeilensummen oder die Spaltensummen zusammenzählen. Passt sie zu
            # einer der beiden, steckt ein Unterschied zur anderen in einer einzelnen Summe: Die meldet ihre Prüfung.
            # Passt sie nur dem Wert nach, kann ein Teil fehlen, der gerade 0 ist.
            if any(_exactly(sheet, line) for line in lines):
                continue
            lines = [line for line in lines if _matches(sheet, line, calc)] or lines
            weights = _weights(sheet, node) if node is not None else None
            lines = sorted(lines, key=lambda ln: -sum(1 for p in set(ln.parts) if weights is not None and weights.get(p)))[:1]
        line = lines[0]
        # Summen nebeneinander in einer Summenzeile, untereinander in einer Summenspalte.
        siblings = bool(exact_at.get((line.vertical, total[0] if line.vertical else total[1]), set()) - {total})
        footing = _footing(sheet, line, calc, siblings)
        if footing is not None:
            found.append(footing[0])
            explained.add(address(total))
            continue
        if rows and columns:
            groups = _groups(sheet, line, calc, rows, columns)
            if groups is not None:
                found.append(groups)
                explained.add(address(total))
    return found, explained
