"""Bereiche, die nicht zu dem passen, was sie rechnen sollen.

Eine Kennzahl heißt „Umsatz alle Nord“, ihre SUMME nimmt aber nur 9 der 10 Nord-Zeilen oder schon eine
Ost-Zeile mit, egal ob die Tabelle sortiert ist. Eine feste Zahl daneben passt nicht mehr zu ihren Zeilen.
SUMMEWENN prüft die Bedingung in B2:B31 und nimmt die Zahlen aus D3:D32, eine Zeile verschoben, oder prüft nur
einen Teil der Tabelle. Die Bedingung „Nord“ findet „Nord “ mit Leerzeichen nicht, oder ist vertippt und
ergibt 0. Und ein Bereich, der genau bei einer Gruppe einer sortierten Tabelle anfängt, aber eine Zeile zu
früh oder zu spät aufhört. Und Bereiche ohne $, die beim Herunterziehen mitrutschen (D2:D31, D3:D32 …).

Der Prüfer sagt genau, was anders ist (welche Zeile fehlt, welche zu viel ist und warum, durch welche Zahl
geteilt wird), und schlägt eine Formel nur vor, wenn sein Formelrechner sie auf dem Blatt nachgerechnet hat.

Alles ohne KI: Der Prüfer liest nur, was im Namen der Kennzahl steht (Werte aus den Textspalten, Monate, Jahre,
Spaltenköpfe, Wörter wie Summe oder Mittelwert), und rechnet selbst nach. Steht dort ein Wort, das er nicht
kennt („Ziel Umsatz Nord“, „Provision Nord“), sagt er lieber nichts.
"""

from __future__ import annotations

import difflib
import itertools
import math
import re
import statistics
from bisect import bisect_left, bisect_right
from dataclasses import dataclass, field
from datetime import date, timedelta

from . import formel, sprache, woerter
from .analyse import Finding, Sheet, _fmt, _german, address
from .sprache import T
from .workbook import _as_number, _left_label, pattern
from .xlsx import column_letters, column_number, number_text, outside_literals, shift

Pos = tuple[int, int]
Area = tuple[int, int, int, int]

MAX_EACH = 3
MAX_LISTED = 3
_STATS = {
    "summe": "SUM", "sum": "SUM", "total": "SUM", "gesamt": "SUM", "insgesamt": "SUM", "gesamtsumme": "SUM",
    "mittelwert": "AVERAGE", "durchschnitt": "AVERAGE", "schnitt": "AVERAGE", "ø": "AVERAGE", "mean": "AVERAGE",
    "average": "AVERAGE", "avg": "AVERAGE", "median": "MEDIAN", "anzahl": "COUNT", "count": "COUNT",
    "max": "MAX", "maximum": "MAX", "höchster": "MAX", "höchste": "MAX", "min": "MIN", "minimum": "MIN",
    "tiefster": "MIN", "tiefste": "MIN", "niedrigster": "MIN", "niedrigste": "MIN", **woerter.STATS,
}
# Eine Spalte von Mittelwerten (Mittlere Liegedauer, Median, Quote) zählt in der Summenzeile nicht zusammen: Dort
# steht der Mittelwert über alle, oder die Quote aus den Summen.
_MEANS = {"mittelwert", "mittelwerte", "mittlere", "mittlerer", "mittleres", "mittel", "durchschnitt", "durchschnittlich",
          "durchschnittliche", "durchschnittlicher", "durchschnittliches", "schnitt", "ø", "median", "mean", "average", "avg"} | woerter.MEANS
_NOT_SUMMED = _MEANS | {"max", "maximum", "höchster", "höchste", "min", "minimum", "tiefster", "tiefste", "niedrigster",
                        "niedrigste", "stabw", "std", "quote", "rate"}
# Wörter, die sagen: über die ganze Spalte, nicht nur eine Gruppe.
_ALL = {"alle", "aller", "allen", "alles", "gesamt", "insgesamt", "total", "summe", "gesamtsumme", "all", "overall"} | woerter.ALL
_FILLER = {"der", "die", "das", "des", "den", "dem", "von", "vom", "in", "im", "für", "fuer", "aus", "und", "nur",
           "je", "pro", "zu", "the", "of", "for", "and", "only", "eur", "euro", "chf", "fr", "usd"} | woerter.FILLER
_CURRENCIES = {"eur", "euro", "chf", "fr", "usd"}
_MONTHS = {
    "januar": 1, "jänner": 1, "jan": 1, "january": 1, "februar": 2, "feb": 2, "february": 2, "märz": 3, "maerz": 3,
    "mär": 3, "mrz": 3, "march": 3, "april": 4, "apr": 4, "mai": 5, "may": 5, "juni": 6, "jun": 6, "june": 6,
    "juli": 7, "jul": 7, "july": 7, "august": 8, "aug": 8, "september": 9, "sep": 9, "sept": 9, "oktober": 10,
    "okt": 10, "october": 10, "oct": 10, "november": 11, "nov": 11, "dezember": 12, "dez": 12, "december": 12, "dec": 12,
}
_MONTH_NAMES = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober",
                "November", "Dezember"]
_MONTH_NAMES_EN = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
                   "November", "December"]
_QUARTERS = {f"q{q}": {3 * q - 2, 3 * q - 1, 3 * q} for q in range(1, 5)}
_PLAIN = {"SUM", "AVERAGE", "COUNT", "COUNTA", "MIN", "MAX", "MEDIAN"}
_ROUNDING = {"ROUND", "ROUNDUP", "ROUNDDOWN"}
_CONDITIONAL = {"SUMIF", "COUNTIF", "AVERAGEIF", "SUMIFS", "COUNTIFS", "AVERAGEIFS", "MINIFS", "MAXIFS"}


def _words(text: str) -> list[str]:
    return re.findall(r"[^\W_]+", woerter.spaced(text).casefold())


def _num(value: float) -> str:
    """Geldbeträge ganz: 26688,97, nicht auf sechs Stellen gerundet. Ein krummes Ergebnis auf vier Stellen."""
    return number_text(round(value, 4 if abs(value) >= 1 else 6))


def _more(n: int) -> str:
    """„ und 3 weitere“ nach einer gekürzten Liste."""
    return T(f" und {n} weitere", f" and {n} more")


def _fn(name: str) -> str:
    """Eine Funktion im Satz: SUMMEWENN, auf Englisch SUMIF."""
    return T(_german(name), name)


def _cells(col: int, rows: list[int]) -> str:
    shown = ", ".join(address((r, col)) for r in rows[:MAX_LISTED])
    return shown + (_more(len(rows) - MAX_LISTED) if len(rows) > MAX_LISTED else "")


def _area(area: Area) -> str:
    r1, c1, r2, c2 = area
    if r2 >= formel.MAX_ROW:
        return f"{column_letters(c1)}:{column_letters(c2)}"
    return f"{address((r1, c1))}:{address((r2, c2))}"


def checks(sheet: Sheet, rows: list[int], columns: dict) -> list[Finding]:
    """columns: je Spalte der Tabelle, was in ihr steht (auswertung._Info: numbers, dates, texts, kind)."""
    found: list[Finding] = []
    # Je Zelle ein Hinweis, der genaueste zuerst: Bereiche verschoben oder Bedingung vertippt, dann was der Name
    # der Kennzahl sagt, zuletzt die feste Anzahl. Kann der Name nicht sagen, was falsch ist, die Anzahl aber
    # schon, gilt die Anzahl.
    counts = {f.key.partition(":")[2]: f for f in _fixed_counts(sheet)}
    named = [counts.pop(f.key.partition(":")[2]) if f.level != "fehler" and f.key.partition(":")[2] in counts else f
             for f in _labelled(sheet, rows, columns)]
    for more in (_conditional(sheet, rows, columns), named, list(counts.values())):
        taken = {f.key.partition(":")[2] for f in found}
        found += [f for f in more if f.key.partition(":")[2] not in taken]
    # Folgt ein Fehler nur aus einem anderen (H3 = H2/10, und H2 ist falsch), zählt nur der erste.
    flagged = {formel.position(f.key.partition(":")[2]) for f in found}
    return [f for f in found if not (_before(sheet, formel.position(f.key.partition(":")[2])) & flagged)]


# --- Was der Name einer Kennzahl verlangt ---


@dataclass
class _Want:
    """„Umsatz alle Nord“: Region ist Nord. „Mittelwert März 2026“: Datum im März 2026, als Mittelwert."""

    conditions: dict[int, dict[str, str]] = field(default_factory=dict)  # Spalte → Text (klein) → wie er dasteht
    months: set[int] = field(default_factory=set)
    year: int | None = None
    date_column: int | None = None
    function: str | None = None
    targets: set[int] = field(default_factory=set)
    everything: bool = False
    heads: dict[str, Pos] = field(default_factory=dict)  # In einer Kreuztabelle: Wert (klein) → seine Kopfzelle

    def matches(self, columns: dict, row: int) -> bool:
        for col, wanted in self.conditions.items():
            if columns[col].texts.get(row, "").strip().casefold() not in wanted:
                return False
        if self.date_column is not None:
            day = _day(columns[self.date_column].dates.get(row))
            if day is None or (self.months and day.month not in self.months) or (self.year and day.year != self.year):
                return False
        return True

    def describe(self, sheet: Sheet) -> str:
        parts = [T(f"{sheet.plan.headers[col]} „{'“ oder „'.join(sorted(wanted.values()))}“",
                   f"{sheet.plan.headers[col]} “{'” or “'.join(sorted(wanted.values()))}”")
                 for col, wanted in sorted(self.conditions.items())]
        if self.date_column is not None:
            when = " ".join(p for p in (_months_text(self.months), str(self.year or "")) if p)
            head = sheet.plan.headers[self.date_column]
            parts.append(T(f"{head} im {when}", f"{head} in {when}") if self.months else
                         T(f"{head} im Jahr {self.year}", f"{head} in {self.year}"))
        return T(" und ", " and ").join(parts)


def _day(serial) -> date | None:
    if not isinstance(serial, float):
        return None
    return date(1899, 12, 30) + timedelta(days=int(serial))


def _runs(numbers: list[int]) -> list[tuple[int, int]]:
    """[1, 2, 3, 7] → [(1, 3), (7, 7)]."""
    runs: list[tuple[int, int]] = []
    for n in sorted(numbers):
        if runs and n == runs[-1][1] + 1:
            runs[-1] = (runs[-1][0], n)
        else:
            runs.append((n, n))
    return runs


def _listing(items: list[str]) -> str:
    return items[0] if len(items) == 1 else T(f"{', '.join(items[:-1])} und {items[-1]}", f"{', '.join(items[:-1])} and {items[-1]}")


def _months_text(months: set[int]) -> str:
    """{1, 2, 7} → „Januar, Februar und Juli“, {1, 2, 3} → „Januar bis März“."""
    names = []
    month = T(_MONTH_NAMES, _MONTH_NAMES_EN)
    for start, end in _runs(list(months)):
        if end - start >= 2:
            names.append(T(f"{month[start - 1]} bis {month[end - 1]}", f"{month[start - 1]} to {month[end - 1]}"))
        else:
            names += [month[m - 1] for m in range(start, end + 1)]
    return _listing(names) if names else ""


def _wanted(label: str, sheet: Sheet, columns: dict) -> _Want | None:
    tokens = _words(label)
    if not tokens:
        return None
    want = _Want()
    known = set(_FILLER)
    # Verglichen wird mit Umlaut gleich ae/oe/ue: „Groesse“ im Namen meint die Spalte „Größe“.
    folded = [woerter.fold(t) for t in tokens]
    for col, info in columns.items():
        header = set(_words(sheet.plan.headers.get(col, "")))
        if {woerter.fold(w) for w in header} & set(folded):
            known |= header | {t for t, f in zip(tokens, folded) if f in {woerter.fold(w) for w in header}}
            if info.kind == "zahl":
                want.targets.add(col)
        if info.kind != "text":
            continue
        for value in {v.strip() for v in info.texts.values() if v.strip()}:
            words = _words(value)
            if not words or set(words) <= _FILLER | set(_STATS) | _ALL:
                continue
            same = [woerter.fold(w) for w in words]
            if any(folded[i:i + len(same)] == same for i in range(len(tokens) - len(same) + 1)):
                want.conditions.setdefault(col, {})[value.casefold()] = value
                known |= set(words) | {t for t, f in zip(tokens, folded) if f in same}
    dated = [col for col, info in columns.items() if info.kind == "datum"]
    for token in tokens:
        if token in _STATS:
            want.function = want.function or _STATS[token]
            known.add(token)
        if token in _ALL:
            want.everything = True
            known.add(token)
        if token in _MONTHS or token in _QUARTERS:
            want.months |= {_MONTHS[token]} if token in _MONTHS else _QUARTERS[token]
            known.add(token)
        if re.fullmatch(r"(19|20)\d\d", token):
            want.year = int(token)
            known.add(token)
    if want.months or want.year:
        if len(dated) != 1:
            return None
        want.date_column = dated[0]
    if any(t not in known for t in tokens):
        return None
    if not (want.conditions or want.date_column is not None or want.everything):
        return None
    return want


def _not_itself(sheet: Sheet, want: _Want | None, row: int, label: str) -> _Want | None:
    """Ein Name ist nicht seine eigene Gruppe: Steht „Kennzahl 2“ in einer Spalte, die die Tabelle als Kategorie
    liest, meint J2 daneben nicht die Zeilen mit „Kennzahl 2“, denn das ist nur J2s eigene Zeile."""
    if want is None:
        return None
    for col, wanted in list(want.conditions.items()):
        cell = sheet.cells.get((row, col))
        if cell is not None and cell.value.strip().casefold() == label.strip().casefold() in wanted:
            del want.conditions[col]
    return want if want.conditions or want.date_column is not None or want.everything else None


def _aggregate(function: str, values: list[float]) -> float:
    if function == "SUM":
        return math.fsum(values)
    if function == "AVERAGE":
        return math.fsum(values) / len(values)
    if function in ("COUNT", "COUNTA"):
        return float(len(values))
    if function == "MEDIAN":
        return float(statistics.median(values))
    return max(values) if function == "MAX" else min(values)


def _decimals(value: float) -> int:
    text = repr(value)
    if "e" in text or "." not in text:
        return 0
    tail = text.split(".")[1]
    return 0 if tail == "0" else len(tail)


# --- Kreuztabellen: der Name kommt aus Zeile und Spalte ---
#
# In einer Übersicht (Regionen untereinander, Produkte nebeneinander) heißt G5 nicht nur „Nord“, wie links
# daneben steht, sondern „Nord“ und „Maus“ aus dem Kopf darüber. Mit nur einem der beiden sähe jede richtige
# SUMMEWENNS($D:$D;$B:$B;$F5;$C:$C;G$4) falsch aus.

_YEAR_HEAD = re.compile(r"(19|20)\d\d")
# Nur Ziffern und Satzzeichen einer Zahl: sicher kein Buchstabe darin (schnell, die meisten Zellen sind Zahlen).
NO_LETTERS = re.compile(r"[\d\s.,'’%+\-]*")


def _head(sheet: Sheet, pos: Pos) -> str:
    """Ein Kopf: ein Wort oder eine Jahreszahl, keine Formel."""
    cell = sheet.cells.get(pos)
    if cell is None or cell.formula is not None:
        return ""
    value = (cell.value or "").replace(":", "").strip()
    if NO_LETTERS.fullmatch(value):
        return value if _YEAR_HEAD.fullmatch(value) and _year_head(sheet, pos) else ""
    return value if (any(ch.isalpha() for ch in value) and len(value) <= 40) or _YEAR_HEAD.fullmatch(value) else ""


def _year_head(sheet: Sheet, pos: Pos) -> bool:
    """Ist die Jahreszahl ein Kopf? Ja in einer Kopfzeile aus Jahren (2025, 2026 …) oder einer Spalte „Jahr“. Nein,
    wenn Zeile und Spalte sonst Zahlen haben: Dann sind 1990 Franken im August ein Betrag wie jeder andere."""
    memo = sheet.memo.get("bereiche.amounts")
    if memo is None:
        rows: set[int] = set()
        cols: set[int] = set()
        for (r, c), cell in sheet.cells.items():
            value = (cell.value or "").strip()
            if cell.formula is None and value and NO_LETTERS.fullmatch(value) and not _YEAR_HEAD.fullmatch(value) \
                    and _as_number(value) is not None:
                rows.add(r)
                cols.add(c)
        memo = sheet.memo["bereiche.amounts"] = {"rows": rows, "cols": cols}
    return pos[0] not in memo["rows"] or pos[1] not in memo["cols"]


def _figure(sheet: Sheet, pos: Pos) -> bool:
    cell = sheet.cells.get(pos)
    if cell is None or _head(sheet, pos):
        return False
    return cell.formula is not None or _number_at(sheet, pos)


def _cross(sheet: Sheet, r: int, c: int) -> list[tuple[str, Pos]] | None:
    """Liegt (r, c) in einer Kreuztabelle mit mindestens zwei Spalten? Dann ihre Köpfe: (Ecke,) Zeile, Spalte."""
    lo, hi = c, c
    while lo > 1 and _figure(sheet, (r, lo - 1)):
        lo -= 1
    while _figure(sheet, (r, hi + 1)):
        hi += 1
    top = r
    while top > 1 and _figure(sheet, (top - 1, c)):
        top -= 1
    if hi == lo or lo < 2 or top < 2:
        return None
    row_head, col_head = _head(sheet, (r, lo - 1)), _head(sheet, (top - 1, c))
    if not row_head or not col_head or any(not _head(sheet, (top - 1, cc)) for cc in range(lo, hi + 1)):
        return None
    parts = [(row_head, (r, lo - 1)), (col_head, (top - 1, c))]
    corner = _head(sheet, (top - 1, lo - 1))
    return ([(corner, (top - 1, lo - 1))] if corner else []) + parts


def _cross_wanted(sheet: Sheet, parts, columns: dict) -> tuple[str, _Want] | None:
    """Der Name aus den Köpfen. Die Ecke („Umsatz“) gehört dazu; versteht er sie nicht, darf sie nur Spaltenköpfe
    nennen („Region / Produkt“). „Anteil“ in der Ecke heißt: Hier stehen keine Summen."""
    def named(use):
        want = _wanted(" ".join(text for text, _ in use), sheet, columns)
        if want is not None:
            for text, pos in use:
                want.heads.setdefault(text.casefold(), pos)
        return want

    full = named(parts)
    if len(parts) == 2:
        return (" / ".join(text for text, _ in parts), full) if full else None
    known = set(_FILLER) | {w for col in columns for w in _words(sheet.plan.headers.get(col, ""))}
    bare = named(parts[1:]) if set(_words(parts[0][0])) <= known else None
    # Nennt die Ecke nur die Spalten („Region / Produkt“), steht sie nicht im Namen.
    if bare and (full is None or (full.targets, full.function) == (bare.targets, bare.function)):
        return " / ".join(text for text, _ in parts[1:]), bare
    return (" / ".join(text for text, _ in parts), full) if full else None


# --- Welche Zeilen eine Formel nimmt ---


@dataclass
class _Use:
    """Welche Zeilen der Zahlenspalte eine Formel nimmt. kind: sum (jede Zeile mit ihrem Faktor, auch doppelt
    oder abgezogen), mean, count, min, max, median (über eine Menge von Zeilen) oder ratio (eine Summe geteilt
    durch eine feste Zahl oder durch die Anzahl einer Menge)."""

    kind: str
    rows: dict[int, float]
    divisor: float | None = None
    counted: list[int] | None = None


_KIND = {"AVERAGE": "mean", "COUNT": "count", "COUNTA": "count", "MIN": "min", "MAX": "max", "MEDIAN": "median",
         "SUMIF": "sum", "SUMIFS": "sum", "AVERAGEIF": "mean", "AVERAGEIFS": "mean", "COUNTIF": "count",
         "COUNTIFS": "count", "MINIFS": "min", "MAXIFS": "max"}
_FUNCTION = {"sum": "SUM", "mean": "AVERAGE", "ratio": "AVERAGE", "count": "COUNT", "min": "MIN", "max": "MAX",
             "median": "MEDIAN"}


def _number_at(sheet: Sheet, pos: Pos) -> bool:
    v = sheet.inputs.get(pos)
    return isinstance(v, float) and not isinstance(v, bool)


def _range_rows(sheet: Sheet, area: Area, target: int) -> list[int] | None:
    """Die Zeilen mit einer Zahl in der Zahlenspalte, die ein Bereich nimmt (auch eine Summenzeile darunter)."""
    r1, c1, r2, c2 = area
    if c1 != target or c2 != target:
        return None
    bottom = min(r2, sheet.extent(None)[0])
    return [r for r in range(r1, bottom + 1) if _number_at(sheet, (r, target))]


def _picked(sheet: Sheet, call, target: int) -> list[int] | None:
    """Die Zeilen, die SUMMEWENN und Verwandte wirklich nehmen, so wie Excel die Bereiche nebeneinanderlegt."""
    parts = _ranges_of(call)
    if parts is None:
        return None
    sums, tests = parts
    if any(a.kind != "range" for a in sums + [t[0] for t in tests]):
        return None
    try:
        checks = [(t[0].value, formel._criterion(_criterion_value(sheet, t[1]))) for t in tests]
    except formel.Unknown:
        return None
    counting = call.value in ("COUNTIF", "COUNTIFS")
    source = (sums[0] if sums else tests[0][0]).value
    if not counting and (source[1] != target or source[3] != target):
        return None
    bottom = sheet.extent(None)[0]
    picked = []
    for i in range(min(source[2], bottom) - source[0] + 1):
        if all(test(sheet.inputs.get((area[0] + i, area[1]))) for area, test in checks):
            row = source[0] + i
            if counting or _number_at(sheet, (row, target)):
                picked.append(row)
    return picked


def _use(sheet: Sheet, node, target: int, data: set[int], depth: int = 0) -> _Use | None:
    if node is None or depth > 8:
        return None
    if node.kind == "call" and node.value in _ROUNDING and node.args:
        return _use(sheet, node.args[0], target, data, depth)
    if node.kind == "ref":
        pos = node.value
        if pos in sheet.nodes and not (pos[0] in data and pos[1] == target):
            return None if pos in sheet.unknown else _use(sheet, sheet.nodes[pos], target, data, depth + 1)
        if pos[1] == target:
            return _Use("sum", {pos[0]: 1.0} if _number_at(sheet, pos) else {})
        return None
    if node.kind == "neg":
        inner = _use(sheet, node.args[0], target, data, depth + 1)
        return _Use("sum", {r: -k for r, k in inner.rows.items()}) if inner and inner.kind == "sum" else None
    if node.kind == "call":
        name = node.value
        if name == "SUM":
            total: dict[int, float] = {}
            for arg in node.args:
                part = _range_rows(sheet, arg.value, target) if arg.kind == "range" else None
                if part is not None:
                    found = {r: 1.0 for r in part}
                else:
                    inner = _use(sheet, arg, target, data, depth + 1)
                    if inner is None or inner.kind != "sum":
                        return None
                    found = inner.rows
                for r, k in found.items():
                    total[r] = total.get(r, 0.0) + k
            return _Use("sum", total)
        if name in ("AVERAGE", "COUNT", "COUNTA", "MIN", "MAX", "MEDIAN"):
            rows: set[int] = set()
            for arg in node.args:
                if arg.kind == "range":
                    part = _range_rows(sheet, arg.value, target)
                elif arg.kind == "ref" and arg.value[1] == target and arg.value not in sheet.nodes:
                    part = [arg.value[0]] if _number_at(sheet, arg.value) else []
                else:
                    part = None
                if part is None:
                    return None
                rows |= set(part)
            return _Use(_KIND[name], {r: 1.0 for r in rows})
        if name in _CONDITIONAL:
            picked = _picked(sheet, node, target)
            return None if picked is None else _Use(_KIND[name], {r: 1.0 for r in picked})
        return None
    if node.kind != "op" or len(node.args) != 2:
        return None
    left, right = node.args
    if node.value in ("+", "-"):
        a, b = _use(sheet, left, target, data, depth + 1), _use(sheet, right, target, data, depth + 1)
        if a is None or b is None or a.kind != "sum" or b.kind != "sum":
            return None
        total = dict(a.rows)
        for r, k in b.rows.items():
            total[r] = total.get(r, 0.0) + (k if node.value == "+" else -k)
        return _Use("sum", {r: k for r, k in total.items() if k})
    if node.value == "*":
        factor, other = (left, right) if left.kind == "num" else (right, left)
        if factor.kind != "num":
            return None
        inner = _use(sheet, other, target, data, depth + 1)
        return _Use("sum", {r: k * factor.value for r, k in inner.rows.items()}) if inner and inner.kind == "sum" else None
    if node.value == "/":
        top = _use(sheet, left, target, data, depth + 1)
        if top is None or top.kind != "sum":
            return None
        if right.kind == "num" and right.value:
            return _Use("ratio", top.rows, divisor=right.value)
        bottom = _use(sheet, right, target, data, depth + 1)
        if bottom is not None and bottom.kind == "count":
            return _Use("ratio", top.rows, counted=sorted(bottom.rows))
    return None


# --- Was die Formel anders nimmt, in Worten ---


def _ident(sheet: Sheet, columns: dict, want: _Want, row: int, why: bool) -> str:
    """„D31 (27.07.2026)“, bei einer Zeile, die nicht dazugehört, mit dem Grund: „D12 (Ost)“."""
    from .auswertung import is_id

    bits = []
    dated = [col for col, info in columns.items() if info.kind == "datum"]
    # Eine Fallnummer nennt die Zeile besser als ein Datum, das mit dem Namen nichts zu tun hat (Geburtsdatum).
    numbered = row in sheet.plan.names and is_id(sheet.plan.headers.get(sheet.plan.name_column, ""))
    if dated and (want.date_column is not None or (not why and not numbered)):
        shown = sheet.shown.get((row, dated[0]))
        if isinstance(shown, str):
            bits.append(shown)
    if why:
        bits += [columns[col].texts.get(row, "").strip() for col in sorted(want.conditions) if columns[col].texts.get(row)]
    if not bits and row in sheet.plan.names:
        bits.append(sheet.plan.word(row))
    if not bits and sheet.plan.last_row < row:
        word = _left_label(sheet.cells, row, min(sheet.plan.headers or [1]) + 1)
        if word:
            bits.append(word)
    return ", ".join(b for b in bits if b)


def _row_list(sheet: Sheet, columns: dict, want: _Want, target: int, rows: list[int], why: bool) -> str:
    shown = []
    for r in rows[:MAX_LISTED]:
        name = _ident(sheet, columns, want, r, why)
        shown.append(f"{address((r, target))} ({name})" if name else address((r, target)))
    if len(rows) > MAX_LISTED:
        return f"{', '.join(shown)}{_more(len(rows) - MAX_LISTED)}"
    return _listing(shown)


def _spans(target: int, rows) -> str:
    """Die Zeilen als Bereiche: „D2:D9 und D28:D30“."""
    parts = [address((a, target)) if a == b else f"{address((a, target))}:{address((b, target))}" for a, b in _runs(list(rows))]
    if len(parts) > 4:
        n, low, high = len(list(rows)), address((min(rows), target)), address((max(rows), target))
        return T(f"{n} Zeilen von {low} bis {high}", f"{n} rows from {low} to {high}")
    return _listing(parts) if parts else T("keine Zeile", "no row")


def _group_name(sheet: Sheet, columns: dict, rows: set[int], data: list[int]) -> str:
    """Ist eine Menge von Zeilen genau eine Gruppe („Region „Ost““) oder alle Zeilen?"""
    if rows == set(data):
        return _all_rows()
    if not rows:
        return ""
    single = []
    for col, info in sorted(columns.items()):
        values = {info.texts.get(r, "").strip() for r in rows} if info.kind == "text" else set()
        if len(values) == 1 and (value := values.pop()):
            single.append((col, value))
    # Eine Gruppe („Region „Ost““) oder zwei zusammen („Region „Nord“ und Produkt „Maus““).
    for size in (1, 2):
        for combo in itertools.combinations(single, size):
            if {r for r in data if all(columns[col].texts.get(r, "").strip() == v for col, v in combo)} == rows:
                return T(" und ", " and ").join(_value(sheet.plan.headers[col], v) for col, v in combo)
    return ""


def _all_rows() -> str:
    return T("alle Zeilen", "all rows")


def _value(header: str, value: str) -> str:
    """„Region „Ost““, auf Englisch „Region “Ost”“."""
    return T(f"{header} „{value}“", f"{header} “{value}”")


def _n(n: int, one: str, many: str) -> str:
    """Die englische Form zur Anzahl: „is“ oder „are“, „row“ oder „rows“."""
    return one if n == 1 else many


def _explain(sheet: Sheet, columns: dict, want: _Want, function: str, target: int, wanted: list[int],
             use: _Use, data: list[int]) -> tuple[str, list[str]]:
    """Wie die Formel rechnet und was daran anders ist als der Name verlangt. ("", []) wenn sich nichts zeigen lässt."""
    taken = {r for r, k in use.rows.items() if k}
    named = _group_name(sheet, columns, taken, data)
    group = named and taken != set(data)
    rows_text = T(f"die Zeilen mit {named}", f"the rows with {named}") if group else _spans(target, taken)
    rows_dative = T(f"den Zeilen mit {named}", f"the rows with {named}") if group else rows_text
    parts = []
    if use.kind != {"SUM": "sum", "AVERAGE": "mean", "COUNT": "count", "COUNTA": "count", "MIN": "min", "MAX": "max",
                    "MEDIAN": "median"}[function] and not (use.kind == "ratio" and function == "AVERAGE"):
        said = T({"sum": "eine Summe", "mean": "einen Mittelwert", "count": "eine Anzahl", "min": "das Minimum",
                  "max": "das Maximum", "median": "den Median", "ratio": "eine Summe geteilt durch eine Zahl"}[use.kind],
                 {"sum": "a sum", "mean": "an average", "count": "a count", "min": "the minimum",
                  "max": "the maximum", "median": "the median", "ratio": "a sum divided by a number"}[use.kind])
        return (T(f"rechnet {said} über {rows_text}", f"calculates {said} over {rows_text}"),
                [T(f"der Name verlangt {_SAID[function]}", f"the name asks for {_SAID_EN[function]}")])
    missing = [r for r in wanted if not use.rows.get(r)]
    extra = sorted(r for r in taken if r not in set(wanted))
    if use.kind in ("sum", "ratio"):
        odd = sorted(r for r in wanted if use.rows.get(r) not in (None, 0.0, 1.0))
        doubled = [r for r in odd if use.rows[r] == 2.0]
        subtracted = [r for r in odd if use.rows[r] < 0]
        if doubled:
            listed, n = _row_list(sheet, columns, want, target, doubled, False), len(doubled)
            parts.append(T(f"{listed} {'zählt' if n == 1 else 'zählen'} doppelt", f"{listed} {_n(n, 'counts', 'count')} twice"))
        if subtracted:
            listed, n = _row_list(sheet, columns, want, target, subtracted, False), len(subtracted)
            parts.append(T(f"{listed} {'wird' if n == 1 else 'werden'} abgezogen", f"{listed} {_n(n, 'is', 'are')} subtracted"))
    if missing:
        listed, n = _row_list(sheet, columns, want, target, missing, False), len(missing)
        parts.append(T(f"{listed} {'fehlt' if n == 1 else 'fehlen'}", f"{listed} {_n(n, 'is', 'are')} missing"))
    if extra:
        listed, n = _row_list(sheet, columns, want, target, extra, True), len(extra)
        parts.append(T(f"{listed} {'gehört' if n == 1 else 'gehören'} nicht dazu", f"{listed} {_n(n, 'does', 'do')} not belong"))
    plain = use.kind == "sum" and all(k == 1.0 for k in use.rows.values())
    how = T({"sum": f"summiert {rows_text}" if plain else f"rechnet mit {rows_dative}", "mean": f"mittelt {rows_text}",
             "count": f"zählt {rows_text}", "min": f"nimmt das Minimum von {rows_dative}",
             "max": f"nimmt das Maximum von {rows_dative}", "median": f"nimmt den Median von {rows_dative}"}.get(use.kind, ""),
            {"sum": f"sums {rows_text}" if plain else f"calculates with {rows_dative}", "mean": f"averages {rows_text}",
             "count": f"counts {rows_text}", "min": f"takes the minimum of {rows_dative}",
             "max": f"takes the maximum of {rows_dative}", "median": f"takes the median of {rows_dative}"}.get(use.kind, ""))
    if use.kind == "ratio":
        if use.divisor is not None:
            by = T(f"die feste Zahl {_num(use.divisor)}", f"the fixed number {_num(use.divisor)}")
            if use.divisor != len(wanted):
                parts.append(T(f"geteilt wird durch {_num(use.divisor)} statt durch {len(wanted)}",
                               f"it divides by {_num(use.divisor)} instead of {len(wanted)}"))
        else:
            counted = set(use.counted or [])
            who = _group_name(sheet, columns, counted, data)
            by = T(f"die Anzahl von {who or _spans(target, counted)} ({len(counted)})",
                   f"the count of {who or _spans(target, counted)} ({len(counted)})")
            if counted != set(wanted):
                parts.append(T(f"geteilt wird durch {len(counted)} statt durch {len(wanted)}",
                               f"it divides by {len(counted)} instead of {len(wanted)}"))
        return T(f"teilt die Summe von {rows_dative} durch {by}", f"divides the sum of {rows_dative} by {by}"), parts
    return how, parts


def _identify(sheet: Sheet, columns: dict, want: _Want, function: str, target: int, wanted: list[int],
              value: float, expected: float, data: list[int], tolerance: float) -> str:
    """Woher eine Zahl kommt, die nicht stimmt: genau eine Zeile zu wenig oder zu viel, oder genau die Zahl einer
    anderen Gruppe. Geraten wird nicht: Nur was genau aufgeht, steht da."""
    numbers = columns[target].numbers
    close = lambda a, b: abs(a - b) <= tolerance + 1e-9 * max(1.0, abs(b))  # noqa: E731
    others = [r for r in data if r in numbers and r not in set(wanted)]
    if function == "SUM":
        hits = [r for r in wanted if close(value + numbers[r], expected)]
        if len(hits) == 1:
            listed = _row_list(sheet, columns, want, target, hits, False)
            return T(f"Es fehlt genau {listed}.", f"Exactly {listed} is missing.")
        hits = [r for r in others if close(value - numbers[r], expected)]
        if len(hits) == 1:
            listed = _row_list(sheet, columns, want, target, hits, True)
            return T(f"Zu viel ist genau {listed}.", f"Exactly {listed} is extra.")
    for col, info in sorted(columns.items()):
        if info.kind != "text":
            continue
        for name in sorted({t.strip() for t in info.texts.values() if t.strip()}):
            group = [r for r in data if r in numbers and info.texts.get(r, "").strip() == name]
            if not group or group == wanted:
                continue
            if close(value, _aggregate(function, [numbers[r] for r in group])):
                return T(f"Das ist genau {_SAID[function]} von {sheet.plan.headers[col]} „{name}“.",
                         f"That is exactly {_SAID_EN[function]} of {sheet.plan.headers[col]} “{name}”.")
    every = [r for r in data if r in numbers]
    if every != wanted and every and close(value, _aggregate(function, [numbers[r] for r in every])):
        return T(f"Das ist {_SAID[function]} über alle Zeilen.", f"That is {_SAID_EN[function]} over all rows.")
    if function == "AVERAGE" and value:
        total = math.fsum(numbers[r] for r in wanted)
        k = total / value
        if abs(k - round(k)) < 1e-6 and round(k) != len(wanted) and round(k) > 0:
            return T(f"Das ist die Summe dieser Zeilen geteilt durch {round(k)}, es sind aber {len(wanted)}.",
                     f"That is the sum of these rows divided by {round(k)}, but there are {len(wanted)}.")
    return ""


# --- Formeln vorschlagen, nur nachgerechnet ---


def _german_formula(english: str) -> str:
    """SUMIFS(D:D,B:B,"Nord") → SUMMEWENNS(D:D;B:B;"Nord"), Texte in Anführungszeichen bleiben."""
    out = []
    for i, part in enumerate(re.split(r'("(?:[^"]|"")*")', english)):
        if i % 2:
            out.append(part)
            continue
        part = re.sub(r"\b([A-Z][A-Z0-9.]*)\(", lambda m: _german(m.group(1)) + "(", part.replace("_xlfn.", ""))
        # Wahrheitswerte heißen im deutschen Excel WAHR und FALSCH, sonst nimmt Excel die Formel nicht an.
        part = re.sub(r"(?<![A-Za-z0-9_.!'])(TRUE|FALSE)(?![A-Za-z0-9_(!])", lambda m: "WAHR" if m.group(1) == "TRUE" else "FALSCH", part)
        out.append(re.sub(r"(?<![A-Z\d])(\d+)\.(\d+)", r"\1,\2", part.replace(",", ";")))
    return "".join(out)


def _quoted(text: str) -> str:
    return '"' + re.sub(r"([*?~])", r"~\1", text).replace('"', '""') + '"'


def _terms(want: _Want, columns: dict, data: list[int]) -> list[list[tuple[int, str]]] | None:
    """Je Teil der Formel die Bedingungen (Spalte, Bedingung). Nord oder Ost, Januar bis Februar und Juli: mehrere Teile."""
    texts = [[(col, _quoted(v)) for v in sorted(wanted.values())] for col, wanted in sorted(want.conditions.items())]
    blocks: list = [None]
    if want.date_column is not None:
        year = want.year
        if year is None:
            years = {d.year for r in data if (d := _day(columns[want.date_column].dates.get(r)))}
            if len(years) != 1:
                return None
            year = years.pop()
        runs = _runs(list(want.months)) if want.months else [(1, 12)]
        blocks = [((year, a), (year + (b == 12), b % 12 + 1)) for a, b in runs]
    terms = []
    for combo in itertools.product(*texts, blocks):
        *conditions, block = combo
        crit = list(conditions)
        if block is not None:
            (y1, m1), (y2, m2) = block
            crit += [(want.date_column, f'">="&DATE({y1},{m1},1)'), (want.date_column, f'"<"&DATE({y2},{m2},1)')]
        terms.append(crit)
    return terms if len(terms) <= 6 else None


def _compose(function: str, total: str, terms, span) -> str:
    def part(single: str, plural: str, crit) -> str:
        if not crit:
            return f"{single.removesuffix('IF')}({total})" if single != "COUNTIF" else f"COUNT({total})"
        if len(crit) == 1 and single in ("SUMIF", "AVERAGEIF"):
            return f"{single}({span(crit[0][0])},{crit[0][1]},{total})"
        if len(crit) == 1 and single == "COUNTIF":
            return f"COUNTIF({span(crit[0][0])},{crit[0][1]})"
        pairs = "".join(f",{span(col)},{k}" for col, k in crit)
        return f"{plural}({pairs[1:]})" if plural == "COUNTIFS" else f"{plural}({total}{pairs})"

    if function == "SUM":
        return "+".join(part("SUMIF", "SUMIFS", c) for c in terms)
    if function in ("COUNT", "COUNTA"):
        return "+".join(part("COUNTIF", "COUNTIFS", c) for c in terms)
    if function == "AVERAGE":
        if len(terms) == 1:
            return part("AVERAGEIF", "AVERAGEIFS", terms[0])
        return f"({'+'.join(part('SUMIF', 'SUMIFS', c) for c in terms)})/({'+'.join(part('COUNTIF', 'COUNTIFS', c) for c in terms)})"
    if function in ("MIN", "MAX"):
        each = [part(function + "IF", function + "IFS", c) if c else f"{function}({total})" for c in terms]
        each = [e.replace(f"{function}IF(", f"{function}IFS(") if e.startswith(f"{function}IF(") else e for e in each]
        return each[0] if len(each) == 1 else f"{function}({','.join(each)})"
    return ""


def _divides_by_own_row(node, row: int) -> bool:
    """Teilt die Formel durch etwas aus der eigenen Zeile (=D16/($B$5*B16))? Dann ist jede Zeile eine eigene Quote."""
    stack = [node]
    while stack:
        n = stack.pop()
        if n is None:
            continue
        if n.kind == "op" and n.value == "/" and len(n.args) == 2:
            inner = [n.args[1]]
            while inner:
                m = inner.pop()
                if m.kind == "ref" and isinstance(m.value, tuple) and m.value[0] == row:
                    return True
                inner.extend(a for a in (m.args or []) if isinstance(a, formel.Node))
        stack.extend(a for a in (n.args or []) if isinstance(a, formel.Node))
    return False


def row_rates(sheet: Sheet, col: int, rows: list[int]) -> list[Pos]:
    """Die Zellen einer Spalte, die je Zeile eine Quote rechnen (Auslastung = Pflegetage / (Betten * Tage)): alle mit
    derselben Formel, jede durch einen Wert ihrer eigenen Zeile geteilt. Ihre Summe sagt nichts, und ihr Mittelwert ist
    nicht die Quote über alle Zeilen. Ein Anteil am Total (=B5/$B$12) teilt durch dieselbe Zelle: Der summiert zu 100 %."""
    memo = sheet.memo.setdefault("bereiche.row_rates", {})
    if col in memo:
        return memo[col]
    cells = [(r, col) for r in rows if (r, col) in sheet.nodes and sheet.nodes[(r, col)] is not None
             and sheet.cells[(r, col)].formula]
    found: list[Pos] = []
    if len(cells) >= 3 and len(cells) >= 0.8 * len(rows):
        same = pattern(sheet.cells[cells[0]].formula, cells[0])
        if all(pattern(sheet.cells[p].formula, p) == same for p in cells) and \
                _divides_by_own_row(sheet.nodes[cells[0]], cells[0][0]):
            found = cells
    memo[col] = found
    return found


def verified(sheet: Sheet, english: list[str], expected: float, pos: Pos | None = None, tolerance: float = 0.0) -> str:
    """Die erste Formel, die auf diesem Blatt nachgerechnet genau expected ergibt (bis auf RUNDEN), auf Deutsch.
    Sonst ""."""
    calc, _ = sheet.evaluator()
    for text in english:
        try:
            node = formel.parse(text)
            if pos is not None:
                singles, areas, _ = formel.references(node, sheet.name)
                if pos in singles or any(r1 <= pos[0] <= r2 and c1 <= pos[1] <= c2 for r1, c1, r2, c2 in areas):
                    continue
            got = calc.run(node)
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
            continue
        if isinstance(got, float) and not isinstance(got, bool) and \
                abs(got - expected) <= tolerance + 1e-9 * max(1.0, abs(expected)):
            return _german_formula(text)
    return ""


def _suggest(sheet: Sheet, want: _Want, function: str, target: int, pos: Pos, columns: dict, data: list[int],
             expected: float) -> str:
    terms = _terms(want, columns, data)
    if terms is None or function == "MEDIAN":
        return ""
    candidates = []
    for whole in (True, False):
        def span(col: int, whole=whole) -> str:
            letter = column_letters(col)
            return f"{letter}:{letter}" if whole else f"${letter}${data[0]}:${letter}${data[-1]}"
        formula = _compose(function, span(target), terms, span)
        if formula:
            candidates.append(formula)
    return verified(sheet, candidates, expected, pos)


# --- Die Ursache in der Formel: die kleinste Änderung, die sie richtig macht ---
#
# Statt nur zu sagen, welche Zeilen fehlen, sucht der Prüfer die Stelle in der Formel, die falsch ist: Er probiert
# je eine kleine Änderung (eine Bedingung weg, ein anderer Wert, ein Bereich eine Zeile länger, ein Teiler, eine
# Funktion …), rechnet jede mit dem eigenen Formelrechner nach und nimmt die, die genau den richtigen Wert ergibt.
# Das ist die Ursache, und die geänderte Formel ist die Korrektur, nachgerechnet.


@dataclass
class _Edit:
    kind: str
    node: object  # die ganze geänderte Formel
    info: dict
    size: int = 0  # wie groß die Änderung ist (Zeilen bei Bereichen), kleiner zuerst


# Was zuerst gilt, wenn mehrere Änderungen passen: +D5 doppelt erklärt mehr als ein um eine Zeile kürzerer Bereich.
_ORDER = ["function", "divisor", "remove", "value", "critref", "critcol", "add", "column", "bound", "drop", "missing", "end",
          "start", "sign", "operator", "date"]
_IFS_FROM = {"SUMIFS": 1, "AVERAGEIFS": 1, "MINIFS": 1, "MAXIFS": 1, "COUNTIFS": 0}
_IF_SINGLE = {"SUMIF", "AVERAGEIF", "COUNTIF"}
_SWAPS = {"SUM": ("AVERAGE", "COUNT", "MIN", "MAX", "MEDIAN"), "AVERAGE": ("SUM", "MEDIAN", "COUNT", "MIN", "MAX"),
          "COUNT": ("SUM", "AVERAGE"), "MEDIAN": ("AVERAGE", "SUM"), "MIN": ("MAX", "AVERAGE"), "MAX": ("MIN", "AVERAGE"),
          "SUMIF": ("AVERAGEIF",), "AVERAGEIF": ("SUMIF",), "SUMIFS": ("AVERAGEIFS", "MINIFS", "MAXIFS"),
          "AVERAGEIFS": ("SUMIFS", "MINIFS", "MAXIFS"), "MINIFS": ("MAXIFS", "AVERAGEIFS"), "MAXIFS": ("MINIFS", "AVERAGEIFS")}
_FUNCTION_OF = {"SUMIF": "SUM", "SUMIFS": "SUM", "AVERAGEIF": "AVERAGE", "AVERAGEIFS": "AVERAGE", "MINIFS": "MIN",
                "MAXIFS": "MAX"}


def _kept(formula: str) -> dict:
    """Wie die Bezüge im Original dastehen, mit $: (Art, Wert) → Text."""
    kept = {}
    try:
        for kind, text in formel.tokens(formula.lstrip("=")):
            if kind == "ref":
                kept.setdefault(("ref", formel.position(text.replace("$", ""))), text)
            elif kind in ("range", "whole"):
                kept.setdefault(("range", formel.area(text.replace("$", ""))), text)
            elif kind == "ext":
                # Bezüge auf andere Blätter (Tarife!$A$2:$B$4) genau so, wie sie dastehen.
                sheet, ref = formel._sheet_part(text)  # noqa: SLF001
                if sheet is not None:
                    plain = ref.replace("$", "")
                    key = ("xrange", (sheet, formel.area(plain))) if ":" in plain else ("xref", (sheet, formel.position(plain)))
                    kept.setdefault(key, text)
    except (formel.Unknown, ValueError):
        pass
    return kept


def _text(node, kept: dict) -> str | None:
    """Zurück zum Formeltext, wie Excel ihn speichert (englisch, Kommas). None, wenn es nicht sicher geht."""
    kind = node.kind
    if kind == "num":
        v = node.value
        return str(int(v)) if v == int(v) and abs(v) < 1e15 else repr(v)
    if kind == "str":
        return '"' + node.value.replace('"', '""') + '"'
    if kind == "bool":
        return "TRUE" if node.value else "FALSE"
    if kind == "empty":
        return ""
    if kind == "ref":
        return kept.get(("ref", node.value)) or address(node.value)
    if kind in ("xref", "xrange"):
        if (kind, node.value) in kept:
            return kept[(kind, node.value)]
        sheet, where = node.value
        name = sheet if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", sheet) else "'" + sheet.replace("'", "''") + "'"
        return f"{name}!{address(where) if kind == 'xref' else f'{address(where[:2])}:{address(where[2:])}'}"
    if kind == "range":
        if ("range", node.value) in kept:
            return kept[("range", node.value)]
        r1, c1, r2, c2 = node.value
        if r1 == 1 and r2 >= formel.MAX_ROW:
            return f"{column_letters(c1)}:{column_letters(c2)}"
        return f"{address((r1, c1))}:{address((r2, c2))}"
    if kind in ("neg", "pct"):
        inner = _text(node.args[0], kept)
        if inner is None:
            return None
        if node.args[0].kind == "op":
            inner = f"({inner})"
        return f"-{inner}" if kind == "neg" else f"{inner}%"
    if kind == "call":
        args = [_text(a, kept) for a in node.args]
        return None if any(a is None for a in args) else f"{node.value}({','.join(args)})"
    if kind == "op":
        power = formel._BINARY[node.value]
        parts = []
        for i, child in enumerate(node.args):
            part = _text(child, kept)
            if part is None:
                return None
            if child.kind == "op" and (formel._BINARY[child.value] < power or (i == 1 and formel._BINARY[child.value] == power)):
                part = f"({part})"
            parts.append(part)
        return f"{parts[0]}{node.value}{parts[1]}"
    return None


def _styled(like: str, pos: Pos) -> str:
    """Die Adresse pos mit denselben $ wie like ($F5 → $F4)."""
    match = re.fullmatch(r"(\$?)[A-Z]{1,3}(\$?)\d+", like.replace(" ", ""))
    if not match:
        return address(pos)
    return f"{match.group(1)}{column_letters(pos[1])}{match.group(2)}{pos[0]}"


def _paths(node, path=()):
    yield path, node
    for i, child in enumerate(node.args):
        if isinstance(child, formel.Node):
            yield from _paths(child, path + (i,))


def _put(node, path, new):
    if not path:
        return new
    args = list(node.args)
    args[path[0]] = _put(args[path[0]], path[1:], new)
    return formel.Node(node.kind, node.value, tuple(args))


def _pairs(call) -> list[tuple[int, object, object]]:
    """Die Bedingungen von SUMMEWENN(S) und Verwandten: (Stelle, Bereich, Bedingung)."""
    if call.value in _IFS_FROM:
        start = _IFS_FROM[call.value]
        return [(i, call.args[i], call.args[i + 1]) for i in range(start, len(call.args) - 1, 2)]
    if call.value in _IF_SINGLE and len(call.args) >= 2:
        return [(0, call.args[0], call.args[1])]
    return []


def _column(node) -> int | None:
    return node.value[1] if node.kind == "range" and node.value[1] == node.value[3] else None


def _value_column(columns: dict, value: str) -> int | None:
    """Die Textspalte, in der ein Wert steht („Nord“ → Region)."""
    hits = [col for col, info in columns.items() if info.kind == "text"
            and any(t.strip().casefold() == value.strip().casefold() for t in info.texts.values())]
    return hits[0] if len(hits) == 1 else None


def _edits(sheet: Sheet, node, want: _Want, columns: dict, target: int, rows: list[int], wanted: list[int],
           function: str):
    """Je eine kleine Änderung an der Formel, mit dem, was geändert ist."""
    inside = {q for p, call in _paths(node) if call.kind == "call" and _pairs(call)
              for q, _ in _paths(call, p) if q != p}
    for path, part in _paths(node):
        # Eine andere Funktion: MITTELWERT statt SUMME.
        if part.kind == "call" and want.function and part.value in _SWAPS:
            for other in _SWAPS[part.value]:
                if _FUNCTION_OF.get(other, other) == want.function or other == want.function:
                    yield _Edit("function", _put(node, path, formel.Node("call", other, part.args)),
                                {"old": part.value, "new": other})
        # Ein fester Teiler, der nicht zur Anzahl passt.
        if part.kind == "op" and part.value == "/" and part.args[1].kind == "num" and part.args[1].value != len(wanted):
            yield _Edit("divisor", _put(node, path + (1,), formel.Node("num", float(len(wanted)))),
                        {"old": part.args[1].value, "new": len(wanted)})
        if part.kind == "call" and (pairs := _pairs(part)):
            tested = {_column(area) for _, area, _ in pairs}
            for i, area, crit in pairs:
                col = _column(area)
                # Eine Bedingung, die der Name nicht sagt.
                if col is not None and col not in want.conditions and col != want.date_column:
                    if part.value in _IFS_FROM and len(pairs) >= 2:
                        args = part.args[:i] + part.args[i + 2:]
                        yield _Edit("remove", _put(node, path, formel.Node("call", part.value, args)),
                                    {"col": col, "crit": crit, "area": area})
                    elif not want.conditions and want.date_column is None and part.value in ("SUMIF", "SUMIFS", "AVERAGEIF", "AVERAGEIFS"):
                        total = part.args[0] if part.value.endswith("S") else (part.args[2] if len(part.args) > 2 else part.args[0])
                        plain = "SUM" if part.value.startswith("SUM") else "AVERAGE"
                        yield _Edit("remove", _put(node, path, formel.Node("call", plain, (total,))),
                                    {"col": col, "crit": crit, "area": area})
                # Ein anderer Wert als im Namen.
                if col in want.conditions and crit.kind == "str" and crit.value.strip().casefold() not in want.conditions[col]:
                    for value in want.conditions[col].values():
                        args = list(part.args)
                        # In einer Kreuztabelle aus der Kopfzelle, wie in den Nachbarzellen.
                        head = want.heads.get(value.casefold())
                        args[i + 1] = formel.Node("ref", head) if head else formel.Node("str", value)
                        yield _Edit("value", _put(node, path, formel.Node("call", part.value, tuple(args))),
                                    {"col": col, "old": crit.value, "new": value})
                # Die Bedingung aus der falschen Kopfzelle: $F5 („Ost“) statt $F4 („Nord“).
                said = sheet.inputs.get(crit.value) if crit.kind == "ref" else None
                if col in want.conditions and isinstance(said, str) and said.strip().casefold() not in want.conditions[col]:
                    for key, value in want.conditions[col].items():
                        if key in want.heads and want.heads[key] != crit.value:
                            args = list(part.args)
                            args[i + 1] = formel.Node("ref", want.heads[key])
                            yield _Edit("critref", _put(node, path, formel.Node("call", part.value, tuple(args))),
                                        {"col": col, "old": crit.value, "said": said.strip(), "new": want.heads[key],
                                         "value": value})
                # Die Bedingung in der falschen Spalte gesucht.
                text = crit.value if crit.kind == "str" else (said if isinstance(said, str) else None)
                if text and col is not None and (home := _value_column(columns, text)) not in (None, col):
                    r1, _, r2, _ = area.value
                    args = list(part.args)
                    args[i] = formel.Node("range", (r1, home, r2, home))
                    yield _Edit("critcol", _put(node, path, formel.Node("call", part.value, tuple(args))),
                                {"value": text.strip(), "old": col, "new": home})
            # Eine Bedingung aus dem Namen fehlt.
            for col, values in want.conditions.items():
                if col in tested or len(values) != 1:
                    continue
                value = next(iter(values.values()))
                like = pairs[0][1]
                if like.kind != "range":
                    continue
                r1, _, r2, _ = like.value
                # In einer Kreuztabelle aus der Kopfzelle, wie die anderen Bedingungen auch.
                head = want.heads.get(value.casefold())
                extra = (formel.Node("range", (r1, col, r2, col)),
                         formel.Node("ref", head) if head else formel.Node("str", value))
                if part.value in _IFS_FROM:
                    new = formel.Node("call", part.value, part.args + extra)
                elif part.value == "COUNTIF":
                    new = formel.Node("call", "COUNTIFS", part.args[:2] + extra)
                else:
                    total = part.args[2] if len(part.args) > 2 else part.args[0]
                    new = formel.Node("call", part.value + "S", (total,) + part.args[:2] + extra)
                yield _Edit("add", _put(node, path, new), {"col": col, "value": value})
            # Die Zahlen aus der falschen Spalte.
            if part.value not in ("COUNTIF", "COUNTIFS"):
                at = 0 if part.value in _IFS_FROM else (2 if len(part.args) > 2 else 0)
                total = part.args[at]
                if total.kind == "range" and _column(total) not in (None, target) and target in columns:
                    r1, _, r2, _ = total.value
                    args = list(part.args)
                    args[at] = formel.Node("range", (r1, target, r2, target))
                    yield _Edit("column", _put(node, path, formel.Node("call", part.value, tuple(args))),
                                {"old": _column(total), "new": target})
            # Datumsgrenzen: < statt <=, ein anderer Tag, Monat oder ein anderes Jahr.
            for i, area, crit in pairs:
                yield from _date_edits(node, path + (i + 1,), crit, columns, rows, want)
            continue
        # Bereiche einer Spalte: ein anderes Ende oder ein anderer Anfang, oder eine andere Spalte.
        if part.kind == "range" and _column(part) is not None and part.value[2] < formel.MAX_ROW and path not in inside:
            r1, col, r2, _ = part.value
            if col != target and col in columns and columns[col].kind == "zahl" and target in columns:
                yield _Edit("column", _put(node, path, formel.Node("range", (r1, target, r2, target))),
                            {"old": col, "new": target})
            near = sorted({r for r in rows if abs(r - r2) <= 12} | {rows[-1]})
            for end in near:
                if end != r2 and end >= r1:
                    yield _Edit("end", _put(node, path, formel.Node("range", (r1, col, end, col))),
                                {"area": part.value, "new": end}, abs(end - r2))
            near = sorted({r for r in rows if abs(r - r1) <= 12} | {rows[0]})
            for start in near:
                if start != r1 and start <= r2:
                    yield _Edit("start", _put(node, path, formel.Node("range", (start, col, r2, col))),
                                {"area": part.value, "new": start}, abs(start - r1))
        # SUMME(D2;D5;D5;D9): eine Zelle zu viel in der Liste, oder eine, die fehlt.
        if part.kind == "call" and part.value in _PLAIN and len(part.args) >= 2:
            for i, arg in enumerate(part.args):
                args = part.args[:i] + part.args[i + 1:]
                yield _Edit("drop", _put(node, path, formel.Node("call", part.value, args)), {"term": arg, "sign": ""})
            if all(a.kind in ("ref", "range") and a.value[1] == target for a in part.args):
                covered = {r for a in part.args for r in ([a.value[0]] if a.kind == "ref" else range(a.value[0], a.value[2] + 1))}
                for r in [r for r in wanted if r not in covered][:40]:
                    yield _Edit("missing", _put(node, path, formel.Node("call", part.value, part.args + (formel.Node("ref", (r, target)),))),
                                {"row": r})
        # Ein Teil zu viel oder mit falschem Vorzeichen: +D5, -D12.
        if part.kind == "op" and part.value in ("+", "-"):
            left, right = part.args
            yield _Edit("drop", _put(node, path, left), {"term": right, "sign": part.value})
            if part.value == "+":
                yield _Edit("drop", _put(node, path, right), {"term": left, "sign": "+"})
            flipped = "+" if part.value == "-" else "-"
            yield _Edit("sign", _put(node, path, formel.Node("op", flipped, part.args)), {"term": right, "old": part.value})


def _date_edits(node, path, crit, columns: dict, rows: list[int], want: _Want):
    """">="&DATUM(2025;3;1): die Grenze, die der Name verlangt (erster Tag des Monats, erster Tag danach), < und <=
    tauschen, ein Jahr aus der Tabelle, einen Monat oder Tag daneben."""
    if crit.kind != "op" or crit.value != "&" or crit.args[0].kind != "str" or crit.args[1].kind != "call" \
            or crit.args[1].value != "DATE" or len(crit.args[1].args) != 3 or any(a.kind != "num" for a in crit.args[1].args):
        return
    sign, call = crit.args
    years = {day.year for col, info in columns.items() if info.kind == "datum"
             for r in rows if (day := _day(info.dates.get(r)))}
    runs = _runs(list(want.months)) if want.months else [(1, 12)]
    year = want.year or (next(iter(years)) if len(years) == 1 else None)
    if want.date_column is not None and year and len(runs) == 1:
        (a, b), lower = runs[0], sign.value in (">", ">=")
        y, m = (year, a) if lower else (year + (b == 12), b % 12 + 1)
        bound = formel.Node("op", "&", (formel.Node("str", ">=" if lower else "<"),
                                        formel.Node("call", "DATE", tuple(formel.Node("num", float(v)) for v in (y, m, 1)))))
        if bound != crit:
            yield _Edit("bound", _put(node, path, bound), {"old": crit, "new": bound})
    swap = {"<": "<=", "<=": "<", ">": ">=", ">=": ">"}
    if sign.value in swap:
        yield _Edit("operator", _put(node, path + (0,), formel.Node("str", swap[sign.value])),
                    {"old": crit, "new": formel.Node("op", "&", (formel.Node("str", swap[sign.value]), call))})
    y, m, d = (int(a.value) for a in call.args)
    options = {(year, m, d) for year in years} | {(y, m + 1, d), (y, m - 1, d), (y, m, d + 1), (y, m, d - 1), (y, m, 1)}
    for year, month, day in sorted(options):
        if (year, month, day) == (y, m, d) or not 1 <= month <= 12 or not 1 <= day <= 31:
            continue
        new = formel.Node("call", "DATE", tuple(formel.Node("num", float(v)) for v in (year, month, day)))
        yield _Edit("date", _put(node, path + (1,), new), {"old": crit, "new": formel.Node("op", "&", (sign, new))},
                    abs(year - y) * 400 + abs(month - m) * 31 + abs(day - d))


def _hits(got, expected: float, tolerance: float) -> bool:
    return isinstance(got, float) and not isinstance(got, bool) and \
        abs(got - expected) <= tolerance + 1e-9 * max(1.0, abs(expected))


def _diff(old, new, path=()):
    """Die eine Stelle, an der sich zwei Formeln unterscheiden: (Pfad, neuer Teil)."""
    if old == new:
        return None
    if old.kind == new.kind and old.value == new.value and len(old.args) == len(new.args):
        changed = [i for i, (a, b) in enumerate(zip(old.args, new.args)) if a != b]
        if len(changed) == 1 and isinstance(old.args[changed[0]], formel.Node):
            return _diff(old.args[changed[0]], new.args[changed[0]], path + (changed[0],))
    return path, new


def _at(node, path):
    for i in path:
        if not isinstance(node, formel.Node) or i >= len(node.args):
            return None
        node = node.args[i]
    return node


def _relocate(old, new, rest):
    """Wo ein Teil von old nach der Änderung zu new steht: gleich, eine Stufe höher (+D5 gestrichen) oder an
    einer anderen Stelle der Liste (SUMMEWENN wird SUMMEWENNS). None, wenn es ihn so nicht mehr gibt."""
    if rest and (here := _at(new, rest)) is not None and here == _at(old, rest):
        return rest
    if not rest or not isinstance(old, formel.Node) or rest[0] >= len(old.args):
        return None
    child = old.args[rest[0]]
    if new == child:
        return rest[1:]
    places = [j for j, c in enumerate(new.args) if c == child]
    return ((min(places, key=lambda j: abs(j - rest[0])),) + rest[1:]) if places else None


def _both(node, first: _Edit, second: _Edit):
    """Beide Änderungen in einer Formel. Am selben Bereich ein anderer Anfang und ein anderes Ende."""
    if {first.kind, second.kind} == {"start", "end"} and first.info["area"] == second.info["area"]:
        r1, col, r2, _ = first.info["area"]
        top = next(e.info["new"] for e in (first, second) if e.kind == "start")
        bottom = next(e.info["new"] for e in (first, second) if e.kind == "end")
        where = next(path for path, part in _paths(node) if part.kind == "range" and part.value == first.info["area"])
        return _put(node, where, formel.Node("range", (top, col, bottom, col))) if top <= bottom else None
    for a, b in ((first, second), (second, first)):
        (pa, sa), (pb, sb) = _diff(node, a.node), _diff(node, b.node)
        if pa == pb:
            # Dieselbe Funktion: die eine ändert nur den Namen (MITTELWERTWENNS statt SUMMEWENNS), die andere die
            # Bedingungen.
            here = _at(node, pa)
            if a.kind == "function" and sa.kind == sb.kind == "call" and sa.args == here.args and _FUNCTION_OF.get(sa.value) \
                    and sa.value in _IFS_FROM and sb.value == here.value:
                return _put(node, pa, formel.Node("call", sa.value, sb.args))
            continue
        if pb[:len(pa)] != pa:
            if pa[:len(pb)] != pb:
                return _put(a.node, pb, sb)
            continue
        rest = _relocate(_at(node, pa), sa, pb[len(pa):])
        if rest is not None:
            return _put(a.node, pa + rest, sb)
    return None


def _exact(use: _Use | None, wanted: list[int]) -> bool:
    """Nimmt die korrigierte Formel genau die gemeinten Zeilen, jede einmal?"""
    if use is None or any(k not in (0, 1) for k in use.rows.values()):
        return False
    if {r for r, k in use.rows.items() if k} != set(wanted):
        return False
    return use.kind != "ratio" or use.divisor == len(wanted) or use.counted == sorted(wanted)


def _clean(sheet: Sheet, node, target: int, data: set[int]) -> bool:
    """Hebt sich in der Formel nichts auf? =SUMME(D2:D12)-D12 rechnet zwar richtig, aber die gemeinte Korrektur ist
    =SUMME(D2:D11). Mit Bedingungen (SUMMEWENNS) nimmt die Formel ohnehin nicht alle Zeilen ihrer Bereiche."""
    if any(p.kind == "call" and p.value in _CONDITIONAL for _, p in _paths(node)):
        return True
    use = _use(sheet, node, target, data)
    if use is None:
        return True
    singles, areas, _ = formel.references(node, sheet.name)
    named = {r for r, c in singles if c == target and _number_at(sheet, (r, target))}
    for r1, c1, r2, c2 in areas:
        if c1 <= target <= c2:
            named |= set(_range_rows(sheet, (r1, target, r2, target), target) or [])
    return named <= {r for r, k in use.rows.items() if k}


# So viele Paare von Änderungen rechnet er höchstens nach, die vielversprechendsten zuerst.
MAX_PAIRS = 150


def _repair(sheet: Sheet, calc, pos: Pos, node, want: _Want, columns: dict, target: int, rows: list[int],
            wanted: list[int], function: str, expected: float, tolerance: float) -> _Edit | None:
    """Die kleinste Änderung an der Formel, nach der sie genau richtig rechnet. Gibt es keine einzelne, zwei
    Änderungen an verschiedenen Stellen (Art „pair“), die zusammen genau die gemeinten Zeilen nehmen."""
    hits, tried = [], []
    seen = set()

    def safe(new) -> bool:
        singles, areas, _ = formel.references(new, sheet.name)
        return pos not in singles and not any(r1 <= pos[0] <= r2 and c1 <= pos[1] <= c2 for r1, c1, r2, c2 in areas)

    def run(new):
        try:
            return calc.run(new)
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, IndexError):
            return None

    for edit in _edits(sheet, node, want, columns, target, rows, wanted, function):
        if edit.node in seen:
            continue
        seen.add(edit.node)
        if not safe(edit.node):
            continue
        got = run(edit.node)
        if _hits(got, expected, tolerance):
            hits.append(edit)
        elif isinstance(got, float) and not isinstance(got, bool) and math.isfinite(got):
            tried.append((edit, got))
    data = set(rows)
    best = min(hits, key=lambda e: (_ORDER.index(e.kind), e.size)) if hits else None
    if best is not None and _clean(sheet, best.node, target, data):
        return best
    # Zwei Fehler: Summen ändern sich je Stelle für sich, so sagt der Wert jeder einzelnen Änderung voraus, welche
    # zwei zusammen passen. Nachgerechnet wird trotzdem jede, und sie muss genau die gemeinten Zeilen nehmen.
    base = run(node)
    if not isinstance(base, float) or isinstance(base, bool) or not math.isfinite(base):
        return best
    scale = max(1.0, abs(expected))
    pairs = sorted(((abs(g1 + g2 - base - expected) / scale, i, j) for i, (_, g1) in enumerate(tried)
                    for j, (_, g2) in enumerate(tried) if i < j), key=lambda x: x[0])
    found, done = [], 0
    for _, i, j in pairs:
        if done >= MAX_PAIRS:
            break
        first, second = tried[i][0], tried[j][0]
        if first.kind == second.kind == "function":
            continue
        both = _both(node, first, second)
        if both is None or both in seen or not safe(both):
            continue
        seen.add(both)
        done += 1
        if _hits(run(both), expected, tolerance) and _exact(_use(sheet, both, target, data), wanted):
            ordered = sorted((first, second), key=lambda e: (_ORDER.index(e.kind), e.size))
            found.append(_Edit("pair", both, {"edits": ordered}, sum(e.size for e in ordered)))
    found = [e for e in found if _clean(sheet, e.node, target, data)]
    if not found:
        return best
    # Die kleinere Änderung zuerst: eine Zeile länger und ein Vorzeichen, nicht zwei Zeilen länger und -D11 weg.
    return min(found, key=lambda e: (e.size, sorted(_ORDER.index(x.kind) for x in e.info["edits"])))


def _row_word(n: int) -> str:
    return T({1: "eine Zeile", 2: "zwei Zeilen", 3: "drei Zeilen"}.get(n, f"{n} Zeilen"),
             {1: "one row", 2: "two rows", 3: "three rows"}.get(n, f"{n} rows"))


def _with(sheet: Sheet, columns: dict, target: int, rows: list[int], col: int | None) -> str:
    """„D12 (Laptop), D13 (Tastatur) und 3 weitere“: die Zeilen mit dem Wert in col."""
    shown = []
    for r in rows[:MAX_LISTED]:
        value = columns[col].texts.get(r, "").strip() if col in columns else ""
        if not value and col in columns and columns[col].kind == "datum":
            value = sheet.shown.get((r, col)) if isinstance(sheet.shown.get((r, col)), str) else ""
        shown.append(f"{address((r, target))} ({value})" if value else address((r, target)))
    if len(rows) > MAX_LISTED:
        return f"{', '.join(shown)}{_more(len(rows) - MAX_LISTED)}"
    return _listing(shown)


_VERBS = ("rechnet", "teilt", "filtert", "prüft", "sucht", "calculates", "divides", "filters", "also filters",
          "does not filter", "checks", "looks for")


def _two(sheet: Sheet, repair: _Edit, kept: dict, columns: dict, want: _Want, target: int, wanted: list[int],
         rows: list[int]) -> str:
    """„hat zwei Fehler. Erstens: … Zweitens: …“. Jeder Satz sagt, was dieser Fehler allein bewirkt: gerechnet
    mit der Formel, in der nur er noch steckt."""
    data = set(rows)
    kinds = {e.kind for e in repair.info["edits"]}
    if kinds == {"start", "end"}:
        # Derselbe Bereich an beiden Enden falsch: ein Satz.
        start, end = sorted(repair.info["edits"], key=lambda e: e.kind != "start")
        r1, col, r2, _ = start.info["area"]
        top, bottom = start.info["new"], end.info["new"]
        area = _text(formel.Node("range", start.info["area"]), kept)
        lost = [r for r in wanted if top <= r < r1 or r2 < r <= bottom]
        gone = [r for r in list(range(r1, top)) + list(range(bottom + 1, r2 + 1)) if r in columns[target].numbers]
        ends = T([f"fängt {_row_word(abs(r1 - top))} zu {'spät' if top < r1 else 'früh'} an",
                  f"hört {_row_word(abs(bottom - r2))} zu {'früh auf' if bottom > r2 else 'spät auf'}"],
                 [f"starts {_row_word(abs(r1 - top))} too {'late' if top < r1 else 'early'}",
                  f"ends {_row_word(abs(bottom - r2))} too {'early' if bottom > r2 else 'late'}"])
        text = T(f"nimmt {area}, aber der Bereich {' und '.join(ends)}.", f"takes {area}, but the range {' and '.join(ends)}.")
        if lost:
            listed, n = _row_list(sheet, columns, want, target, lost, False), len(lost)
            text += T(f" Dadurch {'fehlt' if n == 1 else 'fehlen'} {listed}.", f" As a result, {listed} {_n(n, 'is', 'are')} missing.")
        if gone:
            listed, n = _row_list(sheet, columns, want, target, gone, True), len(gone)
            text += T(f" {listed} {'gehört' if n == 1 else 'gehören'} nicht dazu.", f" {listed} {_n(n, 'does', 'do')} not belong.")
        return text
    sentences = []
    for edit in repair.info["edits"]:
        other = next(e for e in repair.info["edits"] if e is not edit)
        use = _use(sheet, other.node, target, data)
        text = _cause(sheet, edit, kept, columns, want, target, wanted, use, rows, "", "", alone=False)
        text = T(f"Die Formel {text}", f"The formula {text}") if text.startswith(_VERBS) else text[:1].upper() + text[1:]
        sentences.append(text)
    return T(f"hat zwei Fehler. Erstens: {sentences[0]} Zweitens: {sentences[1]}",
             f"has two errors. First: {sentences[0]} Second: {sentences[1]}")


def _cause(sheet: Sheet, edit: _Edit, kept: dict, columns: dict, want: _Want, target: int, wanted: list[int],
           use, rows: list[int], how: str, written: str, alone: bool = True) -> str:
    """Was an der Formel falsch ist, als Satz nach „„Umsatz Ost“ in G4 …“. alone=False: ohne „summiert …, aber“
    vorn, für einen von zwei Fehlern."""
    plan = sheet.plan
    head = lambda col: f"{plan.headers.get(col, column_letters(col))}"  # noqa: E731
    shown = lambda n: _german_formula(_text(n, kept) or "")  # noqa: E731
    who = want.describe(sheet) or T("allen Zeilen", "all rows")
    taken = {r for r, k in use.rows.items() if k} if use else set()
    missing = [r for r in wanted if r not in taken] if use else []
    extra = sorted(r for r in taken if r not in set(wanted)) if use else []
    info = edit.info
    if edit.kind == "function":
        asked = _FUNCTION_OF.get(info['new'], info['new'])
        return T(f"rechnet {_german(info['old'])}, der Name verlangt aber {_ASKED[asked]} ({_german(info['new'])}).",
                 f"calculates {info['old']}, but the name asks for {_SAID_EN[asked]} ({info['new']}).")
    if edit.kind == "divisor":
        many = f"Mit {who} gibt es aber" if want.describe(sheet) else "Es sind aber"
        return T(f"teilt durch {_num(info['old'])}. {many} {info['new']} Zeilen.",
                 f"divides by {_num(info['old'])}. But there are {info['new']} rows with {who}." if want.describe(sheet)
                 else f"divides by {_num(info['old'])}. But there are {info['new']} rows.")
    if edit.kind == "remove":
        col, crit = info["col"], info["crit"]
        what = _value(head(col), crit.value) if crit.kind == "str" else f"{head(col)} {shown(crit)}"
        text = T(f"filtert zusätzlich nach {what} ({shown(info['area'])};{shown(crit)}). Das steht nicht im Namen, "
                 f"der nur {who} sagt.",
                 f"also filters by {what} ({shown(info['area'])};{shown(crit)}). That is not in the name, "
                 f"which only says {who}.")
        if missing:
            listed, n = _with(sheet, columns, target, missing, col), len(missing)
            text += T(f" Dadurch {'fehlt' if n == 1 else 'fehlen'} {n} der {len(wanted)} Zeilen: {listed}.",
                      f" As a result, {n} of the {len(wanted)} rows {_n(n, 'is', 'are')} missing: {listed}.")
        return text
    if edit.kind == "value":
        return T(f"prüft {head(info['col'])} „{info['old']}“, der Name sagt aber „{info['new']}“.",
                 f"checks {head(info['col'])} “{info['old']}”, but the name says “{info['new']}”.")
    if edit.kind == "critref":
        old = kept.get(('ref', info['old'])) or address(info['old'])
        new = kept.get(('ref', info['new'])) or address(info['new'])
        return T(f"nimmt {head(info['col'])} aus {old} („{info['said']}“), gemeint ist aber „{info['value']}“ aus {new}.",
                 f"takes {head(info['col'])} from {old} (“{info['said']}”), but “{info['value']}” from {new} is meant.")
    if edit.kind == "critcol":
        old, new = column_letters(info['old']), column_letters(info['new'])
        return T(f"sucht „{info['value']}“ in Spalte {old} ({head(info['old'])}), dort steht es "
                 f"aber nicht. „{info['value']}“ steht in Spalte {new} ({head(info['new'])}).",
                 f"looks for “{info['value']}” in column {old} ({head(info['old'])}), but it is not there. "
                 f"“{info['value']}” is in column {new} ({head(info['new'])}).")
    if edit.kind == "add":
        text = T(f"filtert nicht nach {head(info['col'])} „{info['value']}“, obwohl das im Namen steht.",
                 f"does not filter by {head(info['col'])} “{info['value']}”, although the name says so.")
        if extra:
            listed, n = _with(sheet, columns, target, extra, info['col']), len(extra)
            text += T(f" Dadurch {'zählt' if n == 1 else 'zählen'} {n} Zeilen mit anderem {head(info['col'])} mit: {listed}.",
                      f" As a result, {n} {_n(n, 'row', 'rows')} with a different {head(info['col'])} "
                      f"{_n(n, 'is', 'are')} counted too: {listed}.")
        return text
    if edit.kind == "column":
        old, new = column_letters(info['old']), column_letters(info['new'])
        return T(f"rechnet mit Spalte {old} ({head(info['old'])}) statt mit Spalte {new} ({head(info['new'])}).",
                 f"calculates with column {old} ({head(info['old'])}) instead of column {new} ({head(info['new'])}).")
    lead = T(f"{how or written}, aber ", f"{how or written}, but ") if alone else ""
    if edit.kind in ("end", "start"):
        r1, col, r2, _ = info["area"]
        area = _text(formel.Node("range", info["area"]), kept)
        new = info["new"]
        if edit.kind == "end":
            lost = [r for r in wanted if r2 < r <= new] or [r for r in range(r2 + 1, new + 1)]
            gone = [r for r in range(new + 1, r2 + 1) if r in columns[target].numbers]
            if new > r2:
                listed, n = _row_list(sheet, columns, want, target, lost, False), len(lost)
                return T(f"{lead}der Bereich {area} hört {_row_word(new - r2)} zu früh auf: "
                         f"{listed} {'gehört' if n == 1 else 'gehören'} noch dazu.",
                         f"{lead}the range {area} ends {_row_word(new - r2)} too early: "
                         f"{listed} also {_n(n, 'belongs', 'belong')}.")
            listed, n = _row_list(sheet, columns, want, target, gone, True), len(gone)
            return T(f"{lead}der Bereich {area} reicht {_row_word(r2 - new)} zu weit: "
                     f"{listed} {'gehört' if n == 1 else 'gehören'} nicht mehr dazu.",
                     f"{lead}the range {area} goes {_row_word(r2 - new)} too far: "
                     f"{listed} no longer {_n(n, 'belongs', 'belong')}.")
        lost = [r for r in wanted if new <= r < r1]
        gone = [r for r in range(r1, new) if r in columns[target].numbers]
        if new < r1:
            listed, n = _row_list(sheet, columns, want, target, lost, False), len(lost)
            return T(f"{lead}der Bereich {area} fängt {_row_word(r1 - new)} zu spät an: "
                     f"{listed} {'fehlt' if n == 1 else 'fehlen'}.",
                     f"{lead}the range {area} starts {_row_word(r1 - new)} too late: {listed} {_n(n, 'is', 'are')} missing.")
        listed, n = _row_list(sheet, columns, want, target, gone, True), len(gone)
        return T(f"{lead}der Bereich {area} fängt {_row_word(new - r1)} zu früh an: "
                 f"{listed} {'gehört' if n == 1 else 'gehören'} nicht dazu.",
                 f"{lead}the range {area} starts {_row_word(new - r1)} too early: {listed} {_n(n, 'does', 'do')} not belong.")
    # Eine lange Formel steht nicht ganz da, dann sagt es, was sie nimmt.
    lead = T(f"{written if len(written) <= 90 or not how else how}, aber ",
             f"{written if len(written) <= 90 or not how else how}, but ") if alone else ""
    if edit.kind == "missing":
        row = info["row"]
        listed = _row_list(sheet, columns, want, target, [row], False)
        return T(f"{lead}{listed} fehlt in der Formel.", f"{lead}{listed} is missing from the formula.")
    if edit.kind == "drop":
        term = info["term"]
        text = _german_formula(_text(term, kept) or "")
        if term.kind == "ref" and term.value[1] == target and use and use.rows.get(term.value[0], 0) >= 2:
            inside = next((_text(p, kept) for _, p in _paths(edit.node)
                           if p.kind == "range" and p.value[0] <= term.value[0] <= p.value[2] and p.value[1] == target), "")
            if not inside:
                return T(f"{lead}{address(term.value)} steht zweimal in der Formel und zählt doppelt.",
                         f"{lead}{address(term.value)} appears twice in the formula and counts twice.")
            return T(f"{lead}{address(term.value)} steckt schon in {inside} und wird mit {info['sign']}{text} ein zweites Mal gezählt.",
                     f"{lead}{address(term.value)} is already in {inside} and is counted a second time with {info['sign']}{text}.")
        if term.kind == "ref" and term.value[1] == target and info["sign"] == "-" and term.value[0] in wanted:
            inside = next((_text(p, kept) for _, p in _paths(edit.node)
                           if p.kind == "range" and p.value[0] <= term.value[0] <= p.value[2] and p.value[1] == target), "")
            if inside:
                return T(f"{lead}{address(term.value)} steckt schon in {inside} und wird mit -{text} wieder abgezogen.",
                         f"{lead}{address(term.value)} is already in {inside} and is subtracted again with -{text}.")
            return T(f"{lead}-{text} zieht {address(term.value)} ab, das gehört aber dazu.",
                     f"{lead}-{text} subtracts {address(term.value)}, but it should be included.")
        if term.kind == "ref" and term.value[1] == target:
            name = _ident(sheet, columns, want, term.value[0], True)
            where = T(f"zu {who}", f" to {who}") if want.describe(sheet) else T("dazu", "")
            cell = address(term.value) + (f' ({name})' if name else '')
            if not info["sign"]:
                return T(f"{lead}{cell} gehört nicht {where}.", f"{lead}{cell} does not belong{where}.")
            verb = T("zieht {} ab", "subtracts {}") if info["sign"] == "-" else T("nimmt {} mit", "includes {}")
            return T(f"{lead}{info['sign']}{text} {verb.format(cell)}, und das gehört nicht {where}.",
                     f"{lead}{info['sign']}{text} {verb.format(cell)}, and that does not belong{where}.")
        return T(f"{lead}der Teil {info['sign']}{text} gehört nicht in die Formel.",
                 f"{lead}the part {info['sign']}{text} does not belong in the formula.")
    if edit.kind == "sign":
        text = _german_formula(_text(info["term"], kept) or "")
        if info["old"] == "-":
            return T(f"{lead}{text} wird abgezogen statt dazugezählt.", f"{lead}{text} is subtracted instead of added.")
        return T(f"{lead}{text} wird dazugezählt, gehört aber abgezogen.", f"{lead}{text} is added but should be subtracted.")
    if edit.kind in ("bound", "operator", "date"):
        old, new = shown(info["old"]), shown(info["new"])
        effect = ""
        if missing:
            listed, n = _row_list(sheet, columns, want, target, missing, False), len(missing)
            effect = T(f" Dadurch {'fehlt' if n == 1 else 'fehlen'} {listed}.", f" As a result, {listed} {_n(n, 'is', 'are')} missing.")
        elif extra:
            listed, n = _row_list(sheet, columns, want, target, extra, True), len(extra)
            effect = T(f" Dadurch {'zählt' if n == 1 else 'zählen'} {listed} mit.", f" As a result, {listed} {_n(n, 'is', 'are')} counted too.")
        return T(f"{lead}die Grenze {old} müsste {new} heißen.{effect}", f"{lead}the limit {old} should be {new}.{effect}")
    return ""


# --- Kennzahlen mit Namen ---


_SAID = {"SUM": "die Summe", "AVERAGE": "der Mittelwert", "COUNT": "die Anzahl", "COUNTA": "die Anzahl",
         "MEDIAN": "der Median", "MAX": "das Maximum", "MIN": "das Minimum"}
_SAID_EN = {"SUM": "the sum", "AVERAGE": "the average", "COUNT": "the count", "COUNTA": "the count",
            "MEDIAN": "the median", "MAX": "the maximum", "MIN": "the minimum"}
_ASKED = {"SUM": "die Summe", "AVERAGE": "den Mittelwert", "COUNT": "die Anzahl", "COUNTA": "die Anzahl",
          "MEDIAN": "den Median", "MAX": "das Maximum", "MIN": "das Minimum"}


def _filled_at(sheet: Sheet, pos: Pos) -> bool:
    cell = sheet.cells.get(pos)
    return cell is not None and (cell.formula is not None or cell.value != "")


def _reach(sheet: Sheet, pos: Pos, rows: set[int], columns: dict) -> set[int]:
    """Aus welchen Zahlenspalten der Tabelle eine Formel am Ende rechnet, auch über andere Zellen (H3 = H2/9)."""
    seen, todo, found = {pos}, [pos], set()
    while todo and len(seen) < 5000:
        for p in sheet.precedents.get(todo.pop(), ()):
            if p in seen:
                continue
            seen.add(p)
            if p[0] in rows and p[1] in columns:
                if columns[p[1]].kind == "zahl":
                    found.add(p[1])
            elif p in sheet.nodes:
                todo.append(p)
    return found


def _rounding(node) -> float:
    """Wie weit RUNDEN in der Formel das Ergebnis höchstens verschiebt."""
    step = 0.0
    for call in formel.calls(node):
        if call.value in _ROUNDING:
            digits = call.args[1].value if len(call.args) > 1 and call.args[1].kind == "num" else 0
            step = max(step, 10 ** -digits)
    return step


def _labelled(sheet: Sheet, rows: list[int], columns: dict) -> list[Finding]:
    """Kennzahlen, deren Name sagt, welche Zeilen dazugehören, und deren Formel oder feste Zahl nicht passt."""
    plan = sheet.plan
    data = set(rows)
    found: dict[Pos, Finding] = {}
    used = {n for node in sheet.nodes.values() if node is not None for n in formel.literals(node)}
    calc, _ = sheet.evaluator()
    from .sensitivitaet import tables as sensitivities

    # Köpfe und Felder einer Sensitivität sind Stufen einer Annahme, keine Kennzahlen über Zeilen.
    varied = {p for t in sensitivities(sheet.cells) for p in t.heads() | t.body()}
    for pos, cell in sorted(sheet.cells.items()):
        r, c = pos
        if pos in varied:
            continue
        value = sheet.inputs.get(pos)
        if not isinstance(value, float) or isinstance(value, bool) or pos in sheet.unknown or plan.is_derived(r, c):
            continue
        if c in plan.headers and r in data:
            continue
        cross = _cross(sheet, r, c) if c not in plan.headers else None
        if cross:
            named = _cross_wanted(sheet, cross, columns) if not any("%" in text for text, _ in cross) else None
            if named is None:
                continue
            label, want = named
        else:
            label = _left_label(sheet.cells, r, c) or (plan.names.get(r, "") if c in plan.headers else "")
            # „+ Investitionen“, „– Abschreibungen“: eine Zeile einer Überleitung, die den Posten übernimmt, keine
            # Kennzahl über alle Zeilen dieses Namens.
            if label.strip()[:1] and label.strip()[0] in "+-–−±":
                continue
            # „Umsatz Nord in %“ ist ein Anteil, keine Summe.
            want = _wanted(label, sheet, columns) if label and "%" not in label else None
            want = _not_itself(sheet, want, r, label)
            # „Summe“ unter „Mittlere Liegedauer“: Dort steht der Mittelwert über alle, nicht die Summe der Mittelwerte.
            if want is not None and want.function == "SUM" and c in plan.headers \
                    and (set(_words(plan.headers[c])) & _NOT_SUMMED or woerter.rate_word(_words(plan.headers[c]))
                         or row_rates(sheet, c, rows)):
                continue
        if want is None:
            continue
        node, use = sheet.nodes.get(pos), None
        if node is None:
            # Eine feste Zahl, mit der Formeln rechnen (Total Sitze 577), ist eine Vorgabe, keine Kennzahl.
            if sheet.dependents.get(pos) or value in used:
                continue
            # Sie braucht den Spaltenkopf oder ein Wort wie Summe im Namen: „Nord 0,15“ ist eher ein Satz.
            numeric = [col for col, info in columns.items() if info.kind == "zahl"]
            targets = sorted(want.targets) or (numeric if want.function and len(numeric) == 1 else [])
            if cross and not targets:
                # Womit die Nachbarn in der Kreuztabelle rechnen, gilt auch für die feste Zahl dazwischen.
                near = [(r, cc) for cc in range(c - 3, c + 4) if cc != c] + [(rr, c) for rr in range(r - 3, r + 4) if rr != r]
                reached = [_reach(sheet, p, data, columns) for p in near if p in sheet.nodes and p not in sheet.unknown]
                common = set.intersection(*reached) if reached else set()
                targets = sorted(common) if len(common) == 1 else []
            if len(targets) != 1:
                continue
            target = targets[0]
            function = want.function or "SUM"
        else:
            reach = _reach(sheet, pos, data, columns)
            targets = (want.targets & reach) if want.targets else reach
            if not targets and not cross and len(want.targets) == 1 and len(reach) == 1 and \
                    _measured(sheet.plan.headers.get(next(iter(want.targets)), ""), label):
                other = _other_column(sheet, calc, pos, node, cell.formula or "", label, want, columns, data,
                                      next(iter(want.targets)), next(iter(reach)))
                if other is not None:
                    found[pos] = other
                continue
            if len(targets) != 1:
                continue
            target = next(iter(targets))
            # Was die Formel mit den Zahlen von jetzt ergibt, nicht was Excel zuletzt gespeichert hat.
            try:
                got = calc.run(node)
            except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
                got = value
            if not isinstance(got, float) or isinstance(got, bool):
                continue
            value = got
            use = _use(sheet, node, target, data)
            # In einer Kreuztabelle nur, was sicher aus den Zeilen rechnet: Anteile (G4/G$8) sagen die Köpfe nicht.
            if cross and use is None:
                continue
            # Sagt der Name nicht, wie gerechnet wird, gilt, was die Formel rechnet.
            function = want.function or (_FUNCTION[use.kind] if use else "SUM")
        numbers = columns[target].numbers
        wanted = sorted(row for row in numbers if want.matches(columns, row))
        # Was mit der Summe selbst rechnet (MwSt = Summe * 0,19, Gesamt darunter), gehört nicht hinein: Das wäre ein
        # Kreisbezug.
        derived = _after(sheet, pos)
        wanted = [row for row in wanted if (row, target) not in derived]
        if c == target and any(row < r for row in wanted):
            # Eine Summe in ihrer eigenen Spalte fasst den Block darüber. Was nach einer Leerzeile darunter steht
            # (Vorjahr, Ziel), ist ein eigener Block, auch wenn die Tabelle daneben weiterläuft.
            below = [row for row in wanted if row > r]
            if below and any(not _filled_at(sheet, (row, c)) for row in range(r + 1, below[0])):
                wanted = [row for row in wanted if row < r]
            # Eine Summe steht über oder unter ihren Teilen, nie mitten drin: Liest die Formel nur Zeilen darüber,
            # sind die Zeilen darunter (Umsatz laut Bank, Abweichung) etwas anderes.
            elif below and node is not None:
                read = {p[0] for p in _before(sheet, pos) if p[1] == c}
                for area in sheet.ranges.get(pos, ()):
                    if area[1] <= c <= area[3]:
                        read |= set(range(area[0], area[2] + 1))
                if read and max(read) < r:
                    wanted = [row for row in wanted if row < r]
        # In einer Kreuztabelle gibt es leere Felder (West im Februar): Summe und Anzahl sind dann 0.
        if not wanted and not (cross and function in ("SUM", "COUNT")):
            continue
        expected = _aggregate(function, [numbers[row] for row in wanted])
        step = _rounding(node) if node is not None else 10 ** -_decimals(value)
        tolerance = 0.5 * step
        if abs(value - expected) <= tolerance + 1e-9 * max(1.0, abs(expected)):
            continue
        # Lässt die Formel genau die Codes für „fehlt“ weg (999, -99), rechnet sie richtig; die Codes meldet statistik.
        from .auswertung import codes_in

        codes = codes_in([numbers[row] for row in wanted]) if node is not None else set()
        real = [numbers[row] for row in wanted if numbers[row] not in codes]
        if codes and real and abs(value - (meant := _aggregate(function, real))) <= tolerance + 1e-9 * max(1.0, abs(meant)):
            continue
        header = plan.headers[target]
        who = want.describe(sheet)
        n = len(wanted)
        belongs = (T(f"Mit {who} gibt es {n} {'Zeile' if n == 1 else 'Zeilen'}",
                     f"There {_n(n, 'is', 'are')} {n} {_n(n, 'row', 'rows')} with {who}") if who
                   else T(f"In {header} stehen {n} Zahlen", f"{header} has {n} {_n(n, 'number', 'numbers')}"))
        formula = _suggest(sheet, want, function, target, pos, columns, rows, expected)
        # Der Tipp in beiden Sprachen: Aus dem deutschen liest einsetzen die Korrektur (_tipped).
        if node is None:
            tip_de, tip_en = ((f"Mit ={formula} rechnet die Zelle selbst. Nachgerechnet ergibt das {_num(expected)}, und es "
                               "bleibt richtig, wenn sich Zahlen ändern.",
                               f"With ={formula} the cell calculates it itself. Recalculated, that gives {_num(expected)}, "
                               "and it stays correct when numbers change.") if formula
                              else ("Rechne die Zahl mit einer Formel aus der Tabelle, dann bleibt sie richtig, wenn sich "
                                    "Zahlen ändern.",
                                    "Calculate the number with a formula from the table, then it stays correct when numbers "
                                    "change."))
        else:
            tip_de, tip_en = ((f"Nimm ={formula}. Nachgerechnet ergibt das {_num(expected)}, und es rechnet immer mit genau "
                               "diesen Zeilen, auch wenn welche dazukommen oder neu sortiert wird.",
                               f"Use ={formula}. Recalculated, that gives {_num(expected)}, and it always uses exactly these "
                               "rows, even when rows are added or the table is re-sorted.") if formula
                              else (f"Nimm genau die Zeilen mit {who or 'allen Zahlen'} in die Formel.",
                                    f"Include exactly the rows with {who} in the formula." if who
                                    else "Include all the numbers in the formula."))
        how, parts = _explain(sheet, columns, want, function, target, wanted, use, rows) if use else ("", [])
        why = _identify(sheet, columns, want, function, target, wanted, value, expected, rows, tolerance) if not parts else ""
        # Die Stelle in der Formel, die falsch ist, und die Formel ohne den Fehler, nachgerechnet.
        repair, fixed = None, ""
        if node is not None:
            kept = _kept(cell.formula or "")
            repair = _repair(sheet, calc, pos, node, want, columns, target, rows, wanted, function, expected, tolerance)
            for edit in (repair.info.get("edits", [repair]) if repair else []):
                if edit.kind == "critref":
                    kept.setdefault(("ref", edit.info["new"]), _styled(kept.get(("ref", edit.info["old"]), ""), edit.info["new"]))
                if edit.kind == "value":
                    head = want.heads.get(edit.info["new"].casefold())
                    if head:
                        letters = column_letters(head[1])
                        kept.setdefault(("ref", head), f"${letters}{head[0]}" if head[0] == pos[0] else
                                        f"{letters}${head[0]}" if head[1] == pos[1] else f"${letters}${head[0]}")
                if edit.kind == "add":
                    # Neue Teile so geschrieben wie die alten: $C:$C neben $B:$B, H$3 für den Spaltenkopf.
                    fixed_ranges = [text for (kind, _), text in kept.items() if kind == "range"]
                    for _, part in _paths(edit.node):
                        if part.kind == "range" and ("range", part.value) not in kept and fixed_ranges \
                                and all(text.startswith("$") for text in fixed_ranges):
                            kept[("range", part.value)] = "$" + _text(part, {}).replace(":", ":$")
                    head = want.heads.get(edit.info["value"].casefold())
                    if head:
                        letters = column_letters(head[1])
                        kept.setdefault(("ref", head), f"${letters}{head[0]}" if head[0] == pos[0] else
                                        f"{letters}${head[0]}" if head[1] == pos[1] else f"${letters}${head[0]}")
            english = _text(repair.node, kept) if repair else None
            fixed = verified(sheet, [english], expected, pos, tolerance) if english else ""
        if fixed:
            written = T("rechnet", "calculates") + f" ={_german_formula(_text(node, kept) or cell.formula or '')}"
            if repair.kind == "pair":
                cause = _two(sheet, repair, kept, columns, want, target, wanted, rows)
            else:
                cause = _cause(sheet, repair, kept, columns, want, target, wanted, use, rows, how, written)
            text = T(f"„{label}“ in {address(pos)} {cause} Richtig wäre {_num(expected)} statt {_num(value)}.",
                     f"“{label}” in {address(pos)} {cause} It should be {_num(expected)} instead of {_num(value)}.")
            level = "fehler"
            tip_de, tip_en = _corrections(fixed, expected)
            robust = not any(p.kind == "call" and _pairs(p) for _, p in _paths(repair.node))
            if formula and robust:
                tip_de += f" Sicherer, auch wenn Zeilen dazukommen oder neu sortiert wird: ={formula}."
                tip_en += f" Safer, even when rows are added or the table is re-sorted: ={formula}."
            removed = [e for e in repair.info.get("edits", [repair]) if e.kind == "remove" and e.info["crit"].kind == "str"]
            if removed:
                only = removed[0].info['crit'].value
                tip_de += f" Ist nur „{only}“ gemeint, nenn die Kennzahl „{label} {only}“."
                tip_en += f" If only “{only}” is meant, name the figure “{label} {only}”."
        elif node is None:
            text = T(f"In {address(pos)} steht „{label}“ als feste Zahl {_num(value)}. {belongs}, nachgerechnet ist "
                     f"{_SAID[function]} {_num(expected)}.{f' {why}' if why else ''}",
                     f"{address(pos)} holds “{label}” as the fixed number {_num(value)}. {belongs}; recalculated, "
                     f"{_SAID_EN[function]} is {_num(expected)}.{f' {why}' if why else ''}")
            # Eine feste Zahl kann mit Absicht ein alter Stand sein.
            level = "pruefen"
        elif parts:
            # Liegen die Zeilen nicht beieinander, kann ein einzelner Bereich sie nicht fassen: Das ist die Ursache.
            ranges = [p for _, p in _paths(node) if p.kind == "range"]
            apart = ""
            if len(_runs(wanted)) > 1 and len(ranges) == 1 and not any(p.kind == "call" and _pairs(p) for _, p in _paths(node)):
                apart = T(f" Die Zeilen mit {who or 'diesen Werten'} liegen nicht beieinander ({_spans(target, wanted)}), "
                          "ein einzelner Bereich kann sie nicht fassen.",
                          f" The rows with {who or 'these values'} are not next to each other ({_spans(target, wanted)}), "
                          "so a single range cannot cover them.")
            text = T(f"„{label}“ in {address(pos)} {how}. {belongs}, aber {', und '.join(parts)}.{apart} "
                     f"Richtig wäre {_num(expected)} statt {_num(value)}.",
                     f"“{label}” in {address(pos)} {how}. {belongs}, but {', and '.join(parts)}.{apart} "
                     f"It should be {_num(expected)} instead of {_num(value)}.")
            level = "fehler"
        else:
            text = T(f"„{label}“ in {address(pos)} ergibt {_num(value)}. {belongs}, nachgerechnet ist {_SAID[function]} "
                     f"{_num(expected)}.{f' {why}' if why else ''}",
                     f"“{label}” in {address(pos)} gives {_num(value)}. {belongs}; recalculated, {_SAID_EN[function]} "
                     f"is {_num(expected)}.{f' {why}' if why else ''}")
            level = "fehler" if why else "pruefen"
        tip, fix = _tipped(sheet, f"label:{address(pos)}", tip_de, tip_en)
        found[pos] = Finding(f"label:{address(pos)}", text, level, tip, fix=fix)
    return list(found.values())[:MAX_EACH]


def _corrections(german_formula: str, value: float) -> tuple[str, str]:
    """„Korrigiert: =…, nachgerechnet …“, deutsch und englisch. Daran erkennen einsetzen und das Seitenfenster die
    Korrektur, die der Knopf einsetzt."""
    return (f"Korrigiert: ={german_formula}, nachgerechnet {_num(value)}.",
            f"Corrected: ={german_formula}, recalculated {_num(value)}.")


def _tipped(sheet: Sheet, key: str, german: str, english_tip: str) -> tuple[str, dict | None]:
    """Der Tipp in der Sprache der Anfrage, auf Englisch mit der Korrektur zum Einsetzen: einsetzen.attach liest sie
    sonst erst später aus dem deutschen Tipp („Korrigiert: =…“, „Nimm =…“). So ist sie in beiden Sprachen dieselbe."""
    if not sprache.english():
        return german, None
    from .einsetzen import fix

    return english_tip, (fix(sheet, Finding(key, "", "fehler", german)) if "=" in german else None)


def _measured(header: str, label: str) -> bool:
    """Nennt der Name die Spalte mit einem eigenen Wort („Stunden Nacht“ unter Stunden, „Total EUR“ unter Betrag EUR)?
    „Total“ unter Lohn, Miete und IT meint die Spalte darüber, nicht die Spalte „Total“ rechts daneben. Die Währung
    zählt hier als eigenes Wort: Sie unterscheidet Betrag EUR von Betrag CHF."""
    same = {w for w in _words(header) if woerter.fold(w) in {woerter.fold(v) for v in _words(label)}}
    return bool(same - _ALL - set(_STATS) - (_FILLER - _CURRENCIES) - _MEANS)


def _other_column(sheet: Sheet, calc, pos: Pos, node, formula: str, label: str, want: _Want, columns: dict,
                  data: set[int], target: int, read: int) -> Finding | None:
    """„Stunden Nacht“ =SUMMEWENN(C2:C31;"Nacht";F2:F31): Der Name sagt Stunden, die Formel zählt die Spalte Pause
    zusammen. Gemeldet nur, wenn dieselbe Formel mit der Spalte aus dem Namen genau ergibt, was der Name verlangt."""
    use = _use(sheet, node, read, data)
    if use is None:
        return None
    function = want.function or _FUNCTION[use.kind]
    numbers = columns[target].numbers
    derived = _after(sheet, pos)
    wanted = [row for row in sorted(numbers) if want.matches(columns, row) and (row, target) not in derived]
    if not wanted:
        return None
    expected = _aggregate(function, [numbers[row] for row in wanted])
    tolerance = 0.5 * _rounding(node)
    try:
        value = calc.run(node)
    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
        return None
    if not isinstance(value, float) or isinstance(value, bool) or \
            abs(value - expected) <= tolerance + 1e-9 * max(1.0, abs(expected)):
        return None
    old, new = column_letters(read), column_letters(target)
    kept, areas = _kept(formula), []

    def move(n):
        if n.kind == "range" and n.value[1] == n.value[3] == read:
            r1, _, r2, _ = n.value
            areas.append((r1, r2))
            moved = (r1, target, r2, target)
            if ("range", n.value) in kept:
                kept[("range", moved)] = re.sub(rf"(?<![A-Z]){old}(?![A-Z])", new, kept[("range", n.value)])
            return formel.Node("range", moved)
        if n.kind == "ref" and n.value[1] == read and n.value[0] in data:
            areas.append((n.value[0], n.value[0]))
            moved = (n.value[0], target)
            if ("ref", n.value) in kept:
                kept[("ref", moved)] = re.sub(rf"(?<![A-Z]){old}(?![A-Z])", new, kept[("ref", n.value)])
            return formel.Node("ref", moved)
        if n.kind in ("call", "op", "neg", "pct"):
            return formel.Node(n.kind, n.value, tuple(move(a) if isinstance(a, formel.Node) else a for a in n.args))
        return n

    english = _text(move(node), kept)
    if not english or not areas:
        return None
    fixed = verified(sheet, [english], expected, pos, tolerance)
    if not fixed:
        return None
    plan = sheet.plan
    r1, r2 = areas[0]
    where = lambda col: address((r1, col)) if r1 == r2 else f"{address((r1, col))}:{address((r2, col))}"  # noqa: E731
    text = T(f"„{label}“ in {address(pos)} rechnet mit der Spalte {plan.headers.get(read, old)} ({where(read)}), der Name "
             f"sagt aber {plan.headers.get(target, new)} ({where(target)}). Richtig wäre {_num(expected)} statt "
             f"{_num(value)}.",
             f"“{label}” in {address(pos)} calculates with the column {plan.headers.get(read, old)} ({where(read)}), but "
             f"the name says {plan.headers.get(target, new)} ({where(target)}). It should be {_num(expected)} instead of "
             f"{_num(value)}.")
    tip, fix = _tipped(sheet, f"label:{address(pos)}", *_corrections(fixed, expected))
    return Finding(f"label:{address(pos)}", text, "fehler", tip, fix=fix)


def _after(sheet: Sheet, pos: Pos) -> set[Pos]:
    """Die Formeln, die mit pos rechnen, auch über Formeln dazwischen."""
    seen, todo = set(), [pos]
    while todo and len(seen) < 5000:
        for p in sheet.dependents.get(todo.pop(), ()):
            if p not in seen:
                seen.add(p)
                todo.append(p)
    return seen


def _before(sheet: Sheet, pos: Pos) -> set[Pos]:
    seen, todo = set(), [pos]
    while todo and len(seen) < 5000:
        for p in sheet.precedents.get(todo.pop(), ()):
            if p not in seen:
                seen.add(p)
                todo.append(p)
    return seen


# --- Geteilt durch eine feste Anzahl ---


def _counted(sheet: Sheet, node, depth: int = 0) -> tuple[int, float, str] | None:
    """Wie viele Zahlen eine Summe zusammenzählt (SUMME(D2:D11) → 10), ihre Summe und ein Vorschlag für den
    Mittelwert (englisch, noch nicht nachgerechnet). Auch über eine Zelle (H2/9, wenn in H2 die Summe steht)."""
    if node is None:
        return None
    if node.kind == "ref" and depth < 2 and node.value in sheet.nodes:
        return _counted(sheet, sheet.nodes[node.value], depth + 1)
    while node.kind == "call" and node.value in _ROUNDING and node.args:
        node = node.args[0]
    if node.kind != "call":
        return None
    bottom = sheet.extent(None)[0]
    if node.value == "SUM":
        values = []
        for arg in node.args:
            if arg.kind != "range":
                return None
            r1, c1, r2, c2 = arg.value
            values += [v for (r, col), v in sheet.inputs.items() if r1 <= r <= min(r2, bottom) and c1 <= col <= c2
                       and isinstance(v, float) and not isinstance(v, bool)]
        hint = f"AVERAGE({_area(node.args[0].value)})" if len(node.args) == 1 else ""
        return len(values), math.fsum(values), hint
    if node.value not in ("SUMIF", "SUMIFS"):
        return None
    parts = _ranges_of(node)
    if parts is None:
        return None
    sums, tests = parts
    areas = sums + [t[0] for t in tests]
    if any(a.kind != "range" for a in areas):
        return None
    try:
        checks = [(t[0].value, formel._criterion(_criterion_value(sheet, t[1]))) for t in tests]
    except formel.Unknown:
        return None
    source = (sums[0] if sums else tests[0][0]).value
    picked = []
    for i in range(min(source[2], bottom) - source[0] + 1):
        if all(test(sheet.inputs.get((area[0] + i, area[1]))) for area, test in checks):
            v = sheet.inputs.get((source[0] + i, source[1]))
            if isinstance(v, float) and not isinstance(v, bool):
                picked.append(v)
    hint = ""
    if node.value == "SUMIF" and node.args[1].kind == "str":
        hint = f"AVERAGEIF({_area(tests[0][0].value)},{_quoted(node.args[1].value)},{_area(source)})"
    return len(picked), math.fsum(picked), hint


def _fixed_counts(sheet: Sheet) -> list[Finding]:
    """Eine Summe geteilt durch eine feste Zahl, die nicht mehr zur Anzahl ihrer Zahlen passt (H2/9 bei 10 Zeilen)."""
    found = []
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown:
            continue
        for part in _walk(node):
            if part.kind != "op" or part.value != "/" or len(part.args) != 2 or part.args[1].kind != "num":
                continue
            fixed = part.args[1].value
            counted = _counted(sheet, part.args[0])
            if counted is None and fixed == int(fixed) and fixed not in (100, 1000, 1_000_000):
                share = _fixed_share(sheet, pos, part, fixed)
                if share is not None:
                    found.append(share)
                    break
            if counted is None or fixed != int(fixed) or fixed in (100, 1000, 1_000_000):
                continue
            n, total, hint = counted
            if n < 2 or fixed == n or abs(fixed - n) > max(2, n // 5):
                continue
            label = sheet.figure_label(pos)
            what = (T(f"die Summe aus {address(part.args[0].value)}", f"the sum from {address(part.args[0].value)}")
                    if part.args[0].kind == "ref" else T("die Summe", "the sum"))
            average = re.search(r"durchschnitt|mittel|schnitt|ø|average|mean", label.casefold())
            formula = verified(sheet, [hint], total / n, pos) if hint else ""
            tip, fix = _tipped(sheet, f"count:{address(pos)}", *(
                (f"Nimm ={formula}. Nachgerechnet ergibt das {_num(total / n)}, und es stimmt auch, wenn Zeilen "
                 "dazukommen oder wegfallen.",
                 f"Use ={formula}. Recalculated, that gives {_num(total / n)}, and it stays correct when rows "
                 "are added or removed.") if formula else
                ("Teil durch ANZAHL über denselben Bereich statt durch eine feste Zahl, dann stimmt es auch, wenn "
                 "Zeilen dazukommen oder wegfallen.",
                 "Divide by COUNT over the same range instead of by a fixed number, then it stays correct when "
                 "rows are added or removed.")))
            found.append(Finding(
                f"count:{address(pos)}",
                T(f"{label} teilt {what} durch die feste Zahl {_num(fixed)}, die Summe zählt aber {n} Zahlen zusammen. "
                  f"Geteilt durch {n} wäre es {_num(total / n)}.",
                  f"{label} divides {what} by the fixed number {_num(fixed)}, but the sum adds up {n} numbers. "
                  f"Divided by {n}, it would be {_num(total / n)}."),
                "fehler" if average else "pruefen",
                tip, fix=fix,
            ))
            break
    return found[:MAX_EACH]


def _fixed_share(sheet: Sheet, pos: Pos, part, fixed: float) -> Finding | None:
    """Ein Anteil ZÄHLENWENN(Bereich;…)/71, dessen feste Zahl nicht mehr zur Zahl der Einträge im Bereich passt
    (dort stehen 66 Fälle). Auch über einen Bereich auf einem anderen Blatt."""
    top = part.args[0]
    if top.kind != "call" or top.value not in ("COUNTIF", "COUNTIFS", "COUNTA", "COUNT") or not top.args:
        return None
    area = top.args[0]
    if area.kind not in ("range", "xrange"):
        return None
    calc, _ = sheet.evaluator()
    try:
        values = calc.flat([area])
        share = calc.run(part)
    except (formel.Unknown, RecursionError):
        return None
    if top.value == "COUNT":
        n = sum(1 for v in values if isinstance(v, float) and not isinstance(v, bool))
    else:
        n = sum(1 for v in values if v is not None and v != "")
    if n < 5 or fixed == n or abs(fixed - n) > max(2, n // 5) or not isinstance(share, float):
        return None
    shown = _kept(sheet.cells[pos].formula or "").get((area.kind, area.value)) or T("dem Bereich", "the range")
    label = sheet.figure_label(pos)
    return Finding(
        f"count:{address(pos)}",
        T(f"{label} teilt durch die feste Zahl {_num(fixed)}, in {shown} stehen aber {n} Einträge. "
          f"Geteilt durch {n} wäre es {_num(share * fixed / n)}.",
          f"{label} divides by the fixed number {_num(fixed)}, but {shown} contains {n} entries. "
          f"Divided by {n}, it would be {_num(share * fixed / n)}."),
        "pruefen",
        T("Gehört die feste Zahl zu einer anderen Grundgesamtheit? Sonst teil durch ANZAHL2 über denselben Bereich, "
          "dann stimmt es auch, wenn Zeilen dazukommen oder wegfallen.",
          "Does the fixed number belong to a different population? Otherwise divide by COUNTA over the same range, "
          "then it stays correct when rows are added or removed."),
    )


def _walk(node):
    yield node
    for child in node.args:
        yield from _walk(child)


# --- SUMMEWENN und Verwandte ---


def _criterion_value(sheet: Sheet, node):
    if node.kind == "str":
        return node.value
    if node.kind == "num":
        return float(node.value)
    if node.kind == "ref":
        return sheet.inputs.get(node.value)
    if node.kind in ("op", "call") and not formel.references(node, sheet.name)[1]:
        # „>=“&DATUM(2026;3;1) und ähnliche: ausrechnen.
        try:
            value = sheet.evaluator()[0].run(node)
        except (ValueError, TypeError, ZeroDivisionError) as error:
            raise formel.Unknown("Bedingung") from error
        if isinstance(value, (str, float)) and not isinstance(value, bool):
            return value
    raise formel.Unknown("Bedingung")


def _ranges_of(call) -> tuple[list, list[tuple]] | None:
    """Die Bereiche einer bedingten Formel: wo gezählt wird (falls eigens) und je Bedingung (Bereich, Bedingung)."""
    args = call.args
    if call.value in ("SUMIF", "COUNTIF", "AVERAGEIF"):
        if len(args) < 2:
            return None
        return ([args[2]] if len(args) > 2 else []), [(args[0], args[1])]
    pairs = args if call.value == "COUNTIFS" else args[1:]
    if len(pairs) < 2 or len(pairs) % 2:
        return None
    return ([] if call.value == "COUNTIFS" else [args[0]]), [(pairs[i], pairs[i + 1]) for i in range(0, len(pairs), 2)]


def _conditional(sheet: Sheet, rows: list[int], columns: dict) -> list[Finding]:
    shifted, partial, criteria = [], [], []
    # Zu kurze Bereiche je heruntergezogener Spalte: ein Hinweis für alle ihre Formeln, nicht einer je Zelle.
    short: dict[tuple, list] = {}
    blocks: dict[Pos, list[Pos]] = {}
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown:
            continue
        for call in formel.calls(node):
            if call.value not in _CONDITIONAL:
                continue
            parts = _ranges_of(call)
            if parts is None:
                continue
            sums, tests = parts
            areas = [a for a in sums + [t[0] for t in tests]]
            if any(a.kind != "range" for a in areas):
                continue
            name = f"{_fn(call.value)} in {sheet.figure_label(pos)}"
            starts = {a.value[0] for a in areas}
            if len(starts) > 1:
                first, other = areas[-1].value, next(a.value for a in areas if a.value[0] != areas[-1].value[0])
                gap = other[0] - first[0]
                fixed = (first[0], other[1], first[0] + other[2] - other[0], other[3]) if other[2] < formel.MAX_ROW else None
                shifted.append(Finding(
                    f"shift:{address(pos)}",
                    T(f"In {name} sind die Bereiche gegeneinander verschoben: {_area(first)} und {_area(other)} fangen nicht in "
                      f"derselben Zeile an. So gehört zu jeder Zeile der Wert {abs(gap)} {'Zeile' if abs(gap) == 1 else 'Zeilen'} "
                      f"{'darunter' if gap > 0 else 'darüber'}.",
                      f"The ranges of {name} are shifted against each other: {_area(first)} and {_area(other)} do not start in "
                      f"the same row. So each row gets the value {abs(gap)} {_n(abs(gap), 'row', 'rows')} "
                      f"{'below' if gap > 0 else 'above'}."),
                    "fehler",
                    T(f"Lass alle Bereiche in derselben Zeile anfangen und enden{f' ({_area(first)} und {_area(fixed)})' if fixed else ''}.",
                      f"Make all ranges start and end in the same row{f' ({_area(first)} and {_area(fixed)})' if fixed else ''}."),
                ))
                continue
            try:
                tested = [(t[0].value, formel._criterion(_criterion_value(sheet, t[1])), _criterion_value(sheet, t[1]))
                          for t in tests]
            except formel.Unknown:
                continue
            # Liest die Formel eine andere Tabelle als die des Blatts (Auswertung oben, Rechnungen unten), gilt deren
            # eigene Länge: die Spalte unter ihrem Kopf, so weit sie ohne Lücke gefüllt ist.
            own = rows if _overlaps(tested[0][0], rows) else _own_rows(sheet, tested[0][0])
            hits = _partial_hits(sheet, pos, own, tested, own is not rows)
            if hits:
                block = _block_of(sheet, pos, blocks)
                group = short.setdefault((block[0], call.value, tested[0][0]), [])
                group.append((pos, hits, node, call, tested, areas, own))
            for (area, _, wanted), test in zip(tested, tests):
                criteria += _criterion(sheet, pos, name, area, wanted, test[1])
    for (block, _, _), group in sorted(short.items(), key=lambda item: item[1][0][0]):
        partial.append(_partial(sheet, group[0][6], blocks[block], [item[:6] for item in group]))
    return shifted[:MAX_EACH] + partial[:MAX_EACH] + criteria[:MAX_EACH]


def _block_of(sheet: Sheet, pos: Pos, blocks: dict[Pos, list[Pos]]) -> list[Pos]:
    """Die heruntergezogene Spalte um pos (oder nur pos), einmal je Spalte gesucht: blocks merkt sie sich für jede
    ihrer Zellen."""
    if pos not in blocks:
        from .auswertung import _block, _filled_down

        block = _block(sheet, pos) if _filled_down(sheet, pos) else [pos]
        for cell in block:
            blocks[cell] = block
        blocks.setdefault(pos, [pos])
    return blocks[pos]


def _overlaps(area: Area, rows: list[int]) -> bool:
    return bool(rows) and area[0] <= rows[-1] and area[2] >= rows[0]


def _own_rows(sheet: Sheet, area: Area) -> list[int]:
    """Die Zeilen der Tabelle, aus der eine Spalte kommt: unter ihrem Kopf (Text darüber), lückenlos gefüllt, wie
    die Zellen im Bereich (Formeln oder Werte). Leer, wenn der Bereich schon bis zum Ende reicht."""
    r1, c1, r2, c2 = area
    if c1 != c2 or r2 - r1 < 9 or r2 >= formel.MAX_ROW:
        return []
    head = sheet.cells.get((r1 - 1, c1))
    if head is None or head.formula or not isinstance(head.value, str) or not head.value.strip():
        return []

    def kind(r: int):
        cell = sheet.cells.get((r, c1))
        if cell is None or (not cell.formula and (cell.value is None or str(cell.value).strip() == "")):
            return None
        return "f" if cell.formula else "v"

    kinds = {kind(r) for r in range(r1, r2 + 1)}
    if None in kinds or len(kinds) != 1:
        return []
    end = r2
    while kind(end + 1) in kinds and end - r2 < 100000:
        end += 1
    return list(range(r1, end + 1)) if end > r2 else []


def _partial_hits(sheet: Sheet, pos: Pos, rows: list[int], tested, own: bool = False) -> list[int]:
    """SUMMEWENN prüft nur einen Teil der Tabelle: die Zeilen im Rest, die die Bedingung erfüllen."""
    from .auswertung import _moves

    plan = sheet.plan
    r1, c1, r2, c2 = tested[0][0]
    if (c1 not in plan.headers and not own) or r2 >= formel.MAX_ROW or not rows or (r1 <= rows[0] and r2 >= rows[-1]):
        return []
    # Ein Bereich über eine Zeile (ZÄHLENWENN(B21:G21;"<0") über die Jahre eines Finanzplans) prüft keine Zeilen
    # einer Spalte.
    if c2 > c1 and r1 == r2:
        return []
    if not (r1 <= rows[-1] and r2 >= rows[0]):
        return []
    # Heruntergezogen bis zur eigenen Zeile (ZÄHLENWENN($B$2:B5;B5) für „zum ersten Mal“): Das Ende rutscht mit Absicht.
    if _moves(sheet, pos, tested[0][0]):
        return []
    outside = [r for r in rows if not r1 <= r <= r2]
    return [r for r in outside if all(test(sheet.inputs.get((r, area[1]))) for area, test, _ in tested)]


def _partial(sheet: Sheet, rows: list[int], block: list[Pos], group: list) -> Finding:
    """Ein Hinweis zu einem zu kurzen Bereich: für eine Formel, oder für eine heruntergezogene Spalte zusammen (nach
    neuen Zeilen unten prüfen alle SUMMEWENN($D$2:$D$4001;D2;…) nur bis 4001). Mit der Korrektur, die alle Bereiche
    der Spalte bis ans Ende der Tabelle zieht, wo das sicher geht."""
    from .auswertung import _also, _longer

    pos, hits, node, call, tested, areas = group[0]
    area = tested[0][0]
    r1, c1, r2, _ = area
    name = _fn(call.value)
    every = sorted({r for item in group for r in item[1]})
    whole = ", ".join(f"{column_letters(a.value[1])}:{column_letters(a.value[1])}" for a in areas)
    wanted = (min(r1, rows[0]), max(r2, rows[-1]))
    # Die Bedingung (D4002 in ZÄHLENWENN($D$2:$D$4001;D4002)) zählt nicht mit, nur Bereiche.
    ranges = formel.Node("call", call.value, tuple(a for a in call.args if a.kind == "range"))
    fix = _longer(sheet, pos, call, area, wanted, set(rows), _also(node, ranges, c1, area))
    if fix is not None and not all(p in block for p in (item[0] for item in group)):
        fix = None
    if len(block) == 1:
        n = len(hits)
        text = T(f"{name} in {sheet.figure_label(pos)} prüft nur {_area(area)}, die Tabelle geht aber bis Zeile {rows[-1]}: "
                 f"{n} {'Zeile erfüllt' if n == 1 else 'Zeilen erfüllen'} die Bedingung auch und "
                 f"{'zählt' if n == 1 else 'zählen'} nicht mit ({_cells(c1, hits)}).",
                 f"{name} in {sheet.figure_label(pos)} only checks {_area(area)}, but the table goes down to row {rows[-1]}: "
                 f"{n} more {_n(n, 'row meets', 'rows meet')} the condition and {_n(n, 'is', 'are')} not counted "
                 f"({_cells(c1, hits)}).")
    else:
        col = block[0][1]
        head = sheet.plan.headers.get(col)
        text = T(f"{name} in {address(block[0])} bis {address(block[-1])}{f' ({head})' if head else ''} prüft in jeder "
                 f"Zeile nur {_area(area)}, die Tabelle geht aber bis Zeile {rows[-1]}. ",
                 f"{name} in {address(block[0])} to {address(block[-1])}{f' ({head})' if head else ''} only checks "
                 f"{_area(area)} in every row, but the table goes down to row {rows[-1]}. ")
        lines = f"{'erfüllt' if len(every) == 1 else 'erfüllen'} {_cells(c1, every)}"
        also = (f"{_cells(c1, every)} also {_n(len(every), 'meets', 'meet')} the condition and "
                f"{_n(len(every), 'is', 'are')} not counted")
        if len(group) == 1:
            text += T(f"In {sheet.figure_label(pos)} {lines} die Bedingung auch und {'zählt' if len(every) == 1 else 'zählen'} nicht mit.",
                      f"In {sheet.figure_label(pos)}, {also}.")
        else:
            text += T(f"In {len(group)} dieser Formeln {lines} die Bedingung auch und {'zählt' if len(every) == 1 else 'zählen'} "
                      f"nicht mit, zum Beispiel in {sheet.figure_label(pos)}.",
                      f"In {len(group)} of these formulas, {also}, for example in {sheet.figure_label(pos)}.")
    tip = T(f"Zieh alle Bereiche bis Zeile {rows[-1]}, oder nimm ganze Spalten ({whole}).",
            f"Extend all ranges to row {rows[-1]}, or use whole columns ({whole}).")
    return Finding(f"partial:{address(pos)}", text, "fehler", tip, fix=fix)


def _literal_fix(sheet: Sheet, pos: Pos, node, old: str, new: str) -> tuple[dict | None, str]:
    """Die Korrektur, wenn die Bedingung als Text in der Formel steht („Nrod“ → „Nord“): die Formel mit dem richtigen
    Text, nachgerechnet, und der Satz dazu. (None, "") für eine Bedingung aus einer Zelle oder wenn es nicht sicher ist."""
    from .einsetzen import formulas

    formula = (sheet.cells[pos].formula or "") if pos in sheet.cells else ""
    quoted = '"' + old.replace('"', '""') + '"'
    if node is None or node.kind != "str" or formula.count(quoted) != 1:
        return None, ""
    english = formula.replace(quoted, '"' + new.replace('"', '""') + '"')
    fix = formulas(sheet, {pos: english})
    if fix is None:
        return None, ""
    calc, _ = sheet.evaluator()
    try:
        value = calc.run(formel.parse(english))
    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError):
        return None, ""
    if not isinstance(value, float) or isinstance(value, bool):
        return None, ""
    return fix, T(*_corrections(_german_formula(english), value))


def _criterion(sheet: Sheet, pos: Pos, name: str, area: Area, wanted, node=None) -> list[Finding]:
    """„Nord“ findet „Nord “ mit Leerzeichen nicht, „Nrd“ findet gar nichts. node: wie die Bedingung in der Formel
    steht (ein Text darin bekommt eine Korrektur zum Einsetzen)."""
    if not isinstance(wanted, str) or not wanted.strip() or re.match(r"^(<=|>=|<>|<|>|=)", wanted) or re.search(r"[*?~]", wanted):
        return []
    r1, c1, r2, c2 = area
    target = wanted.casefold()
    if c1 == c2:
        # Eine Spalte: über ein Verzeichnis je Spalte, ohne die Texte jedes Bereichs zu sammeln (ZÄHLENWENN($B$2:B5;B5)
        # heruntergezogen sind 20 000 verschiedene Bereiche).
        rows, folded, trimmed, padded = _column_texts(sheet, c1)
        if bisect_right(rows, r2) == bisect_left(rows, r1):
            return []
        inside = lambda found: found[bisect_left(found, r1):bisect_right(found, r2)]  # noqa: E731
        exact = inside(folded.get(target, []))
        if wanted.strip() == wanted:
            # Ohne Leerzeichen gesucht: Daneben liegen genau die Zellen mit Leerzeichen am Rand.
            near = [(r, c1) for r in inside(padded.get(target, []))]
        else:
            near = [(r, c1) for r in inside(trimmed.get(wanted.strip().casefold(), []))
                    if sheet.shown[(r, c1)].casefold() != target]
        texts = None
    else:
        texts, folded, trimmed = _texts_in(sheet, area)
        if not texts:
            return []
        exact = folded.get(target, [])
        near = sorted(p for p in trimmed.get(wanted.strip().casefold(), []) if texts[p].casefold() != target)
    if near:
        sample = sheet.shown[near[0]]
        where = (T("am Ende", "at the end") if sample.rstrip() != sample
                 else T("am Anfang", "at the start") if sample.lstrip() != sample else "")
        n, cells = len(near), _cells(c1, [p[0] for p in near])
        if wanted.strip() != wanted:
            text = T(f"{name} sucht „{wanted}“ mit Leerzeichen, in {cells} steht aber „{sample}“: "
                     f"{n} {'Zeile zählt' if n == 1 else 'Zeilen zählen'} nicht mit.",
                     f"{name} looks for “{wanted}” with spaces, but {cells} {_n(n, 'contains', 'contain')} “{sample}”: "
                     f"{n} {_n(n, 'row is', 'rows are')} not counted.")
            fix, tip = _literal_fix(sheet, pos, node, wanted, wanted.strip()) if not exact else (None, "")
            if fix is not None:
                return [Finding(f"criterion:{address(pos)}", text.replace("  ", " "), "fehler", tip, fix=fix)]
        else:
            text = T(f"{name} sucht „{wanted}“, in {cells} steht aber „{sample}“ mit einem "
                     f"Leerzeichen {where}: {n} {'Zeile zählt' if n == 1 else 'Zeilen zählen'} nicht mit.",
                     f"{name} looks for “{wanted}”, but {cells} {_n(n, 'contains', 'contain')} “{sample}” with a "
                     f"space {where}: {n} {_n(n, 'row is', 'rows are')} not counted.")
        return [Finding(f"criterion:{address(pos)}", text.replace("  ", " "), "fehler",
                        T("Entfern die Leerzeichen (GLÄTTEN in einer Hilfsspalte, oder Suchen und Ersetzen).",
                          "Remove the spaces (TRIM in a helper column, or Find and Replace)."))]
    if exact:
        return []
    if texts is None and folded.get(target):
        # Der Wert steht in der Spalte, nur nicht im Bereich (ZÄHLENWENN($A$2:$A$4001;A4002) nach einer neuen Zeile):
        # Der Bereich ist zu kurz, nicht die Bedingung falsch geschrieben. Das sagt _partial.
        return []
    if texts is None:
        texts = _texts_in(sheet, area)[0]
    if ("bereiche", "values", area) not in sheet.memo:
        sheet.memo[("bereiche", "values", area)] = {v.strip().casefold(): v.strip() for v in texts.values()}
    values = sheet.memo[("bereiche", "values", area)]
    close = difflib.get_close_matches(target.strip(), list(values), n=1, cutoff=0.75)
    if close:
        meant = values[close[0]]
        fix, tip = _literal_fix(sheet, pos, node, wanted, meant)
        return [Finding(
            f"criterion:{address(pos)}",
            T(f"{name} sucht „{wanted}“, das kommt in {_area(area)} nicht vor, darum zählt sie nichts. Gemeint ist wohl „{meant}“.",
              f"{name} looks for “{wanted}”, which does not occur in {_area(area)}, so it counts nothing. "
              f"“{meant}” is probably meant."),
            "fehler", tip or T(f"Schreib die Bedingung genau so wie in der Tabelle: „{meant}“.",
                               f"Write the condition exactly as in the table: “{meant}”."), fix=fix,
        )]
    return [Finding(
        f"criterion:{address(pos)}",
        T(f"{name} sucht „{wanted}“, das kommt in {_area(area)} nicht vor, darum zählt sie nichts.",
          f"{name} looks for “{wanted}”, which does not occur in {_area(area)}, so it counts nothing."),
        "pruefen", T("Wenn es den Wert noch geben soll, ist das so richtig. Sonst prüf, wie er in der Tabelle geschrieben ist.",
                     "If the value may still turn up, this is fine. Otherwise check how it is written in the table."),
    )]


def _texts_in(sheet: Sheet, area: Area) -> tuple[dict, dict, dict]:
    """Die Texte in einem Bereich und wo sie stehen, nach Kleinschreibung und ohne Leerzeichen am Rand. Einmal je
    Bereich: Eine Spalte mit ZÄHLENWENN($A$2:$A$20001;A2) in jeder Zeile fragt 20 000-mal nach demselben."""
    key = ("bereiche", "texts", area)
    if key not in sheet.memo:
        r1, c1, r2, c2 = area
        texts = {p: v for p, v in sheet.shown.items()
                 if r1 <= p[0] <= r2 and c1 <= p[1] <= c2 and isinstance(v, str) and v.strip()}
        folded: dict[str, list] = {}
        trimmed: dict[str, list] = {}
        for p, v in texts.items():
            folded.setdefault(v.casefold(), []).append(p)
            trimmed.setdefault(v.strip().casefold(), []).append(p)
        sheet.memo[key] = (texts, folded, trimmed)
    return sheet.memo[key]


def _column_texts(sheet: Sheet, col: int) -> tuple[list[int], dict, dict, dict]:
    """Die Zeilen einer Spalte mit Text, und je Text (klein geschrieben, und ohne Leerzeichen am Rand) seine Zeilen,
    dazu nur die Zeilen, in denen er mit Leerzeichen am Rand steht."""
    key = ("bereiche", "column texts", col)
    if key not in sheet.memo:
        rows: list[int] = []
        folded: dict[str, list[int]] = {}
        trimmed: dict[str, list[int]] = {}
        padded: dict[str, list[int]] = {}
        for (r, c), v in sorted(sheet.shown.items()) if col is not None else ():
            if c == col and isinstance(v, str) and v.strip():
                rows.append(r)
                folded.setdefault(v.casefold(), []).append(r)
                trimmed.setdefault(v.strip().casefold(), []).append(r)
                if v.casefold() != v.strip().casefold():
                    padded.setdefault(v.strip().casefold(), []).append(r)
        sheet.memo[key] = (rows, folded, trimmed, padded)
    return sheet.memo[key]


# --- Sortierte Gruppen ohne Namen ---


def _blocks(sheet: Sheet, rows: list[int], col: int) -> list[tuple[str, int, int]]:
    """Zeilen hintereinander mit demselben Text („Nord“ in Zeile 2 bis 11), wie in einer sortierten Tabelle."""
    blocks: list[tuple[str, int, int]] = []
    for r in rows:
        text = sheet.shown.get((r, col))
        if not isinstance(text, str) or not text.strip():
            blocks.append(("", r, r))
        elif blocks and blocks[-1][0] == text.strip() and blocks[-1][2] == r - 1:
            blocks[-1] = (blocks[-1][0], blocks[-1][1], r)
        else:
            blocks.append((text.strip(), r, r))
    return [b for b in blocks if b[0]]


def group(sheet: Sheet, rows: list[int], r1: int, r2: int, col: int) -> tuple[int, str, int, int] | None:
    """Fängt ein Bereich genau bei einer Gruppe an (Region „Nord“), hört aber eine oder zwei Zeilen zu früh
    oder zu spät auf? Passt er in einer anderen Spalte genau auf eine Gruppe, ist er Absicht."""
    near = None
    for other in sorted(sheet.plan.headers):
        if other == col:
            continue
        for name, top, bottom in _blocks(sheet, rows, other):
            if bottom - top < 2:
                continue
            if (top, bottom) == (r1, r2):
                return None
            if near is None and (top == r1 and 1 <= abs(bottom - r2) <= 2 or bottom == r2 and 1 <= abs(top - r1) <= 2):
                near = (other, name, top, bottom)
    return near


def group_finding(sheet: Sheet, pos: Pos, function: str, area: Area, other: int, name: str, top: int, bottom: int) -> Finding:
    r1, c1, r2, _ = area
    missing = [r for r in range(top, bottom + 1) if not r1 <= r <= r2]
    extra = [r for r in range(r1, r2 + 1) if not top <= r <= bottom]
    if missing:
        n = len(missing)
        what = T(f"{n} {'Zeile' if n == 1 else 'Zeilen'} mit „{name}“ "
                 f"{'fehlt' if n == 1 else 'fehlen'} ({_cells(c1, missing)})",
                 f"{n} {_n(n, 'row', 'rows')} with “{name}” {_n(n, 'is', 'are')} missing ({_cells(c1, missing)})")
    else:
        n = len(extra)
        what = T(f"{_cells(c1, extra)} {'gehört' if n == 1 else 'gehören'} schon nicht mehr zu „{name}“",
                 f"{_cells(c1, extra)} no longer {_n(n, 'belongs', 'belong')} to “{name}”")
    letter, key = column_letters(c1), column_letters(other)
    tip_de, tip_en = f"Nimm {address((top, c1))}:{address((bottom, c1))}.", f"Use {address((top, c1))}:{address((bottom, c1))}."
    values = [v for r in range(top, bottom + 1) if isinstance(v := sheet.inputs.get((r, c1)), float) and not isinstance(v, bool)]
    if function in ("SUM", "AVERAGE") and values:
        expected = _aggregate(function, values)
        english = f"{'SUMIF' if function == 'SUM' else 'AVERAGEIF'}({key}:{key},{_quoted(name)},{letter}:{letter})"
        if formula := verified(sheet, [english], expected, pos):
            tip_de += (f" Sicherer ist ={formula}, nachgerechnet {_num(expected)}. Das stimmt auch, wenn Zeilen dazukommen "
                       "oder anders sortiert sind.")
            tip_en += (f" Safer is ={formula}, recalculated {_num(expected)}. That stays correct even when rows are added "
                       "or sorted differently.")
    tip, fix = _tipped(sheet, f"group:{address(pos)}", tip_de, tip_en)
    return Finding(
        f"group:{address(pos)}",
        T(f"{_german(function)} in {sheet.figure_label(pos)} nimmt {address((r1, c1))}:{address((r2, c1))}. In {sheet.plan.headers[other]} "
          f"steht „{name}“ aber in den Zeilen {top} bis {bottom}: {what}.",
          f"{function} in {sheet.figure_label(pos)} takes {address((r1, c1))}:{address((r2, c1))}. But in {sheet.plan.headers[other]}, "
          f"“{name}” is in rows {top} to {bottom}: {what}."),
        "fehler", tip, fix=fix,
    )


# --- Bereiche, die beim Herunterkopieren mitrutschen ---


# Funktionen, deren Bereich meist die ganze Tabelle sein soll. Nachschlagen (SVERWEIS …) prüft analyse.py.
_WHOLE = _PLAIN | _CONDITIONAL | {"SUMPRODUCT", "RANK", "RANK.EQ", "RANK.AVG", "LARGE", "SMALL", "STDEV", "STDEV.S",
                                  "STDEV.P", "VAR", "VAR.S", "VAR.P", "PERCENTILE", "PERCENTILE.INC", "QUARTILE",
                                  "QUARTILE.INC", "PERCENTRANK", "PERCENTRANK.INC"}
_RANGE_TEXT = re.compile(r"(?<![A-Za-z0-9_.$!\\])(\$?)([A-Z]{1,3})(\$?)(\d{1,7}):(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(!])")


def _copied_runs(sheet: Sheet) -> list[list[Pos]]:
    """Zellen untereinander mit derselben Formel, so wie Excel sie herunterzieht (gleiche Z1S1-Form)."""
    by_col: dict[int, list[int]] = {}
    for (r, c), cell in sheet.cells.items():
        if cell.formula is not None and (r, c) in sheet.nodes and (r, c) not in sheet.unknown:
            by_col.setdefault(c, []).append(r)
    runs = []
    for c, rows in sorted(by_col.items()):
        run: list[Pos] = []
        for r in sorted(rows):
            here = pattern(sheet.cells[(r, c)].formula, (r, c))
            if run and r == run[-1][0] + 1 and here == pattern(sheet.cells[run[-1]].formula, run[-1]):
                run.append((r, c))
            else:
                if len(run) >= 2:
                    runs.append(run)
                run = [(r, c)]
        if len(run) >= 2:
            runs.append(run)
    return runs


def _anchor(formula: str, areas: set[Area]) -> str:
    """D2:D31 → D$2:D$31, nur für die Bereiche in areas: Die Spalte darf mitwandern, die Zeilen nicht."""
    def fix(match: re.Match) -> str:
        dc1, l1, dr1, n1, dc2, l2, dr2, n2 = match.groups()
        area = (int(n1), column_number(l1), int(n2), column_number(l2))
        if dr1 or dr2 or area not in areas:
            return match.group(0)
        return f"{dc1}{l1}${n1}:{dc2}{l2}${n2}"

    return outside_literals(formula, lambda part: _RANGE_TEXT.sub(fix, part))


def _same(a, b) -> bool:
    if isinstance(a, float) and isinstance(b, float) and not isinstance(a, bool) and not isinstance(b, bool):
        return abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))
    return a == b


def sliding(sheet: Sheet, rows: list[int]) -> tuple[list[Finding], set[str]]:
    """SUMME(D2:D31), SUMMEWENN(B2:B31;…) oder RANG(D2;D2:D31) ohne $ heruntergezogen: In der nächsten Zeile
    nimmt die Formel D3:D32, ganz unten nur noch die letzten Zeilen. Gibt die Hinweise und die Zellen, die sie
    erklären (andere Hinweise zu diesen Zellen sagen dasselbe ungenauer)."""
    if len(rows) < 3:
        return [], set()
    plan = sheet.plan
    calc, _ = sheet.evaluator()
    found, explained = [], set()
    for run in _copied_runs(sheet):
        first = run[0]
        node = sheet.nodes[first]
        areas = set()
        for call in formel.calls(node):
            if call.value not in _WHOLE:
                continue
            for arg in call.args:
                if arg.kind != "range":
                    continue
                r1, c1, r2, c2 = arg.value
                # Der Bereich der ersten Zelle nimmt die Tabelle von ihrer ersten Zeile an, mindestens zur Hälfte.
                if c1 not in plan.headers or r2 >= formel.MAX_ROW or not (r1 <= rows[0] and r2 >= rows[len(rows) // 2]):
                    continue
                if any(r1 <= p[0] <= r2 and c1 <= p[1] <= c2 for p in run):
                    continue
                areas.add(arg.value)
        if not areas:
            continue
        fixed = _anchor(sheet.cells[first].formula, areas)
        if fixed == sheet.cells[first].formula:
            continue  # Die Zeilen haben schon $.
        wrong, sample = [], None
        for pos in run[1:]:
            try:
                now = calc.run(sheet.nodes[pos])
                right = calc.run(formel.parse(shift(fixed, pos[0] - first[0], 0)))
            except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
                continue
            if not _same(now, right):
                wrong.append(pos)
                sample = sample or (pos, now, right)
        last = run[-1]
        area = sorted(areas)[0]
        moved = (area[0] + last[0] - first[0], area[1], area[2] + last[0] - first[0], area[3])
        lost = [r for r in rows if r < moved[0]]
        where = f"{address(run[1])}:{address(last)}" if len(run) > 2 else address(run[1])
        name = next((c.value for c in formel.calls(node) if any(a.kind == "range" and a.value in areas for a in c.args)), "")
        second = _area((area[0] + run[1][0] - first[0], area[1], area[2] + run[1][0] - first[0], area[3]))
        n = len(lost)
        head = T(f"{where} {'sind' if len(run) > 2 else 'ist'} aus {address(first)} heruntergezogen, und der Bereich "
                 f"{_area(area)} in {_german(name)} rutscht dabei mit, weil vor den Zeilen die $-Zeichen fehlen: "
                 f"{address(run[1])} nimmt {second}"
                 + (f", {address(last)} schon {_area(moved)}" if len(run) > 2 else "")
                 + f". Ganz unten in {address(last)} {'fehlt' if n == 1 else 'fehlen'} so schon {n} "
                 f"{'Zeile' if n == 1 else 'Zeilen'} vom Anfang der Tabelle.",
                 f"{where} {'are' if len(run) > 2 else 'is'} filled down from {address(first)}, and the range "
                 f"{_area(area)} in {name} moves along, because the row numbers have no $ signs: "
                 f"{address(run[1])} takes {second}"
                 + (f", {address(last)} already {_area(moved)}" if len(run) > 2 else "")
                 + f". At the bottom, in {address(last)}, {n} {_n(n, 'row', 'rows')} from the start of the table "
                 f"{_n(n, 'is', 'are')} already missing.")
        if wrong:
            pos, now, right = sample
            k = len(wrong)
            text = T(f"{head} {k} von {len(run) - 1} Ergebnissen {'ist' if k == 1 else 'sind'} dadurch "
                     f"falsch, zum Beispiel {sheet.figure_label(pos)}: {_fmt(now)} statt {_fmt(right)}.",
                     f"{head} As a result, {k} of {len(run) - 1} results {_n(k, 'is', 'are')} wrong, for example "
                     f"{sheet.figure_label(pos)}: {_fmt(now)} instead of {_fmt(right)}.")
            level = "fehler"
        else:
            text = T(f"{head} Heute ändert das zufällig nichts, weil in den fehlenden Zeilen nichts steht, was mitzählt. "
                     "Wird die Tabelle anders sortiert oder kommen Zeilen dazu, stimmt es nicht mehr.",
                     f"{head} Right now this happens to change nothing, because the missing rows contain nothing that "
                     "counts. If the table is sorted differently or rows are added, it will no longer be correct.")
            level = "pruefen"
        key = f"slide:{address(sample[0] if wrong else run[1])}"
        tip, fix = _tipped(sheet, key,
                           f"Schreib in {address(first)} ={_german_formula(fixed)} und zieh es wieder herunter (die $-Zeichen "
                           f"setzt F4, am Mac cmd+T). Dann nimmt jede Zelle die ganze Tabelle {_area(area)}.",
                           f"In {address(first)}, enter ={_german_formula(fixed)} and fill it down again (F4 adds the $ "
                           f"signs, cmd+T on a Mac). Then every cell takes the whole table {_area(area)}.")
        found.append(Finding(key, text, level, tip, fix=fix))
        explained |= {address(p) for p in run}
        if len(found) >= 2:
            break
    return found, explained


_ROLLING = re.compile(r"gleitend|rollierend|rollend|laufend|rolling|moving|trailing|letzten? \d+|last \d+|\(\d+ ?(?:mon|monate|"
                      r"tage|wochen|jahre|perioden|quartale|m|t|w|j|q)\b|\bma\d*\b|\bsma\b|mobile|glissant|mobile|scorrevole")
_FIXED_SPAN = re.compile(r"basis|referenz|bezug|baseline|ausgangs|start|anfang|q[1-4]\b|quartal \d|jan(?:uar)?\s*(?:-|bis|–)"
                         r"|fest|fix")


def frozen(sheet: Sheet, rows: list[int]) -> tuple[list[Finding], set[str]]:
    """Das Gegenstück zu sliding(): ein gleitender Mittelwert (die letzten 3 Monate), dessen Bereich $B$5:$B$7 mit
    $-Zeichen festgehalten ist. Heruntergezogen nimmt dann jede Zeile dieselben 3 Monate, die Spalte zeigt überall
    dieselbe Zahl. Gemeint ist ein Fenster, das mitwandert: in der ersten Zeile B5:B7, darunter B6:B8 usw."""
    if len(rows) < 4:
        return [], set()
    plan = sheet.plan
    data = set(rows)
    calc, _ = sheet.evaluator()
    found, explained = [], set()
    for run in _copied_runs(sheet):
        first, last = run[0], run[-1]
        if len(run) < 3 or first[0] not in data or first[1] not in plan.headers:
            continue
        text = sheet.cells[first].formula or ""
        node = sheet.nodes[first]
        areas = []
        for call in formel.calls(node):
            if call.value not in _WHOLE:
                continue
            for arg in call.args:
                if arg.kind == "range" and arg.value[1] == arg.value[3] and arg.value[1] in plan.headers:
                    areas.append((call.value, arg.value))
        if len(areas) != 1:
            continue
        name, area = areas[0]
        r1, col, r2, _ = area
        size = r2 - r1 + 1
        # Ein kurzes Stück der Tabelle (nicht die ganze Spalte), in allen Zeilen gleich, weil beide Zeilen ein $ haben.
        if not 2 <= size <= max(2, len(rows) // 2) or not (r1 in data and r2 in data) or col == first[1]:
            continue
        held = [m for m in _RANGE_TEXT.finditer(text)
                if (int(m.group(4)), column_number(m.group(2)), int(m.group(8)), column_number(m.group(6))) == area]
        if len(held) != 1 or not (held[0].group(3) and held[0].group(7)):
            continue
        header = plan.headers[first[1]]
        words = header.casefold()
        named = bool(_ROLLING.search(words)) and not _FIXED_SPAN.search(words)
        # Ohne Wort im Namen: Das Fenster endet genau in der Zeile der ersten Formel (die 3 Monate bis hierher), und
        # die Formeln fangen erst an, wenn es voll ist. Das tut nur ein gleitender Mittelwert.
        aligned = r2 == first[0] and r1 == rows[0] and first[0] - rows[0] == size - 1
        if not (named or aligned) or _FIXED_SPAN.search(words):
            continue
        offset = (r1 - first[0], r2 - first[0])
        loose = text[:held[0].start()] + f"{held[0].group(2)}{held[0].group(4)}:{held[0].group(6)}{held[0].group(8)}" \
            + text[held[0].end():]
        changes, wrong, sample = {}, [], None
        for pos in run:
            moved = shift(loose, pos[0] - first[0], 0)
            lo, hi = pos[0] + offset[0], pos[0] + offset[1]
            if lo not in data or hi not in data or "#" in moved:
                changes = {}
                break
            changes[pos] = moved
            try:
                now = calc.run(sheet.nodes[pos])
                right = calc.run(formel.parse(moved))
            except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
                changes = {}
                break
            if not _same(now, right):
                wrong.append(pos)
                sample = (pos, now, right, lo, hi)
        if not changes or not wrong:
            continue
        from . import einsetzen

        fix = einsetzen.formulas(sheet, changes)
        pos, now, right, lo, hi = sample
        span = (lambda a, b: T(f"{plan.names[a]} bis {plan.names[b]}", f"{plan.names[a]} to {plan.names[b]}")
                if a in plan.names and b in plan.names else "")
        said = span(r1, r2)
        there = span(lo, hi)
        where = f"{address(first)}:{address(last)}"
        why = (T(f"„{header}“ ist ein Mittelwert, der mitwandert", f"“{header}” is a moving average") if named
               else T(f"Das Fenster endet in {address(first)} genau in der eigenen Zeile, wie bei einem gleitenden Mittelwert",
                      f"The window in {address(first)} ends exactly in its own row, as with a moving average"))
        k = len(wrong)
        text_out = T(f"{where} ({header}) {'nehmen' if len(run) > 1 else 'nimmt'} in jeder Zeile denselben Bereich "
                     f"{held[0].group(0)}{f' ({said})' if said else ''}: Die $-Zeichen halten ihn fest, darum steht dort überall "
                     f"{_fmt(calc.run(node))}. {why}, also nimmt jede Zeile die {size} Zeilen bis zu ihrer eigenen, "
                     f"{address(pos)} etwa {_area((lo, col, hi, col))}{f' ({there})' if there else ''}. Richtig wäre dort "
                     f"{_fmt(right)} statt {_fmt(now)}; {k} von {len(run)} Zellen sind falsch.",
                     f"{where} ({header}) {'take' if len(run) > 1 else 'takes'} the same range "
                     f"{held[0].group(0)}{f' ({said})' if said else ''} in every row: the $ signs hold it in place, so it "
                     f"shows {_fmt(calc.run(node))} everywhere. {why}, so each row should take the {size} rows up to its own, "
                     f"{address(pos)} for example {_area((lo, col, hi, col))}{f' ({there})' if there else ''}. It should be "
                     f"{_fmt(right)} there instead of {_fmt(now)}; {k} of {len(run)} {_n(k, 'cell is', 'cells are')} wrong.")
        tip = T(f"Schreib in {address(first)} ={_german_formula(changes[first])} ohne $ und zieh es herunter bis {address(last)}, "
                f"dann wandert der Bereich mit.",
                f"In {address(first)}, enter ={_german_formula(changes[first])} without $ and fill it down to {address(last)}, "
                f"then the range moves along.")
        item = Finding(f"window:{address(first)}", text_out, "fehler" if named else "pruefen", tip, fix=fix)
        found.append(item)
        explained |= {address(p) for p in run}
        if len(found) >= 2:
            break
    return found, explained


# --- Eine Formel in Worten: welche Zeilen sie nimmt (für die Herleitung) ---


def _month_phrase(pairs: set[tuple[int, int]]) -> str:
    """{(2026, 2)} → „Februar 2026“, {(2026, 1), (2026, 2), (2026, 3)} → „Januar bis März 2026“."""
    years = {y for y, _ in pairs}
    if len(years) != 1:
        return ""
    months = {m for _, m in pairs}
    if len(_runs(list(months))) > 1 and len(months) > 2:
        return ""
    return f"{_months_text(months)} {years.pop()}"


def _filters(sheet: Sheet, columns: dict, rows: set[int], data: list[int]) -> list[tuple[str, set[int], bool]]:
    """Was alle diese Zeilen gemeinsam haben, je mit allen Zeilen der Tabelle, auf die es zutrifft:
    ein Wert in einer Textspalte („Region „Ost““) oder die Monate oder Tage in einer Datumsspalte. Das dritte
    Feld sagt, ob es sich mit einem anderen verbinden lässt: Eine Spanne von Tagen passt sonst zu oft zufällig."""
    found = []
    for col, info in sorted(columns.items()):
        if info.kind == "text":
            values = {info.texts.get(r, "").strip() for r in rows}
            if len(values) == 1 and (value := values.pop()):
                found.append((_value(sheet.plan.headers[col], value),
                              {r for r in data if info.texts.get(r, "").strip() == value}, True))
        elif info.kind == "datum" and all(r in info.dates for r in rows):
            days = {r: _day(info.dates[r]) for r in data if r in info.dates}
            months = {(days[r].year, days[r].month) for r in rows}
            phrase = _month_phrase(months)
            head = sheet.plan.headers[col]
            if phrase:
                found.append((T(f"{head} im {phrase}", f"{head} in {phrase}"),
                              {r for r, d in days.items() if (d.year, d.month) in months}, True))
            low, high = min(days[r] for r in rows), max(days[r] for r in rows)
            if low != high:
                found.append((T(f"{head} vom {low:%d.%m.%Y} bis {high:%d.%m.%Y}", f"{head} from {low:%d.%m.%Y} to {high:%d.%m.%Y}"),
                              {r for r, d in days.items() if low <= d <= high}, False))
    return found


def _named(sheet: Sheet, columns: dict, rows: set[int], data: list[int], target: int,
           conditional: bool) -> tuple[str, str] | None:
    """Die Zeilen so, wie man sie nennen würde, im Dativ und Akkusativ: („den 12 Zeilen mit Region „Ost““, „die 12 …“).
    Bei einer Formel mit Bedingungen zählen zwei Bedingungen zusammen („Ost im Februar“) vor „fast eine Gruppe“, bei
    einem festen Bereich ist „alle mit Ost außer D10“ die bessere Beschreibung."""
    n = len(rows)
    if rows == set(data):
        return T(f"allen {n} Zeilen der Tabelle", f"all {n} rows of the table"), T(f"alle {n} Zeilen der Tabelle", f"all {n} rows of the table")

    filters = _filters(sheet, columns, rows, data)

    def exact(size: int) -> tuple[str, str] | None:
        for combo in itertools.combinations(filters, size):
            if (size == 1 or all(joins for _, _, joins in combo)) and set.intersection(*(m for _, m, _ in combo)) == rows:
                what = T(" und ", " and ").join(text for text, _, _ in combo)
                return T(f"den {n} Zeilen mit {what}", f"the {n} rows with {what}"), T(f"die {n} Zeilen mit {what}", f"the {n} rows with {what}")
        return None

    found = (exact(1) or (exact(2) if conditional else None) or _all_but(rows, data, target)
             or _nearly(sheet, columns, rows, data, target))
    return found or (None if conditional else exact(2))


def _all_but(rows: set[int], data: list[int], target: int) -> tuple[str, str] | None:
    """Fast die ganze Tabelle: „die 37 Zeilen E2:E38, also alle außer E39:E41“."""
    missing = set(data) - rows
    if not missing or rows - set(data) or len(missing) > 5 or len(rows) < 2 * len(missing):
        return None
    spans, but = _spans(target, sorted(rows)), _spans(target, sorted(missing))
    if len(_runs(sorted(rows))) > 2:
        return None
    n = len(rows)
    return (T(f"den {n} Zeilen {spans}, also allen außer {but}", f"the {n} rows {spans}, that is all except {but}"),
            T(f"die {n} Zeilen {spans}, also alle außer {but}", f"the {n} rows {spans}, that is all except {but}"))


def _nearly(sheet: Sheet, columns: dict, rows: set[int], data: list[int], target: int) -> tuple[str, str] | None:
    """Fast eine Gruppe: alle Zeilen mit Region „Ost“ außer einer, oder dazu eine aus einer anderen Gruppe."""
    n = len(rows)
    best = None
    for col, info in sorted(columns.items()):
        if info.kind != "text":
            continue
        for value in {info.texts.get(r, "").strip() for r in rows} - {""}:
            group = {r for r in data if info.texts.get(r, "").strip() == value}
            missing, extra = sorted(group - rows), sorted(rows - group)
            if len(missing) + len(extra) <= 2 and len(group & rows) >= 3 * (len(missing) + len(extra)):
                key = len(missing) + len(extra)
                if best is None or key < best[0]:
                    best = (key, col, value, missing, extra)
                elif key == best[0]:
                    best = (key, None, None, [], [])
    if best and best[1] is not None:
        _, col, value, missing, extra = best
        head = _value(sheet.plan.headers[col], value)
        rest = []
        if missing:
            rest.append(T("außer ", "except ") + _listing([address((r, target)) for r in missing]))
        if extra:
            rest.append(T("dazu ", "plus ") + _listing([f"{address((r, target))} ({columns[col].texts.get(r, '').strip() or T('leer', 'empty')})"
                                                       for r in extra]))
        what = T(f"{head} {' und '.join(rest)}" if missing else f"{head} und {rest[0]}", f"{head} {', '.join(rest)}" if missing else f"{head}, {rest[0]}")
        return T(f"den {n} Zeilen mit {what}", f"the {n} rows with {what}"), T(f"die {n} Zeilen mit {what}", f"the {n} rows with {what}")
    return None


def _rows_phrase(sheet: Sheet, columns: dict, taken: set[int], data: list[int], target: int,
                 conditional: bool = False) -> tuple[str, str]:
    """Dativ und Akkusativ für die Zeilen, die eine Formel nimmt, mit den Zellen dazu, wo es hilft."""
    inside = {r for r in taken if r in set(data)}
    outside = sorted(taken - inside)
    # Leere Zellen nimmt keine Formel mit: Verglichen wird mit den Zeilen, in denen eine Zahl steht.
    numbered = [r for r in data if _number_at(sheet, (r, target))]
    if inside <= set(numbered):
        data = numbered
    if len(taken) == 1:
        only = address((next(iter(taken)), target))
        return only, only
    spans = _spans(target, sorted(inside)) if inside else ""
    named = _named(sheet, columns, inside, data, target, conditional) if len(inside) >= 2 else None
    if named:
        dative, accusative = named
        if len(_runs(sorted(inside))) <= 2 and spans not in dative:
            dative, accusative = f"{dative} ({spans})", f"{accusative} ({spans})"
    elif len(_runs(sorted(inside))) > 4:
        low, high = address((min(inside), target)), address((max(inside), target))
        dative = accusative = T(f"{len(inside)} Zeilen zwischen {low} und {high}", f"{len(inside)} rows between {low} and {high}")
    elif inside:
        dative = accusative = spans
    else:
        dative = accusative = ""
    if outside:
        named_rows = []
        for r in outside[:MAX_LISTED]:
            word = sheet.plan.word(r) or _left_label(sheet.cells, r, min(sheet.plan.headers or [1]) + 1)
            named_rows.append(f"{address((r, target))} ({word})" if word else address((r, target)))
        more = _listing(named_rows) + (_more(len(outside) - MAX_LISTED) if len(outside) > MAX_LISTED else "")
        note = T(f"{more}, die nicht zur Tabelle {'gehört' if len(outside) == 1 else 'gehören'}",
                 f"{more}, which {_n(len(outside), 'does', 'do')} not belong to the table")
        dative = T(f"{dative} und dazu {note}", f"{dative}, plus {note}") if dative else note
        accusative = T(f"{accusative} und dazu {note}", f"{accusative}, plus {note}") if accusative else note
    return dative, accusative


def _factors(sheet: Sheet, columns: dict, use: _Use, target: int) -> str:
    """Was in einer Summe nicht einfach einmal zählt: „D11 zählt doppelt, D12 wird abgezogen“."""
    doubled = sorted(r for r, k in use.rows.items() if k == 2.0)
    subtracted = sorted(r for r, k in use.rows.items() if k < 0)
    other = sorted(r for r, k in use.rows.items() if k not in (1.0, 2.0) and k > 0)
    parts = []
    if doubled:
        listed, n = _listing([address((r, target)) for r in doubled]), len(doubled)
        parts.append(T(f"{listed} {'zählt' if n == 1 else 'zählen'} doppelt", f"{listed} {_n(n, 'counts', 'count')} twice"))
    if subtracted:
        listed, n = _listing([address((r, target)) for r in subtracted]), len(subtracted)
        parts.append(T(f"{listed} {'wird' if n == 1 else 'werden'} abgezogen", f"{listed} {_n(n, 'is', 'are')} subtracted"))
    if other:
        parts.append(_listing([T(f"{address((r, target))} mal {_num(use.rows[r])}", f"{address((r, target))} times {_num(use.rows[r])}")
                               for r in other]))
    return ", ".join(parts)


def _sentence(sheet: Sheet, columns: dict, use: _Use, target: int, data: list[int], conditional: bool) -> str:
    header = sheet.plan.headers[target]
    taken = {r for r, k in use.rows.items() if k}
    # Eine einzelne Zelle (=D5) sagt die Formel schon selbst.
    if not taken or (len(taken) == 1 and use.kind == "sum" and not conditional):
        return ""
    dative, accusative = _rows_phrase(sheet, columns, taken, data, target, conditional)
    factors = set(use.rows.values()) - {0.0}
    if use.kind == "sum":
        if factors == {1.0}:
            return T(f"Summiert {header} aus {dative}.", f"Sums {header} from {dative}.")
        if len(factors) == 1:
            factor = _num(factors.pop())
            return T(f"Summiert {header} aus {dative} und nimmt das mal {factor}.",
                     f"Sums {header} from {dative} and multiplies it by {factor}.")
        factored = _factors(sheet, columns, use, target)
        return T(f"Rechnet mit {header} aus {dative}: {factored}.", f"Calculates with {header} from {dative}: {factored}.")
    if use.kind == "count":
        return T(f"Zählt {accusative}.", f"Counts {accusative}.") if len(taken) > 1 else T(f"Zählt {dative}.", f"Counts {dative}.")
    if use.kind == "ratio":
        if use.divisor is not None:
            n = len(taken)
            same = (T(" (so viele Zeilen sind es)", " (that is how many rows there are)") if use.divisor == n
                    else T(f", es sind aber {n} Zeilen", f", but there {_n(n, 'is', 'are')} {n} {_n(n, 'row', 'rows')}"))
            return T(f"Teilt die Summe von {header} aus {dative} durch die feste Zahl {_num(use.divisor)}{same}.",
                     f"Divides the sum of {header} from {dative} by the fixed number {_num(use.divisor)}{same}.")
        counted = set(use.counted or [])
        by = T("dieselben Zeilen", "the same rows") if counted == taken else \
            _rows_phrase(sheet, columns, counted, data, target, conditional)[1]
        return T(f"Teilt die Summe von {header} aus {dative} durch die Anzahl ({len(counted)}) für {by}.",
                 f"Divides the sum of {header} from {dative} by the count ({len(counted)}) of {by}.")
    said = T({"mean": "den Mittelwert", "min": "das Minimum", "max": "das Maximum", "median": "den Median"}[use.kind],
             {"mean": "the average", "min": "the minimum", "max": "the maximum", "median": "the median"}[use.kind])
    return T(f"Nimmt {said} von {header} aus {dative}.", f"Takes {said} of {header} from {dative}.")


def in_words(sheet: Sheet, pos: Pos, rows: list[int], columns: dict) -> str:
    """Welche Zeilen der Tabelle die Formel in pos nimmt, in einem Satz: „Summiert Umsatz (€) aus den 12 Zeilen mit
    Region „Ost“ (D2:D13).“ So sieht man auch ohne Hinweis, ob eine Formel rechnet, was sie soll. "" wenn unklar."""
    node = sheet.nodes.get(pos)
    if node is None or pos in sheet.unknown or not rows or sheet.plan.is_derived(*pos):
        return ""
    data = set(rows)
    # Eine Formel, die nur mit ihrer eigenen Zeile rechnet (=D5*1,19), sagt das schon selbst.
    if all(p[0] == pos[0] for p in sheet.precedents.get(pos, ())):
        return ""
    targets = sorted(_reach(sheet, pos, data, columns))
    if node.kind == "op" and node.value == "/" and len(node.args) == 2 and len(targets) == 2:
        # Eine Quote aus zwei Spalten: Umsatz durch Menge.
        parts = []
        for side in node.args:
            for target in targets:
                use = _use(sheet, side, target, data)
                if use and use.kind == "sum" and set(use.rows.values()) == {1.0}:
                    parts.append((target, {r for r, k in use.rows.items() if k}))
                    break
        if len(parts) == 2 and parts[0][0] != parts[1][0] and parts[0][1] and parts[1][1]:
            (top, above), (bottom, below) = parts
            where = _rows_phrase(sheet, columns, above, rows, top)[0]
            other = T("denselben Zeilen", "the same rows") if above == below else _rows_phrase(sheet, columns, below, rows, bottom)[0]
            return T(f"Teilt die Summe von {sheet.plan.headers[top]} aus {where} durch die Summe von "
                     f"{sheet.plan.headers[bottom]} aus {other}.",
                     f"Divides the sum of {sheet.plan.headers[top]} from {where} by the sum of "
                     f"{sheet.plan.headers[bottom]} from {other}.")
        return ""
    # Zählen geht auch über Textspalten (ZÄHLENWENN(B:B;"Ost")): dann zeigt er die Zellen der Spalte, die er prüft.
    tested = next((part.value[1] for _, part in _paths(node) if part.kind == "range" and part.value[1] in sheet.plan.headers), None)
    for target in targets or ([tested] if tested is not None else []):
        use = _use(sheet, node, target, data)
        if use is None or (not targets and use.kind != "count"):
            continue
        return _sentence(sheet, columns, use, target, rows, any(c.value in _CONDITIONAL for c in formel.calls(node)))
    return ""
