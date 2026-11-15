"""Excel-Formeln selbst lesen und nachrechnen, ohne Excel und ohne KI.

Ein kleiner Parser für die Formeln, wie sie in der Datei stehen (englische Funktionsnamen, Komma als
Trenner, ohne „=“), und ein Rechner für die üblichen Funktionen. Was er nicht kennt (andere Blätter,
Namen, SVERWEIS …), gibt None: Dann gilt der Wert, den Excel zeigt, und es wird nichts behauptet.
"""

from __future__ import annotations

import calendar
import math
import re
import statistics
import threading
from bisect import bisect_left, bisect_right
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, timedelta

from .xlsx import column_letters, column_number

Pos = tuple[int, int]


class Unknown(Exception):
    """Diese Formel kann der Rechner nicht (sicher) nachrechnen."""


@dataclass(frozen=True)
class Err:
    """Ein Fehlerwert wie #DIV/0!."""

    code: str


@dataclass(frozen=True)
class Node:
    kind: str  # num, str, bool, err, ref, range, op, neg, pct, call, unknown
    value: object = None
    args: tuple = ()


_TOKEN = re.compile(
    r"""\s*(?:
    (?P<str>"(?:[^"]|"")*")
  | (?P<err>\#(?:DIV/0!|N/A|NV|VALUE!|WERT!|REF!|BEZUG!|NAME\?|NUM!|ZAHL!|NULL!))
  | (?P<ext>(?:'(?:[^']|'')+'|\[[^\]]*\][^\W\d][\w.]*|[^\W\d][\w.]*)!(?:\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d{1,7}:\$?\d{1,7}))
  | (?P<range>\$?[A-Z]{1,3}\$?\d{1,7}:\$?[A-Z]{1,3}\$?\d{1,7})
  | (?P<whole>\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d{1,7}:\$?\d{1,7})
  | (?P<func>(?:_xlfn\.|_xlws\.)?[A-Z][A-Z0-9.]*(?=\())
  | (?P<ref>\$?[A-Z]{1,3}\$?\d{1,7})(?![A-Za-z0-9_(\[])
  | (?P<num>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?|\.\d+)
  | (?P<bool>TRUE|FALSE)(?![A-Za-z0-9_(])
  | (?P<name>[A-Za-z_\\][A-Za-z0-9_.]*(?:\[[^\]]*(?:\[[^\]]*\][^\]]*)*\])?)
  | (?P<op><=|>=|<>|[-+*/^&=<>%(),;{}@])
    )""",
    re.VERBOSE,
)

_REF_PARTS = re.compile(r"\$?([A-Z]{1,3})\$?(\d{1,7})")
# So groß ist ein Blatt in Excel: Ganze Spalten (A:A) und Zeilen (3:3) reichen bis hierhin.
MAX_ROW, MAX_COL = 1_048_576, 16_384
_BINARY = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5}


def position(text: str) -> Pos:
    match = _REF_PARTS.fullmatch(text)
    return int(match.group(2)), column_number(match.group(1))


def area(text: str) -> tuple[int, int, int, int]:
    """„B2:C9“, „A:C“ (ganze Spalten) oder „3:5“ (ganze Zeilen) als (r1, c1, r2, c2)."""
    a, b = text.replace("$", "").split(":")
    if a.isdigit():
        return min(int(a), int(b)), 1, max(int(a), int(b)), MAX_COL
    if b.isalpha():
        c1, c2 = column_number(a), column_number(b)
        return 1, min(c1, c2), MAX_ROW, max(c1, c2)
    (r1, c1), (r2, c2) = position(a), position(b)
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def _sheet_part(text: str) -> tuple[str | None, str]:
    """„'Daten 2024'!B3“ → („Daten 2024“, „B3“); eine andere Mappe ([1]…) → (None, …)."""
    name, _, ref = text.rpartition("!")
    if name.startswith("["):
        return None, ref
    if name.startswith("'"):
        name = name[1:-1].replace("''", "'")
        if name.startswith("["):
            return None, ref
    return name, ref


def tokens(formula: str) -> list[tuple[str, str]]:
    found, at = [], 0
    text = formula.strip().lstrip("=")
    # Bis zum letzten Zeichen, das kein Leerzeichen ist (ohne den Rest jedes Mal neu anzusehen).
    end = len(text.rstrip())
    while at < end:
        match = _TOKEN.match(text, at)
        if not match or match.end() == at:
            raise Unknown(f"unlesbar bei {text[at:at + 10]!r}")
        kind = match.lastgroup
        found.append((kind, match.group(kind)))
        at = match.end()
    return found


# Zelladressen und Bereiche: Darin unterscheiden sich heruntergezogene Formeln, sonst in nichts.
_SLOTS = ("ref", "range", "whole", "ext")


def _leaf(kind: str, text: str) -> Node:
    """Eine Zelladresse oder ein Bereich als Knoten."""
    if kind == "ref":
        return Node("ref", position(text))
    if kind in ("range", "whole"):
        return Node("range", area(text))
    sheet, ref = _sheet_part(text)
    if sheet is None:
        return Node("unknown", text)  # Eine andere Mappe.
    if ":" in ref:
        return Node("xrange", (sheet, area(ref)))
    return Node("xref", (sheet, position(ref.replace("$", ""))))


class _Parser:
    def __init__(self, formula: str, items: list[tuple[str, str]] | None = None, slots: bool = False) -> None:
        self.items = tokens(formula) if items is None else items
        self.at = 0
        # slots: Adressen werden zu Platzhaltern (Node("slot", i)), aus denen eine Vorlage für dieselbe Formel in
        # anderen Zeilen wird.
        self.slots = slots
        self.count = 0

    def peek(self) -> tuple[str, str] | None:
        return self.items[self.at] if self.at < len(self.items) else None

    def take(self) -> tuple[str, str]:
        item = self.items[self.at]
        self.at += 1
        return item

    def expect(self, op: str) -> None:
        item = self.peek()
        if item != ("op", op):
            raise Unknown(f"„{op}“ erwartet")
        self.at += 1

    def parse(self) -> Node:
        node = self.expression(0)
        if self.peek() is not None:
            raise Unknown("Rest übrig")
        return node

    def expression(self, least: int) -> Node:
        """Rangfolge wie in Excel: Vergleich < & < + - < * / < ^, alles von links."""
        left = self.unary()
        while True:
            item = self.peek()
            if item is None or item[0] != "op" or item[1] not in _BINARY or _BINARY[item[1]] < least:
                return left
            power = _BINARY[item[1]]
            self.take()
            # Excel rechnet auch ^ von links: 2^3^2 = 64.
            right = self.expression(power + 1)
            left = Node("op", item[1], (left, right))

    def unary(self) -> Node:
        item = self.peek()
        if item in (("op", "-"), ("op", "+")):
            self.take()
            inner = self.unary()
            return Node("neg", None, (inner,)) if item[1] == "-" else inner
        if item == ("op", "@"):
            # Excel schreibt „@“ vor Formeln aus älteren Versionen (=@INDEX(…), =@STABW.S(…)): Schnittmenge mit
            # der eigenen Zeile. Vor einer Zelle oder Funktion mit einem Wert ändert das nichts; vor einem Bereich
            # rechnet der Prüfer nicht nach.
            self.take()
            after = self.peek()
            inner = self.unary()
            if after is not None and (after[0] in ("range", "whole") or (after[0] == "ext" and ":" in after[1])):
                return Node("unknown", "@")
            return inner
        node = self.atom()
        while self.peek() == ("op", "%"):
            self.take()
            node = Node("pct", None, (node,))
        return node

    def atom(self) -> Node:
        if self.peek() is None:
            raise Unknown("Formel endet zu früh")
        kind, text = self.take()
        if kind == "num":
            return Node("num", float(text))
        if kind == "str":
            return Node("str", text[1:-1].replace('""', '"'))
        if kind == "bool":
            return Node("bool", text == "TRUE")
        if kind == "err":
            return Node("err", text)
        if kind in _SLOTS:
            if self.slots:
                self.count += 1
                return Node("slot", self.count - 1)
            return _leaf(kind, text)
        if kind == "name":
            return Node("unknown", text)
        if kind == "func":
            name = re.sub(r"^_xl(?:fn|ws)\.", "", text).upper()
            self.expect("(")
            args: list[Node] = []
            if self.peek() != ("op", ")"):
                while True:
                    if self.peek() in (("op", ","), ("op", ";"), ("op", ")")):
                        args.append(Node("empty"))
                    else:
                        args.append(self.expression(0))
                    if self.peek() in (("op", ","), ("op", ";")):
                        self.take()
                        continue
                    break
            self.expect(")")
            return Node("call", name, tuple(args))
        if (kind, text) == ("op", "("):
            node = self.expression(0)
            self.expect(")")
            return node
        raise Unknown(f"unerwartet: {text}")


# Während einer Prüfung liest parse() jede Formel nur einmal: Dieselbe Formel lesen viele Prüfungen, und bei großen
# Blättern ist das Lesen das Teuerste. Je Thread, damit zwei Prüfungen zugleich sich nicht stören.
_parsed = threading.local()


@contextmanager
def remembered():
    """Solange das läuft, merkt sich parse() die gelesenen Formeln (auch die unlesbaren)."""
    if getattr(_parsed, "nodes", None) is not None:
        yield
        return
    _parsed.nodes = {}
    try:
        yield
    finally:
        _parsed.nodes = None


def remember(kind: str, key, compute):
    """Während remembered(): einmal ausrechnen, dann merken (wie parse). Sonst jedes Mal neu."""
    nodes = getattr(_parsed, "nodes", None)
    if nodes is None:
        return compute()
    store = nodes.get(kind)
    if store is None:
        store = nodes[kind] = {}
    try:
        return store[key]
    except KeyError:
        value = store[key] = compute()
        return value


def parse(formula: str) -> Node:
    nodes = getattr(_parsed, "nodes", None)
    if nodes is None:
        return _Parser(formula).parse()
    found = nodes.get(formula)
    if found is None:
        try:
            found = _from_template(formula, nodes)
        except Unknown as err:
            found = err
        nodes[formula] = found
    if isinstance(found, Unknown):
        raise Unknown(*found.args)
    return found


def _from_template(formula: str, nodes: dict) -> Node:
    """Heruntergezogene Formeln (=H2/(I2/100)^2, =H3/(I3/100)^2 …) haben denselben Aufbau: einmal lesen, für die
    übrigen nur die Adressen einsetzen."""
    items = tokens(formula)
    try:
        leaves = [_leaf(kind, text) for kind, text in items if kind in _SLOTS]
    except (AttributeError, ValueError):
        return _Parser(formula, items).parse()
    leaf = iter(leaves)
    key = tuple(next(leaf).kind if kind in _SLOTS else (kind, text) for kind, text in items)
    templates = nodes.get("\x00templates")
    if templates is None:
        templates = nodes["\x00templates"] = {}
    template = templates.get(key)
    if template is None:
        try:
            template = _Parser(formula, items, slots=True).parse()
        except Unknown as err:
            template = err
        templates[key] = template
    if isinstance(template, Unknown):
        raise Unknown(*template.args)
    return _fill(template, leaves)


def _fill(node: Node, leaves: list[Node]) -> Node:
    if node.kind == "slot":
        return leaves[node.value]
    if not node.args:
        return node
    args = tuple(_fill(a, leaves) for a in node.args)
    return Node(node.kind, node.value, args)


def references(node: Node, own: str | None = None) -> tuple[set[Pos], list[tuple[int, int, int, int]], bool]:
    """Die Zellen und Bereiche dieses Blattes, die eine Formel liest, und ob sie etwas Unbekanntes liest.

    own: der Name dieses Blattes. „Daten!B3“ auf dem Blatt „Daten“ zählt dann wie „B3“."""
    cells: set[Pos] = set()
    ranges: list[tuple[int, int, int, int]] = []
    unknown = False
    mine = own.casefold() if own else None

    def walk(n: Node) -> None:
        nonlocal unknown
        if n.kind == "ref" or (n.kind == "xref" and n.value[0].casefold() == mine):
            cells.add(n.value if n.kind == "ref" else n.value[1])
        elif n.kind == "range" or (n.kind == "xrange" and n.value[0].casefold() == mine):
            ranges.append(n.value if n.kind == "range" else n.value[1])
        elif n.kind == "unknown":
            unknown = True
        for child in n.args:
            walk(child)

    walk(node)
    return cells, ranges, unknown


def reach(node: Node, own: str | None = None) -> dict[str | None, int]:
    """Bis zu welcher Zeile eine Formel je Blatt liest (None: dieses Blatt, sonst klein geschrieben)."""
    mine = own.casefold() if own else None
    found: dict[str | None, int] = {}

    def walk(n: Node) -> None:
        if n.kind in ("ref", "range"):
            sheet, last = None, n.value[0] if n.kind == "ref" else n.value[2]
        elif n.kind in ("xref", "xrange"):
            sheet = n.value[0].casefold()
            sheet = None if sheet == mine else sheet
            last = n.value[1][0] if n.kind == "xref" else n.value[1][2]
        else:
            sheet = last = None
        if last is not None and last > found.get(sheet, 0):
            found[sheet] = last
        for child in n.args:
            walk(child)

    walk(node)
    return found


def foreign(node: Node, own: str | None = None) -> set[str]:
    """Die anderen Blätter (klein geschrieben), die eine Formel liest."""
    mine = own.casefold() if own else None
    found: set[str] = set()

    def walk(n: Node) -> None:
        if n.kind in ("xref", "xrange") and n.value[0].casefold() != mine:
            found.add(n.value[0].casefold())
        for child in n.args:
            walk(child)

    walk(node)
    return found


_FINANCE = {"NPV", "XNPV", "IRR", "PMT", "PV", "FV", "IPMT", "PPMT", "NPER"}


def _pmt(rate: float, nper: float, pv: float, fv: float, kind: float) -> float:
    if nper == 0:
        raise _Fail(Err("#NUM!"))
    if rate == 0:
        return -(pv + fv) / nper
    grow = (1 + rate) ** nper
    return -(rate * (pv * grow + fv)) / ((1 + rate * kind) * (grow - 1))


def _fv(rate: float, nper: float, pmt: float, pv: float, kind: float) -> float:
    if rate == 0:
        return -(pv + pmt * nper)
    grow = (1 + rate) ** nper
    return -(pv * grow + pmt * (1 + rate * kind) * (grow - 1) / rate)


def _irr(values: list[float], guess: float) -> float:
    """Der interne Zinsfuß wie IKV: Newton vom Schätzwert aus, sonst Halbieren zwischen -99 % und 1000 %."""
    if not any(v > 0 for v in values) or not any(v < 0 for v in values):
        raise _Fail(Err("#NUM!"))

    def npv(r: float) -> float:
        return math.fsum(v / (1 + r) ** i for i, v in enumerate(values))

    r = guess
    for _ in range(50):
        f = npv(r)
        d = math.fsum(-i * v / (1 + r) ** (i + 1) for i, v in enumerate(values))
        if d == 0:
            break
        step = f / d
        r -= step
        if r <= -1:
            break
        if abs(step) < 1e-12:
            return r
    low, high = -0.99, 10.0
    if npv(low) * npv(high) > 0:
        raise _Fail(Err("#NUM!"))
    for _ in range(200):
        mid = (low + high) / 2
        if npv(low) * npv(mid) <= 0:
            high = mid
        else:
            low = mid
    return (low + high) / 2


def _has_area(node: Node) -> bool:
    if node.kind in ("range", "xrange"):
        return True
    return node.kind in ("op", "neg", "pct") and any(_has_area(a) for a in node.args)


def calls(node: Node) -> list[Node]:
    """Alle Funktionsaufrufe in einer Formel, außen zuerst."""
    found = [node] if node.kind == "call" else []
    for child in node.args:
        found += calls(child)
    return found


# Argumente, die keine Größe sind, sondern sagen, wie gerechnet wird: die Spalte bei SVERWEIS, die
# Stellen bei RUNDEN, die 9 bei TEILERGEBNIS. Eine 2 dort ist keine feste Zahl, die in eine Zelle gehört.
_SETTINGS = {
    "VLOOKUP": {2, 3}, "HLOOKUP": {2, 3}, "MATCH": {2}, "XMATCH": {2, 3}, "XLOOKUP": {4, 5}, "INDEX": {1, 2, 3},
    "ROUND": {1}, "ROUNDUP": {1}, "ROUNDDOWN": {1}, "TRUNC": {1}, "FIXED": {1, 2}, "TEXT": {1},
    "SUBTOTAL": {0}, "AGGREGATE": {0, 1}, "CHOOSE": {0}, "OFFSET": {1, 2, 3, 4},
    "LEFT": {1}, "RIGHT": {1}, "MID": {1, 2}, "LARGE": {1}, "SMALL": {1}, "RANK": {2}, "RANK.EQ": {2}, "RANK.AVG": {2},
    "WEEKDAY": {1}, "WEEKNUM": {1}, "EOMONTH": {1}, "EDATE": {1}, "DATEDIF": {2}, "YEARFRAC": {2}, "ADDRESS": {2, 3, 4},
    # Seiten und Typ bei T.TEST, das Alpha bei KONFIDENZ und T.INV.2S: wie getestet wird, keine Größe aus den Daten.
    "T.TEST": {2, 3}, "TTEST": {2, 3}, "CONFIDENCE": {0}, "CONFIDENCE.NORM": {0}, "CONFIDENCE.T": {0},
    "T.INV.2T": {0}, "TINV": {0}, "T.INV": {0}, "NORM.S.INV": {0}, "NORMSINV": {0}, "T.DIST": {2}, "TDIST": {2},
    "NORM.S.DIST": {1}, "PERCENTILE": {1}, "PERCENTILE.INC": {1}, "PERCENTILE.EXC": {1}, "QUARTILE": {1},
    "QUARTILE.INC": {1}, "QUARTILE.EXC": {1},
    # Monat und Tag in DATUM(2026;4;0), die Teile einer Uhrzeit: ein Kalenderdatum, keine Annahme. Das Jahr zählt.
    "DATE": {1, 2}, "TIME": {0, 1, 2},
}


def literals(node: Node, exponents: bool = True) -> list[float]:
    """Die Zahlen, die direkt in der Formel stehen (ohne Einstellungen wie die Spalte bei SVERWEIS).
    exponents=False: ohne Hochzahlen (die 2 in ^2)."""
    found: list[float] = []

    def walk(n: Node) -> None:
        if n.kind == "num":
            found.append(n.value)
        skip = _SETTINGS.get(n.value, ()) if n.kind == "call" else ()
        if not exponents and n.kind == "op" and n.value == "^" and n.args[1].kind in ("num", "neg"):
            skip = (1,)
        for i, child in enumerate(n.args):
            if i not in skip:
                walk(child)

    walk(node)
    return found


def depth(node: Node, name: str = "IF") -> int:
    """Wie tief Funktionen dieses Namens ineinander stecken."""
    inner = max((depth(a, name) for a in node.args), default=0)
    return inner + (1 if node.kind == "call" and node.value == name else 0)


# --- Rechnen ---


_REFERENCES = ("OFFSET", "INDIRECT")


def dynamic(node: Node) -> bool:
    """Liest die Formel über BEREICH.VERSCHIEBEN oder INDIREKT Zellen, die erst beim Rechnen feststehen?"""
    todo = [node]
    while todo:
        n = todo.pop()
        if n.kind == "call" and n.value in _REFERENCES:
            return True
        todo.extend(n.args)
    return False


def _reference_node(sheet: str | None, top: int, left: int, bottom: int, right: int) -> Node:
    if (top, left) == (bottom, right):
        return Node("ref", (top, left)) if sheet is None else Node("xref", (sheet, (top, left)))
    area = (top, left, bottom, right)
    return Node("range", area) if sheet is None else Node("xrange", (sheet, area))


def _upper_outside_quotes(text: str) -> str:
    """„daten!b5“ wie „DATEN!B5“: Excel liest Bezüge in jeder Schreibweise, Blattnamen in Hochkommas bleiben."""
    parts = re.split(r"('(?:[^']|'')*')", text)
    return "".join(part if part.startswith("'") else part.upper() for part in parts)


def _num(value) -> float:
    if isinstance(value, Err):
        raise _Fail(value)
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, float):
        return value
    if isinstance(value, str):
        number = _as_float(value)
        if number is None:
            raise _Fail(Err("#VALUE!"))
        return number
    raise Unknown("kein Wert")


_DATE_TEXT = re.compile(r"(\d{1,2})\.(\d{1,2})\.(\d{4})(?: (\d{1,2}):(\d{2})(?::(\d{2}))?)?")
_TIME_TEXT = re.compile(r"(\d{1,2}):(\d{2})(?::(\d{2}))?")


def serial_of(text: str) -> float | None:
    """„01.10.2024“, „01.10.2024 08:30“ oder „08:30“ (so zeigt der Prüfer Datumszellen) als Excel-Zahl."""
    text = text.strip()
    match = _DATE_TEXT.fullmatch(text)
    if match:
        try:
            days = (date(int(match[3]), int(match[2]), int(match[1])) - date(1899, 12, 30)).days
        except ValueError:
            return None
        seconds = int(match[4]) * 3600 + int(match[5]) * 60 + int(match[6] or 0) if match[4] else 0
        return days + seconds / 86400
    match = _TIME_TEXT.fullmatch(text)
    if match and int(match[1]) < 24 and int(match[2]) < 60 and int(match[3] or 0) < 60:
        return (int(match[1]) * 3600 + int(match[2]) * 60 + int(match[3] or 0)) / 86400
    return None


def _as_float(text: str) -> float | None:
    """Text, mit dem Excel rechnen kann: „12“, „3,5“ oder ein Datum."""
    try:
        return float(text.strip().replace(",", "."))
    except ValueError:
        return serial_of(text)


def _numbers(values) -> list[float]:
    """Zahlen aus Bereichen wie SUMME sie zählt: Text und Leeres nicht."""
    out = []
    for v in values:
        if isinstance(v, Err):
            raise _Fail(v)
        if isinstance(v, bool):
            continue
        if isinstance(v, float):
            out.append(v)
    return out


class _Fail(Exception):
    def __init__(self, err: Err) -> None:
        self.err = err


def _round(x: float, digits: float, how: str) -> float:
    factor = 10 ** int(digits)
    y = abs(x) * factor
    if how == "half":
        y = math.floor(y + 0.5 + 1e-9)
    elif how == "up":
        y = math.ceil(y - 1e-9)
    else:
        y = math.floor(y + 1e-9)
    return math.copysign(y / factor, x)


def _criterion(test):
    """„>5“, „RN“, „R*“, 3, "": wie ZÄHLENWENN vergleicht."""
    if isinstance(test, bool):
        return lambda v: isinstance(v, bool) and v == test
    if isinstance(test, float):
        return lambda v: isinstance(v, float) and not isinstance(v, bool) and v == test
    text = "" if test is None else str(test)
    match = re.match(r"^(<=|>=|<>|<|>|=)?(.*)$", text, re.S)
    op, rest = match.group(1) or "=", match.group(2)
    if rest == "" and op in ("=", "<>"):
        # "" und "=" finden leere Zellen, "<>" alle nicht leeren.
        empty = lambda v: v is None or v == ""  # noqa: E731
        return empty if op == "=" else (lambda v: not empty(v))
    try:
        number = float(rest.replace(",", "."))
    except ValueError:
        number = None
    if number is not None:
        compare = {"=": lambda v: v == number, "<>": lambda v: v != number, "<": lambda v: v < number,
                   ">": lambda v: v > number, "<=": lambda v: v <= number, ">=": lambda v: v >= number}[op]
        return lambda v: isinstance(v, float) and compare(v)
    if op not in ("=", "<>"):
        raise Unknown("Textvergleich mit < >")
    if re.search(r"(?<!~)[*?]", rest):
        # Platzhalter: * beliebig viel, ? ein Zeichen, ~ davor heißt das Zeichen selbst.
        pattern = "".join(
            re.escape(part[1]) if part.startswith("~") and len(part) == 2 else ".*" if part == "*" else "." if part == "?" else re.escape(part)
            for part in re.findall(r"~[*?~]|[*?]|[^*?~]+|~", rest)
        )
        wild = re.compile(pattern, re.I | re.S)
        same = lambda v: isinstance(v, str) and wild.fullmatch(v) is not None  # noqa: E731
    else:
        plain = re.sub(r"~([*?~])", r"\1", rest).lower()
        same = lambda v: isinstance(v, str) and v.lower() == plain  # noqa: E731
    return same if op == "=" else (lambda v: not same(v))


def _equal_key(test):
    """Wonach ZÄHLENWENN und Verwandte mit einfachem Gleich suchen („Nord“, 5, WAHR), als Schlüssel für ein Verzeichnis
    der Werte (_value_key). None, wenn die Bedingung mehr kann (>5, Platzhalter, leer): dann Zelle für Zelle."""
    if isinstance(test, bool):
        return ("b", test)
    if isinstance(test, float):
        return ("n", test)
    if not isinstance(test, str):
        return None
    text = test[1:] if test.startswith("=") else test
    if not text or text[0] in "<>" or re.search(r"[*?~]", text):
        return None
    try:
        return ("n", float(text.replace(",", ".")))
    except ValueError:
        return ("s", text.lower())


def _value_key(value):
    """Der Schlüssel eines Werts, so wie _criterion ihn bei einfachem Gleich vergleicht."""
    if isinstance(value, bool):
        return ("b", value)
    if isinstance(value, float):
        return ("n", value)
    if isinstance(value, str):
        return ("s", value.lower())
    return None


# Ab so vielen Zeilen zählt ZÄHLENWENN mit einfachem Gleich über ein Verzeichnis je Spalte statt Zelle für Zelle.
MATCH_INDEX = 32
# Jede endliche Zahl ist ein ganzzahliges Vielfaches von 2^-1074. Unter 2^1000 (Excel geht nur bis 1e308) läuft
# auch die Summe von 200 000 Zellen nicht über.
_EXACT = 1074
_HUGE = 2.0 ** 1000


class Evaluator:
    """Rechnet Formeln mit den Werten aus value(pos). Bereiche werden flach gelesen.

    other(blatt, pos) liefert Werte anderer Blätter (sonst gilt „Daten!B3“ als unbekannt),
    extent(blatt) wie weit ein Blatt reicht (letzte Zeile, letzte Spalte), damit A:A nicht eine Million Zellen liest.
    """

    def __init__(self, value, constants: dict[float, float] | None = None, other=None, extent=None, static: bool = False,
                 shared=None) -> None:
        self.value = value
        # shared(blatt): ein anderes Blatt, dessen Werte sich nie ändern, mit gemerkten Bereichen (grids) und
        # Suchverzeichnissen (found) für alle Rechnungen darüber, oder None.
        self.shared = shared
        # Feste Zahlen in Formeln ersetzen (für „was wäre, wenn 577 ein Prozent größer wäre?“).
        self.constants = constants or {}
        self.other = other
        self.extent = extent
        # static: Die Werte ändern sich nicht (Nachrechnen mit den Werten aus Excel). Dann merkt er sich Bereiche
        # und Suchspalten, sonst liest jeder SVERWEIS in 2000 Zeilen die ganze Tabelle neu.
        self._grids: dict | None = {} if static else None
        self._found: dict = {}
        # Je Bereich (oder Bereiche nebeneinander): welche Zeilen welchen Wert haben, für ZÄHLENWENN und Verwandte mit
        # einfachem Gleich. Innerhalb einer Rechnung ändern sich die Werte nicht, auch nicht die der Formeln.
        self._rows_by_key: dict = {}
        self._read: dict = {}
        self._sums: dict = {}
        self._conditions: dict = {}
        # Für die Herleitung: welche Zelle SVERWEIS, XVERWEIS und INDEX geliefert haben, als (Funktion, Blatt, Zelle).
        self.hits: list[tuple[str, str | None, Pos]] | None = None
        # nested(pos): Steht in der Zelle dieses Blatts selbst ein TEILERGEBNIS (oder AGGREGAT)? Die lässt
        # TEILERGEBNIS weg, damit die Gesamtsumme die Zwischensummen nicht doppelt zählt. None: unbekannt.
        self.nested = None
        # Ausgeblendete Zeilen dieses Blatts (hidden, aus welchem Grund auch immer) und die davon, die ein Filter
        # ausblendet (filtered). TEILERGEBNIS(109; …) lässt die einen weg, TEILERGEBNIS(9; …) nur die anderen.
        # None: unbekannt, dann zählt alles.
        self.hidden: frozenset[int] | None = None
        self.filtered: frozenset[int] | None = None

    def run(self, node: Node):
        try:
            return self.eval(node)
        except _Fail as fail:
            return fail.err

    # --- Bereiche ---

    def _get(self, sheet: str | None, pos: Pos):
        if sheet is None:
            return self.value(pos)
        if self.other is None:
            raise Unknown(f"anderes Blatt {sheet}")
        return self.other(sheet, pos)

    def _bounds(self, sheet: str | None, area) -> tuple[int, int, int, int]:
        r1, c1, r2, c2 = area
        if (r2 >= MAX_ROW or c2 >= MAX_COL) and self.extent is not None:
            rows, cols = self.extent(sheet)
            r2, c2 = min(r2, max(rows, r1)), min(c2, max(cols, c1))
        if (r2 - r1 + 1) * (c2 - c1 + 1) > 200_000:
            raise Unknown("Bereich zu groß")
        return r1, c1, r2, c2

    def area_of(self, node: Node) -> tuple[str | None, tuple[int, int, int, int]]:
        if node.kind == "range":
            return None, self._bounds(None, node.value)
        if node.kind == "xrange":
            return node.value[0], self._bounds(node.value[0], node.value[1])
        if node.kind == "ref":
            r, c = node.value
            return None, (r, c, r, c)
        if node.kind == "xref":
            r, c = node.value[1]
            return node.value[0], (r, c, r, c)
        raise Unknown("kein Bereich")

    def area_key(self, node: Node) -> tuple:
        sheet, area = self.area_of(node)
        return (sheet, *area)

    def _fixed(self, sheet: str | None):
        """Wo Bereiche eines Blatts gemerkt werden, deren Werte sich nicht ändern, oder None."""
        if sheet is not None and self.shared is not None:
            found = self.shared(sheet)
            if found is not None:
                return found
        return self if self._grids is not None else None

    @property
    def grids(self) -> dict | None:
        return self._grids

    @property
    def found(self) -> dict:
        return self._found

    def grid(self, node: Node) -> list[list]:
        sheet, (r1, c1, r2, c2) = self.area_of(node)
        key = (sheet, r1, c1, r2, c2)
        # Eine einzelne Zelle ist schnell gelesen: nur Bereiche merken.
        fixed = self._fixed(sheet) if r2 > r1 or c2 > c1 else None
        if fixed is not None and key in fixed.grids:
            return fixed.grids[key]
        rows = [[self._get(sheet, (r, c)) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]
        if fixed is not None:
            fixed.grids[key] = rows
        return rows

    def _exact(self, key, line: list, where: tuple) -> int:
        """Die erste genaue Fundstelle, bei unveränderlichen Werten über ein gemerktes Verzeichnis."""
        fixed = self._fixed(where[-1][0])
        if fixed is None or (isinstance(key, str) and re.search(r"(?<!~)[*?]", key)):
            return _find(key, line, 0)
        index = fixed.found.get(where)
        if index is None:
            index = {}
            for i, v in enumerate(line):
                index.setdefault(_norm(v), i)
            fixed.found[where] = index
        at = index.get(_norm(key))
        if at is None or _norm(key) is None:
            raise _Fail(Err("#N/A"))
        return at

    def _column(self, sheet: str | None, c: int, r1: int, r2: int) -> dict[int, object]:
        """Die Werte einer Spalte, mindestens von r1 bis r2, jede Zelle einmal je Rechnung gelesen. Mehr braucht es
        nicht: Eine Formel, die einmal gelesen ist, hat ihren Wert (value merkt ihn sich), und nur eine noch nicht
        fertige Formel kann einen Kreisbezug melden."""
        found = self._read.get((sheet, c))
        if found is None or r2 < found[0] - 1 or r1 > found[1] + 1:
            # Gelesen wird immer ein Stück ohne Lücke; liegt der Bereich anderswo, fängt es dort neu an.
            found = self._read[(sheet, c)] = [r1, r1 - 1, {}]
        start, end, values = found
        for r in range(r1, start):
            values[r] = self._get(sheet, (r, c))
        found[0] = min(start, r1)
        for r in range(end + 1, r2 + 1):
            values[r] = self._get(sheet, (r, c))
            found[1] = r
        return values

    def _matching(self, areas: list[Node], wanted: list) -> list[int] | None:
        """Die Stellen im Bereich (von 0 an, Zeile für Zeile), an denen jeder Bereich seinen gesuchten Wert hat, wenn
        alle Bedingungen ein einfaches Gleich sind und jeder Bereich eine Spalte über dieselben Zeilen ist; sonst None.
        Das Verzeichnis je Spalte entsteht einmal je Rechnung und wächst mit: Eine Duplikat-Spalte
        ZÄHLENWENN($A$2:$A$20001;A2) oder ZÄHLENWENN($A$2:A2;A2) liest Spalte A so einmal statt 20 000-mal."""
        keys = tuple(_equal_key(w) for w in wanted)
        if any(k is None for k in keys):
            return None
        spans = [self.area_of(a) for a in areas]
        sheet, (r1, _, r2, _) = spans[0]
        if any(c1 != c2 or (s, a1, a2) != (sheet, r1, r2) for s, (a1, c1, a2, c2) in spans) or r2 - r1 < MATCH_INDEX:
            return None
        columns = tuple(c1 for _, (_, c1, _, _) in spans)
        read = [self._column(sheet, c, r1, r2) for c in columns]
        found = self._rows_by_key.get((sheet, columns))
        if found is None or r1 < found[0] or r1 > found[1] + 1:
            found = self._rows_by_key[(sheet, columns)] = [r1, r1 - 1, {}]
        start, end, index = found
        for r in range(end + 1, r2 + 1):
            row = tuple(_value_key(values[r]) for values in read)
            if None not in row:
                index.setdefault(row, []).append(r)
            found[1] = r
        rows = index.get(keys, [])
        return [r - r1 for r in rows[bisect_left(rows, r1):bisect_right(rows, r2)]]

    def _summed(self, node: Node) -> tuple[float, int] | None:
        """SUMME und Anzahl der Zahlen eines Bereichs, wenn er eine lange Spalte ist; sonst None. Je Spalte zählt es
        einmal von oben zusammen: Eine Laufsumme SUMME($E$2:E20001) liest Spalte E so einmal statt 20 000-mal. Genau
        wie math.fsum, weil es ganz genau zählt (in Vielfachen der kleinsten Zahl) und erst am Ende rundet."""
        sheet, (r1, c1, r2, c2) = self.area_of(node)
        if c1 != c2 or r2 - r1 < MATCH_INDEX:
            return None
        values = self._column(sheet, c1, r1, r2)
        found = self._sums.get((sheet, c1))
        if found is None or r1 < found[0] or r1 > found[0] + len(found[1]) - 1:
            found = self._sums[(sheet, c1)] = [r1, [0], [0], [0], []]
        start, sums, counts, zeros, errors = found
        for r in range(start + len(sums) - 1, r2 + 1):
            v = values[r]
            number = isinstance(v, float) and not isinstance(v, bool)
            if number and not abs(v) < _HUGE:
                # Unendlich oder so groß, dass math.fsum unterwegs überläuft: dann Zelle für Zelle.
                del self._sums[(sheet, c1)]
                return None
            if isinstance(v, Err):
                errors.append(r)
            if number:
                n, d = v.as_integer_ratio()
                sums.append(sums[-1] + (n << (_EXACT - d.bit_length() + 1)))
                counts.append(counts[-1] + 1)
                zeros.append(zeros[-1] + (v == 0 and math.copysign(1.0, v) < 0))
            else:
                sums.append(sums[-1])
                counts.append(counts[-1])
                zeros.append(zeros[-1])
        at = bisect_left(errors, r1)
        if at < len(errors) and errors[at] <= r2:
            raise _Fail(values[errors[at]])
        a, b = r1 - start, r2 - start + 1
        total, count = sums[b] - sums[a], counts[b] - counts[a]
        if not total:
            # Null mit dem Vorzeichen, das math.fsum gibt (nur -0 in allen Zellen).
            return math.fsum([-0.0] if count and zeros[b] - zeros[a] == count else []), count
        return total / (1 << _EXACT), count

    def _conditional(self, name: str, areas: list[Node], wanted: list, source: Node | None):
        """ZÄHLENWENN, SUMMEWENN, MITTELWERTWENN und die mit S über das Verzeichnis je Spalte (_matching); None, wo das
        nicht geht. Je Rechnung gemerkt: SUMMEWENN($C$2:$C$20001;C2;$E$2:$E$20001) in jeder Zeile fragt für jede
        Station Tausende Male dasselbe."""
        keys = tuple(_equal_key(w) for w in wanted)
        if any(k is None for k in keys):
            return None
        key = (name, tuple(self.area_key(a) for a in areas), keys, None if source is None else self.area_key(source))
        if key not in self._conditions:
            rows = self._matching(areas, wanted)
            if rows is None:
                return None
            found = [None] * len(rows) if source is None else self._at(source, rows, self._size(areas[0]))
            if found is None:
                return None
            picked = [v for v in found if isinstance(v, float) and not isinstance(v, bool)]
            if name in ("COUNTIF", "COUNTIFS"):
                got = float(len(rows))
            elif name in ("SUMIF", "SUMIFS"):
                got = math.fsum(picked)
            elif name in ("AVERAGEIF", "AVERAGEIFS"):
                got = math.fsum(picked) / len(picked) if picked else Err("#DIV/0!")
            else:
                got = ((min if name == "MINIFS" else max)(picked)) if picked else 0.0
            self._conditions[key] = got
        got = self._conditions[key]
        if isinstance(got, Err):
            raise _Fail(got)
        return got

    def _at(self, node: Node, offsets: list[int], size: int) -> list | None:
        """Die Werte eines Bereichs an diesen Stellen, wenn er eine Spalte mit size Zellen ist; sonst None."""
        sheet, (r1, c1, r2, c2) = self.area_of(node)
        if c1 != c2 or r2 - r1 + 1 != size:
            return None
        values = self._column(sheet, c1, r1, r2)
        return [values[r1 + i] for i in offsets]

    def _size(self, node: Node) -> int:
        _, (r1, c1, r2, c2) = self.area_of(node)
        return (r2 - r1 + 1) * (c2 - c1 + 1)

    def cells(self, area, sheet: str | None = None) -> list:
        r1, c1, r2, c2 = self._bounds(sheet, area)
        return [self._get(sheet, (r, c)) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]

    def values_of(self, node: Node) -> list:
        """Die Werte eines Bereichs (auch auf einem anderen Blatt), Zeile für Zeile."""
        return [v for row in self.grid(node) for v in row]

    def _filter(self, args: tuple) -> list:
        """FILTER(Bereich; Bedingung) in MEDIAN, STABW und Co.: die Werte, deren Bedingung WAHR ist. Leere Zellen
        gibt FILTER wie Excel als 0 zurück."""
        if len(args) < 2 or not self.is_area(args[0]):
            raise Unknown("FILTER ohne Bereich")
        values = self.values_of(args[0])
        keep = self._mask(args[1], len(values))
        picked = [0.0 if v is None else v for v, k in zip(values, keep, strict=True) if k]
        if not picked:
            if len(args) > 2 and args[2].kind != "empty":
                return [self.eval(args[2])]
            raise _Fail(Err("#CALC!"))
        return picked

    def _mask(self, node: Node, size: int) -> list[bool]:
        """Eine Bedingung über einen Bereich, Zelle für Zelle: (C2:C9<>999)*(C2:C9<>"")."""
        if node.kind == "op" and node.value in ("*", "+"):
            a, b = self._mask(node.args[0], size), self._mask(node.args[1], size)
            return [x and y if node.value == "*" else x or y for x, y in zip(a, b, strict=True)]
        if node.kind == "op" and node.value in ("=", "<>", "<", ">", "<=", ">=") and self.is_area(node.args[0]):
            values = self.values_of(node.args[0])
            if len(values) != size:
                raise _Fail(Err("#VALUE!"))
            other = self.eval(node.args[1])
            out = []
            for v in values:
                try:
                    out.append(bool(self.operator(node.value, v, other)))
                except Unknown:
                    # Text und Zahl: nie gleich, aber größer oder kleiner sagen wir nicht.
                    if node.value not in ("=", "<>"):
                        raise
                    out.append(node.value == "<>")
            return out
        raise Unknown("FILTER-Bedingung")

    @staticmethod
    def is_area(node: Node) -> bool:
        return node.kind in ("range", "xrange")

    def flat(self, nodes) -> list:
        out = []
        for n in nodes:
            if n.kind in ("range", "xrange"):
                out += self.values_of(n)
            elif n.kind == "call" and n.value == "FILTER":
                out += self._filter(n.args)
            elif n.kind in ("ref", "xref"):
                out.append(self.eval(n))
            elif n.kind == "empty":
                continue
            else:
                v = self.eval(n)
                # Direkt eingetippte Zahlen und Wahrheitswerte zählen bei SUMME mit.
                out.append(float(v) if isinstance(v, bool) else v)
        return out

    def eval(self, n: Node):
        kind = n.kind
        if kind == "num":
            return self.constants.get(n.value, n.value)
        if kind in ("str", "bool"):
            return n.value
        if kind == "err":
            raise _Fail(Err(n.value))
        if kind == "ref":
            return self.value(n.value)
        if kind == "xref":
            return self._get(n.value[0], n.value[1])
        if kind in ("range", "xrange"):
            raise Unknown("Bereich ohne Funktion")
        if kind == "unknown":
            raise Unknown(str(n.value))
        if kind == "empty":
            return None
        if kind == "neg":
            return -_num(self.eval(n.args[0]))
        if kind == "pct":
            return _num(self.eval(n.args[0])) / 100
        if kind == "op":
            return self.operator(n.value, self.eval(n.args[0]), self.eval(n.args[1]))
        if kind == "call":
            return self.call(n.value, n.args)
        raise Unknown(kind)

    def _subtotal(self, args: tuple):
        """TEILERGEBNIS(9; C2:C17): wie SUMME (1 Mittelwert, 2 Anzahl … 11 Varianz), aber ohne die Zellen im Bereich,
        in denen selbst ein TEILERGEBNIS steht, und ohne weggefilterte Zeilen; 101 bis 111 lassen auch von Hand
        ausgeblendete weg. Kennt der Prüfer die ausgeblendeten Zeilen nicht, rechnet er, als seien alle sichtbar
        (analyse.py vergleicht darum nicht mit Excel, wenn es anders herauskommt)."""
        if len(args) < 2:
            raise _Fail(Err("#VALUE!"))
        number = _num(self.eval(args[0]))
        base = int(number) % 100 if number == int(number) and int(number) in (*range(1, 12), *range(101, 112)) else 0
        if not base:
            raise _Fail(Err("#VALUE!"))
        if self.nested is None:
            raise Unknown("TEILERGEBNIS ohne Wissen, wo andere stehen")
        away = (self.hidden if number > 100 else self.filtered) or frozenset()
        values = []
        for node in args[1:]:
            if node.kind == "range":
                r1, c1, r2, c2 = self._bounds(None, node.value)
                cells = [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
            elif node.kind == "ref":
                cells = [node.value]
            else:
                raise Unknown("TEILERGEBNIS über ein anderes Blatt")
            values += [self.value(pos) for pos in cells if pos[0] not in away and not self.nested(pos)]
        if base == 3:
            for v in values:
                if isinstance(v, Err):
                    raise _Fail(v)
            return float(sum(1 for v in values if v is not None and v != ""))
        numbers = _numbers(values)
        if base == 9:
            return math.fsum(numbers)
        if base == 2:
            return float(len(numbers))
        if base == 6:
            return math.prod(numbers) if numbers else 0.0
        if base in (4, 5):
            return (max if base == 4 else min)(numbers) if numbers else 0.0
        if not numbers or len(numbers) < 2 and base in (7, 10):
            raise _Fail(Err("#DIV/0!"))
        if base == 1:
            return math.fsum(numbers) / len(numbers)
        if base == 7:
            return statistics.stdev(numbers)
        if base == 8:
            return statistics.pstdev(numbers)
        if base == 10:
            return statistics.variance(numbers)
        return statistics.pvariance(numbers)

    def _finance(self, name: str, args: tuple):
        """Barwert, Rate, Zins und Tilgung wie in Excel (Zahlungen am Ende der Periode; am Anfang nur, wo die Formel
        einfach ist)."""
        if name == "NPV":
            rate = _num(self.eval(args[0]))
            values = [v for v in self.flat(args[1:]) if isinstance(v, float) and not isinstance(v, bool)]
            if rate == -1:
                raise _Fail(Err("#DIV/0!"))
            return math.fsum(v / (1 + rate) ** (i + 1) for i, v in enumerate(values))
        if name == "XNPV":
            rate = _num(self.eval(args[0]))
            values, dates = _numbers(self.flat(args[1:2])), _numbers(self.flat(args[2:3]))
            if len(values) != len(dates) or not values:
                raise _Fail(Err("#NUM!"))
            return math.fsum(v / (1 + rate) ** ((d - dates[0]) / 365) for v, d in zip(values, dates))
        if name == "IRR":
            values = _numbers(self.flat(args[:1]))
            guess = _num(self.arg(args, 1, 0.1))
            return _irr(values, guess)
        numbers = [_num(self.arg(args, i, 0.0)) for i in range(len(args))]
        if name == "PMT":
            rate, nper, pv, fv, kind = (numbers + [0.0, 0.0])[:5]
            return _pmt(rate, nper, pv, fv, kind)
        if name == "PV":
            rate, nper, pmt, fv, kind = (numbers + [0.0, 0.0])[:5]
            if rate == 0:
                return -(pmt * nper + fv)
            return -(pmt * (1 + rate * kind) * ((1 + rate) ** nper - 1) / rate + fv) / (1 + rate) ** nper
        if name == "FV":
            rate, nper, pmt, pv, kind = (numbers + [0.0, 0.0])[:5]
            return _fv(rate, nper, pmt, pv, kind)
        if name in ("IPMT", "PPMT"):
            rate, per, nper, pv, fv, kind = (numbers + [0.0, 0.0])[:6]
            if kind != 0:
                raise Unknown(f"{name} mit Zahlung am Anfang")
            if not 1 <= per <= nper:
                raise _Fail(Err("#NUM!"))
            pmt = _pmt(rate, nper, pv, fv, 0.0)
            interest = _fv(rate, per - 1, pmt, pv, 0.0) * rate
            return interest if name == "IPMT" else pmt - interest
        if name == "NPER":
            rate, pmt, pv, fv, kind = (numbers + [0.0, 0.0])[:5]
            if rate == 0:
                if pmt == 0:
                    raise _Fail(Err("#NUM!"))
                return -(pv + fv) / pmt
            top, bottom = pmt * (1 + rate * kind) - fv * rate, pmt * (1 + rate * kind) + pv * rate
            if bottom == 0 or top / bottom <= 0:
                raise _Fail(Err("#NUM!"))
            return math.log(top / bottom) / math.log(1 + rate)
        raise Unknown(f"Funktion {name}")

    def _array(self, n: Node):
        """Ein Argument von SUMMENPRODUKT Zelle für Zelle: ein Bereich als Liste, ein Ausdruck über Bereichen ebenso
        (SUMMENPRODUKT(C30:G30/(1+A51)^C31:G31)), alles andere als ein Wert."""
        if self.is_area(n):
            return self.values_of(n)
        if n.kind == "op" and n.value != "&" and _has_area(n):
            a, b = self._array(n.args[0]), self._array(n.args[1])
            if isinstance(a, list) and isinstance(b, list) and len(a) != len(b):
                raise _Fail(Err("#VALUE!"))
            size = len(a) if isinstance(a, list) else len(b)
            out = []
            for i in range(size):
                x = a[i] if isinstance(a, list) else a
                y = b[i] if isinstance(b, list) else b
                if isinstance(x, bool) and n.value not in ("=", "<>"):
                    x = float(x)
                if isinstance(y, bool) and n.value not in ("=", "<>"):
                    y = float(y)
                try:
                    out.append(self.operator(n.value, x, y))
                except _Fail as fail:
                    out.append(fail.err)
            return out
        if n.kind == "call" and n.value in _ELEMENTWISE and n.args and _has_area(n.args[0]) and \
                not any(_has_area(a) for a in n.args[1:]):
            # SUMMENPRODUKT(ABS(D2:D7)): die Funktion Zelle für Zelle.
            rest = tuple(Node("num", _num(self.eval(a))) for a in n.args[1:])
            out = []
            for v in self._array(n.args[0]):
                if isinstance(v, Err):
                    out.append(v)
                    continue
                try:
                    out.append(self.call(n.value, (Node("num", _num(v)),) + rest))
                except _Fail as fail:
                    out.append(fail.err)
            return out
        if n.kind in ("neg", "pct") and _has_area(n):
            inner = self._array(n.args[0])
            return [(-_num(v) if n.kind == "neg" else _num(v) / 100) for v in inner]
        return self.eval(n)

    def operator(self, op, a, b):
        if op == "&":
            return f"{_text(a)}{_text(b)}"
        if op in ("=", "<>", "<", ">", "<=", ">="):
            if isinstance(a, Err):
                raise _Fail(a)
            if isinstance(b, Err):
                raise _Fail(b)
            # Leer ist gleich "" und gleich 0, je nachdem, womit verglichen wird.
            if a is None and isinstance(b, str):
                a = ""
            if b is None and isinstance(a, str):
                b = ""
            x, y = _comparable(a), _comparable(b)
            if type(x) is not type(y):
                raise Unknown("Vergleich Zahl mit Text")
            return {"=": x == y, "<>": x != y, "<": x < y, ">": x > y, "<=": x <= y, ">=": x >= y}[op]
        x, y = _num(a), _num(b)
        if op == "+":
            return x + y
        if op == "-":
            return x - y
        if op == "*":
            return x * y
        if op == "/":
            if y == 0:
                raise _Fail(Err("#DIV/0!"))
            return x / y
        if op == "^":
            if x == 0 and y <= 0:
                raise _Fail(Err("#DIV/0!" if y < 0 else "#NUM!"))
            if x < 0 and y != int(y):
                # Gerade Wurzel aus einer negativen Zahl gibt #ZAHL!; was Excel bei anderen Brüchen tut, wissen wir nicht.
                root = 1 / y
                if abs(root - round(root)) < 1e-9 and round(root) % 2 == 0:
                    raise _Fail(Err("#NUM!"))
                raise Unknown("Potenz einer negativen Zahl")
            try:
                return float(x ** y)
            except (OverflowError, ZeroDivisionError, ValueError):
                raise _Fail(Err("#NUM!")) from None
        raise Unknown(op)

    def arg(self, args: tuple, i: int, default=None):
        """Das i-te Argument ausgerechnet, oder default, wenn es fehlt oder leer ist."""
        if i >= len(args) or args[i].kind == "empty":
            return default
        return self.eval(args[i])

    # --- BEREICH.VERSCHIEBEN und INDIREKT: Bezüge, die erst beim Rechnen feststehen ---

    def reference(self, node: Node) -> Node:
        """Der Bezug, den BEREICH.VERSCHIEBEN oder INDIREKT liefert, als ref, range, xref oder xrange."""
        if node.kind in ("ref", "range", "xref", "xrange"):
            return node
        if node.kind != "call" or node.value not in _REFERENCES:
            raise Unknown("kein Bezug")
        return self._offset(node.args) if node.value == "OFFSET" else self._indirect(node.args)

    def _whole(self, arg: Node) -> int:
        value = self.eval(arg)
        if isinstance(value, str) and _as_float(value) is None:
            raise _Fail(Err("#VALUE!"))
        number = _num(value)
        if not math.isfinite(number):
            raise _Fail(Err("#NUM!"))
        return math.trunc(number)

    def _offset(self, args: tuple) -> Node:
        """BEREICH.VERSCHIEBEN(Bezug; Zeilen; Spalten; [Höhe]; [Breite])."""
        if not 3 <= len(args) <= 5:
            raise Unknown("BEREICH.VERSCHIEBEN mit falscher Zahl von Argumenten")
        base = self.reference(args[0])
        sheet = base.value[0] if base.kind in ("xref", "xrange") else None
        where = base.value[1] if sheet is not None else base.value
        r1, c1, r2, c2 = where if len(where) == 4 else (*where, *where)
        rows, cols = self._whole(args[1]), self._whole(args[2])
        height = self._whole(args[3]) if len(args) > 3 and args[3].kind != "empty" else r2 - r1 + 1
        width = self._whole(args[4]) if len(args) > 4 and args[4].kind != "empty" else c2 - c1 + 1
        if height < 1 or width < 1:
            # Excel und LibreOffice gehen mit 0 und negativen Größen verschieden um.
            raise Unknown("BEREICH.VERSCHIEBEN ohne Größe")
        top, left = r1 + rows, c1 + cols
        bottom, right = top + height - 1, left + width - 1
        if top < 1 or left < 1 or bottom > MAX_ROW or right > MAX_COL:
            raise _Fail(Err("#REF!"))
        return _reference_node(sheet, top, left, bottom, right)

    def _indirect(self, args: tuple) -> Node:
        """INDIREKT(Text; [A1]): nur die A1-Schreibweise, Bezüge auf dieses oder ein anderes Blatt der Mappe."""
        if not 1 <= len(args) <= 2:
            raise Unknown("INDIREKT mit falscher Zahl von Argumenten")
        if len(args) == 2 and args[1].kind != "empty":
            a1 = self.eval(args[1])
            if isinstance(a1, Err):
                raise _Fail(a1)
            if not (a1 is True or (isinstance(a1, float) and a1 != 0)):
                raise Unknown("INDIREKT in Z1S1-Schreibweise")
        text = self.eval(args[0])
        if isinstance(text, Err):
            raise _Fail(text)
        if not isinstance(text, str) or not text.strip():
            raise _Fail(Err("#REF!"))
        try:
            node = parse(_upper_outside_quotes(text.strip()))
        except Unknown:
            raise Unknown("INDIREKT mit unbekanntem Text") from None
        if node.kind in ("ref", "range", "xref", "xrange"):
            return node
        if node.kind == "unknown":
            raise Unknown("INDIREKT mit einem Namen")
        raise _Fail(Err("#REF!"))

    def _address(self, args: tuple) -> str:
        """ADRESSE(Zeile; Spalte; [Art]; [A1]; [Blatt]): der Text eines Bezugs, wie ihn INDIREKT liest."""
        if not 2 <= len(args) <= 5:
            raise Unknown("ADRESSE mit falscher Zahl von Argumenten")
        row, col = self._whole(args[0]), self._whole(args[1])
        kind = self._whole(args[2]) if len(args) > 2 and args[2].kind != "empty" else 1
        if len(args) > 3 and args[3].kind != "empty":
            a1 = self.eval(args[3])
            if isinstance(a1, Err):
                raise _Fail(a1)
            if not (a1 is True or (isinstance(a1, float) and a1 != 0)):
                raise Unknown("ADRESSE in Z1S1-Schreibweise")
        if not (1 <= row <= MAX_ROW and 1 <= col <= MAX_COL) or kind not in (1, 2, 3, 4):
            raise _Fail(Err("#VALUE!"))
        letters = column_letters(col)
        text = {1: f"${letters}${row}", 2: f"{letters}${row}", 3: f"${letters}{row}", 4: f"{letters}{row}"}[kind]
        if len(args) > 4 and args[4].kind != "empty":
            sheet = self.eval(args[4])
            if isinstance(sheet, Err):
                raise _Fail(sheet)
            sheet = _text(sheet)
            if sheet:
                quoted = "'" + sheet.replace("'", "''") + "'" if re.search(r"[^\w.]", sheet) or sheet[:1].isdigit() else sheet
                text = f"{quoted}!{text}"
        return text

    def call(self, name: str, args: tuple):  # noqa: C901 - eine Funktion je Zweig
        if name in _REFERENCES:
            found = self.reference(Node("call", name, args))
            if found.kind in ("range", "xrange"):
                sheet = found.value[0] if found.kind == "xrange" else None
                r1, c1, r2, c2 = found.value[1] if sheet is not None else found.value
                if (r1, c1) != (r2, c2):
                    raise Unknown("Bereich ohne Funktion")
                return self._get(sheet, (r1, c1))
            return self.eval(found)
        if name == "ADDRESS":
            return self._address(args)
        if any(a.kind == "call" and a.value in _REFERENCES for a in args):
            # SUMME(BEREICH.VERSCHIEBEN(…)): Die Funktion bekommt den Bereich, als stünde er so da.
            args = tuple(self.reference(a) if a.kind == "call" and a.value in _REFERENCES else a for a in args)
        if name in _LOOKUPS:
            return getattr(self, f"_{name.lower().replace('.', '_')}")(args)
        if name in _TEXT_FUNCTIONS:
            return self._text_call(name, args)
        if name in _DATE_FUNCTIONS:
            return self._date_call(name, args)
        if name in ("STDEVP", "VARP"):
            name = name[:-1] + ".P"
        if name == "SUBTOTAL":
            return self._subtotal(args)
        if name == "AVERAGEA":
            # Text in Bereichen zählt als 0, WAHR als 1, Leeres gar nicht.
            values = []
            for v in self.flat(args):
                if isinstance(v, Err):
                    raise _Fail(v)
                if v is not None:
                    values.append(v if isinstance(v, float) and not isinstance(v, bool) else float(v is True))
            if not values:
                raise _Fail(Err("#DIV/0!"))
            return math.fsum(values) / len(values)
        if name in ("SUM", "AVERAGE", "COUNT") and len(args) == 1 and args[0].kind in ("range", "xrange"):
            summed = self._summed(args[0])
            if summed is not None:
                total, count = summed
                if name == "SUM":
                    return total
                if name == "COUNT":
                    return float(count)
                if not count:
                    raise _Fail(Err("#DIV/0!"))
                return total / count
        if name in ("SUM", "AVERAGE", "MIN", "MAX", "COUNT", "MEDIAN", "PRODUCT", "STDEV", "STDEV.S", "STDEV.P", "VAR", "VAR.S", "VAR.P"):
            values = _numbers(self.flat(args))
            if name == "SUM":
                return math.fsum(values)
            if name == "COUNT":
                return float(len(values))
            if name == "PRODUCT":
                return math.prod(values) if values else 0.0
            if name in ("MIN", "MAX"):
                return (min if name == "MIN" else max)(values) if values else 0.0
            if not values:
                raise _Fail(Err("#DIV/0!"))
            if name == "AVERAGE":
                return math.fsum(values) / len(values)
            if name == "MEDIAN":
                return float(statistics.median(values))
            if len(values) < 2 and name not in ("STDEV.P", "VAR.P"):
                raise _Fail(Err("#DIV/0!"))
            if name in ("STDEV", "STDEV.S"):
                return statistics.stdev(values)
            if name == "STDEV.P":
                return statistics.pstdev(values)
            return statistics.pvariance(values) if name == "VAR.P" else statistics.variance(values)
        if name in ("LARGE", "SMALL"):
            values = sorted(_numbers(self.flat(args[:1])), reverse=name == "LARGE")
            k = int(_num(self.eval(args[1])))
            if not 1 <= k <= len(values):
                raise _Fail(Err("#NUM!"))
            return values[k - 1]
        if name in ("PERCENTILE", "PERCENTILE.INC", "QUARTILE", "QUARTILE.INC"):
            values = sorted(_numbers(self.flat(args[:1])))
            k = _num(self.eval(args[1]))
            if name.startswith("QUARTILE"):
                if k not in (0, 1, 2, 3, 4):
                    raise _Fail(Err("#NUM!"))
                k /= 4
            if not values or not 0 <= k <= 1:
                raise _Fail(Err("#NUM!"))
            at = k * (len(values) - 1)
            low = math.floor(at)
            high = min(low + 1, len(values) - 1)
            return values[low] + (at - low) * (values[high] - values[low])
        if name in ("RANK", "RANK.EQ"):
            x = _num(self.eval(args[0]))
            values = _numbers(self.flat(args[1:2]))
            if x not in values:
                raise _Fail(Err("#N/A"))
            ascending = _truth(self.arg(args, 2, 0.0))
            return float(1 + sum(v < x if ascending else v > x for v in values))
        if name in ("ROWS", "COLUMNS") and len(args) == 1 and self.is_area(args[0]):
            r1, c1, r2, c2 = args[0].value if args[0].kind == "range" else args[0].value[1]
            return float(r2 - r1 + 1 if name == "ROWS" else c2 - c1 + 1)
        if name == "COUNTA":
            return float(sum(v is not None and v != "" for v in self.flat(args)))
        if name == "COUNTBLANK":
            return float(sum(v is None or v == "" for v in self.flat(args)))
        if name in _FINANCE:
            return self._finance(name, args)
        if name in ("ROUND", "ROUNDUP", "ROUNDDOWN"):
            x = _num(self.eval(args[0]))
            digits = _num(self.eval(args[1])) if len(args) > 1 else 0.0
            return _round(x, digits, {"ROUND": "half", "ROUNDUP": "up", "ROUNDDOWN": "down"}[name])
        if name == "MROUND":
            x, step = _num(self.eval(args[0])), _num(self.eval(args[1]))
            if step == 0:
                return 0.0
            if x * step < 0:
                raise _Fail(Err("#NUM!"))
            return _round(x / step, 0, "half") * step
        if name in ("INT", "TRUNC"):
            x = _num(self.eval(args[0]))
            if name == "INT":
                return float(math.floor(x))
            digits = int(_num(self.arg(args, 1, 0.0)))
            return math.trunc(x * 10 ** digits) / 10 ** digits
        if name == "ABS":
            return abs(_num(self.eval(args[0])))
        if name == "SIGN":
            x = _num(self.eval(args[0]))
            return float((x > 0) - (x < 0))
        if name == "SQRT":
            x = _num(self.eval(args[0]))
            if x < 0:
                raise _Fail(Err("#NUM!"))
            return math.sqrt(x)
        if name in ("LN", "LOG10", "EXP"):
            x = _num(self.eval(args[0]))
            if name == "EXP":
                return math.exp(x)
            if x <= 0:
                raise _Fail(Err("#NUM!"))
            return math.log(x) if name == "LN" else math.log10(x)
        if name in ("TRUE", "FALSE") and not args:
            # WAHR() und FALSCH() gibt es auch als Funktion; LibreOffice speichert FALSCH so.
            return name == "TRUE"
        if name == "NA" and not args:
            raise _Fail(Err("#N/A"))
        if name == "PI" and not args:
            return math.pi
        if name in _ANGLES and len(args) == 1:
            x = _num(self.eval(args[0]))
            if name in ("ASIN", "ACOS") and not -1 <= x <= 1:
                raise _Fail(Err("#NUM!"))
            return float(_ANGLES[name](x))
        if name == "ATAN2" and len(args) == 2:
            # Excel nimmt zuerst x, dann y (umgekehrt wie math.atan2).
            x, y = _num(self.eval(args[0])), _num(self.eval(args[1]))
            if x == 0 and y == 0:
                raise _Fail(Err("#DIV/0!"))
            return math.atan2(y, x)
        if name == "FACT" and len(args) == 1:
            x = math.trunc(_num(self.eval(args[0])))
            if not 0 <= x <= 170:
                raise _Fail(Err("#NUM!"))
            return float(math.factorial(x))
        if name == "CEILING" and len(args) == 2:
            x, step = _num(self.eval(args[0])), _num(self.eval(args[1]))
            if step == 0:
                return 0.0
            if x > 0 > step:
                raise _Fail(Err("#NUM!"))
            return float(math.ceil(round(x / step, 9)) * step)
        if name == "POWER":
            return self.operator("^", self.eval(args[0]), self.eval(args[1]))
        if name == "MOD":
            x, y = _num(self.eval(args[0])), _num(self.eval(args[1]))
            if y == 0:
                raise _Fail(Err("#DIV/0!"))
            return x - y * math.floor(x / y)
        if name == "QUOTIENT" and len(args) == 2:
            x, y = _num(self.eval(args[0])), _num(self.eval(args[1]))
            if y == 0:
                raise _Fail(Err("#DIV/0!"))
            return float(math.trunc(x / y))
        if name == "FLOOR" and len(args) == 2:
            x, step = _num(self.eval(args[0])), _num(self.eval(args[1]))
            if step == 0:
                raise _Fail(Err("#DIV/0!"))
            if x > 0 > step:
                raise _Fail(Err("#NUM!"))
            # Beide negativ rundet Excel zur Null hin, also auch dort: abrunden von x / step.
            return float(math.floor(x / step) * step)
        if name == "FLOOR.MATH" and 1 <= len(args) <= 3:
            x, step = _num(self.eval(args[0])), abs(_num(self.arg(args, 1, 1.0)))
            if step == 0:
                return 0.0
            toward = x < 0 and _num(self.arg(args, 2, 0.0)) != 0
            return float((math.ceil(x / step) if toward else math.floor(x / step)) * step)
        if name == "IF":
            test = self.eval(args[0])
            if isinstance(test, Err):
                raise _Fail(test)
            chosen = args[1] if _truth(test) else (args[2] if len(args) > 2 else None)
            if chosen is None:
                return False
            return None if chosen.kind == "empty" else self.eval(chosen)
        if name == "IFS":
            for i in range(0, len(args) - 1, 2):
                test = self.eval(args[i])
                if isinstance(test, Err):
                    raise _Fail(test)
                if _truth(test):
                    return self.eval(args[i + 1])
            raise _Fail(Err("#N/A"))
        if name == "SWITCH":
            key = self.eval(args[0])
            rest = args[1:]
            for i in range(0, len(rest) - 1, 2):
                if _same(self.eval(rest[i]), key):
                    return self.eval(rest[i + 1])
            if len(rest) % 2:
                return self.eval(rest[-1])
            raise _Fail(Err("#N/A"))
        if name == "CHOOSE":
            k = int(_num(self.eval(args[0])))
            if not 1 <= k < len(args):
                raise _Fail(Err("#VALUE!"))
            return self.eval(args[k])
        if name in ("IFERROR", "IFNA"):
            try:
                v = self.eval(args[0])
            except _Fail as fail:
                if name == "IFERROR" or fail.err.code in ("#N/A", "#NV"):
                    return self.eval(args[1])
                raise
            if isinstance(v, Err) and (name == "IFERROR" or v.code in ("#N/A", "#NV")):
                return self.eval(args[1])
            return v
        if name in ("ISBLANK", "ISNUMBER", "ISTEXT", "ISERROR", "ISNA", "ISERR", "ISLOGICAL"):
            try:
                v = self.eval(args[0])
            except _Fail as fail:
                v = fail.err
            if name == "ISBLANK":
                return v is None
            if name == "ISNUMBER":
                return isinstance(v, float) and not isinstance(v, bool)
            if name == "ISTEXT":
                return isinstance(v, str)
            if name == "ISLOGICAL":
                return isinstance(v, bool)
            if name == "ISNA":
                return isinstance(v, Err) and v.code in ("#N/A", "#NV")
            if name == "ISERR":
                return isinstance(v, Err) and v.code not in ("#N/A", "#NV")
            return isinstance(v, Err)
        if name == "N":
            v = self.eval(args[0])
            return v if isinstance(v, float) and not isinstance(v, bool) else float(v) if isinstance(v, bool) else 0.0
        if name in ("AND", "OR"):
            values = [_truth(v) for v in self.flat(args) if v is not None]
            return all(values) if name == "AND" else any(values)
        if name == "NOT":
            return not _truth(self.eval(args[0]))
        if name == "MMULT" and len(args) == 2 and all(self.is_area(a) for a in args):
            # Zeile mal Spalte (eine Markov-Kette: Verteilung des Vorjahrs mal Spalte der Übergangsmatrix). Ergibt das
            # mehr als einen Wert, verteilt Excel ihn auf mehrere Zellen: das rechnet Gegenprobe nicht nach.
            left, right = self.grid(args[0]), self.grid(args[1])
            if len(left[0]) != len(right):
                raise _Fail(Err("#VALUE!"))
            if len(left) != 1 or len(right[0]) != 1:
                raise Unknown("MMULT mit mehreren Werten")
            if any(isinstance(v, (str, bool)) or v is None for v in left[0] + [r[0] for r in right]):
                raise _Fail(Err("#VALUE!"))
            return sum(_num(a) * _num(r[0]) for a, r in zip(left[0], right))
        if name == "SUMPRODUCT":
            columns = [self._array(a) for a in args]
            columns = [c if isinstance(c, list) else [c] for c in columns]
            if len({len(c) for c in columns}) != 1:
                raise _Fail(Err("#VALUE!"))
            # WAHR zählt hier als 1 (SUMMENPRODUKT((B2:B9="Nord")*C2:C9)), Text und Leeres als 0.
            return math.fsum(math.prod(float(v) if isinstance(v, (float, bool)) else 0.0 for v in row)
                             for row in zip(*columns))
        if name in ("SUMIF", "COUNTIF", "AVERAGEIF"):
            if not self.is_area(args[0]):
                raise Unknown("Bedingung ohne Bereich")
            wanted = self.eval(args[1])
            test = _criterion(wanted)
            if name == "COUNTIF":
                fast = self._conditional(name, [args[0]], [wanted], None)
                return fast if fast is not None else float(sum(test(v) for v in self.values_of(args[0])))
            own = not (len(args) > 2 and self.is_area(args[2]))
            if not own and self._size(args[2]) != self._size(args[0]):
                raise Unknown("ungleiche Bereiche")
            fast = self._conditional(name, [args[0]], [wanted], args[0] if own else args[2])
            if fast is not None:
                return fast
            where = self.values_of(args[0])
            source = where if own else self.values_of(args[2])
            picked = [s for w, s in zip(where, source) if test(w) and isinstance(s, float)]
            if name == "SUMIF":
                return math.fsum(picked)
            if not picked:
                raise _Fail(Err("#DIV/0!"))
            return math.fsum(picked) / len(picked)
        if name in ("SUMIFS", "COUNTIFS", "AVERAGEIFS", "MINIFS", "MAXIFS"):
            return self._ifs(name, args)
        if name == "DATE" and len(args) == 3:
            return _excel_date(*(math.trunc(_num(self.eval(a))) for a in args))
        if name in _TESTS:
            return self._test_call(name, args)
        if name in _PAIRED:
            return self._paired(name, args)
        if name in _SHAPE:
            return _shape(name, _numbers(self.flat(args)))
        if name == "STANDARDIZE" and len(args) == 3:
            x, mean, sd = (_num(self.eval(a)) for a in args)
            if sd <= 0:
                raise _Fail(Err("#NUM!"))
            return (x - mean) / sd
        raise Unknown(f"Funktion {name}")

    def _paired(self, name: str, args: tuple) -> float:
        """KORREL, STEIGUNG und Verwandte: Zahlenpaare an gleicher Stelle. Steht in einem der beiden Bereiche Text
        oder nichts, lässt Excel das ganze Paar weg; ungleich große Bereiche sind #NV."""
        if name in ("FORECAST", "FORECAST.LINEAR"):
            if len(args) != 3:
                raise Unknown(f"{name} mit {len(args)} Argumenten")
            x = _num(self.eval(args[0]))
            ys, xs = args[1], args[2]
        else:
            if len(args) != 2:
                raise Unknown(f"{name} mit {len(args)} Argumenten")
            x, (ys, xs) = None, args
        left, right = self.flat([ys]), self.flat([xs])
        if len(left) != len(right):
            raise _Fail(Err("#N/A"))
        for v in left + right:
            if isinstance(v, Err):
                raise _Fail(v)
        pairs = [(a, b) for a, b in zip(left, right)
                 if isinstance(a, float) and not isinstance(a, bool) and isinstance(b, float) and not isinstance(b, bool)]
        n = len(pairs)
        if n < (1 if name == "COVARIANCE.P" or name == "COVAR" else 2):
            raise _Fail(Err("#DIV/0!"))
        ma = math.fsum(a for a, _ in pairs) / n
        mb = math.fsum(b for _, b in pairs) / n
        sab = math.fsum((a - ma) * (b - mb) for a, b in pairs)
        saa = math.fsum((a - ma) ** 2 for a, _ in pairs)
        sbb = math.fsum((b - mb) ** 2 for _, b in pairs)
        if name in ("COVARIANCE.P", "COVAR"):
            return sab / n
        if name == "COVARIANCE.S":
            return sab / (n - 1)
        # Erster Bereich: y (abhängig), zweiter: x. Bei KORREL ist die Reihenfolge gleich.
        if name in ("CORREL", "PEARSON", "RSQ"):
            if not saa or not sbb:
                raise _Fail(Err("#DIV/0!"))
            r = sab / math.sqrt(saa * sbb)
            return r * r if name == "RSQ" else r
        if not sbb:
            raise _Fail(Err("#DIV/0!"))
        slope = sab / sbb
        if name == "SLOPE":
            return slope
        if name == "INTERCEPT":
            return ma - slope * mb
        if name == "STEYX":
            if n < 3:
                raise _Fail(Err("#DIV/0!"))
            return math.sqrt(max(0.0, (saa - sab * sab / sbb) / (n - 2)))
        return ma + slope * (x - mb)

    def _test_call(self, name: str, args: tuple):  # noqa: C901 - eine Funktion je Zweig
        """T.TEST, Konfidenzintervalle und die t- und Normalverteilung, so wie Excel sie rechnet."""
        if len(args) != _TESTS[name]:
            raise Unknown(f"{name} mit {len(args)} Argumenten")
        if name in ("T.TEST", "TTEST"):
            if not (self.is_area(args[0]) and self.is_area(args[1])):
                raise Unknown("T.TEST ohne zwei Bereiche")
            tails, kind = math.trunc(_num(self.eval(args[2]))), math.trunc(_num(self.eval(args[3])))
            if tails not in (1, 2) or kind not in (1, 2, 3):
                raise _Fail(Err("#NUM!"))
            t, df = t_statistic(self.values_of(args[0]), self.values_of(args[1]), kind)
            return tails * t_tail(abs(t), df)
        if name in ("CONFIDENCE", "CONFIDENCE.NORM", "CONFIDENCE.T"):
            alpha, sd, size = (_num(self.eval(a)) for a in args[:3])
            size = math.trunc(size)
            if not 0 < alpha < 1 or sd <= 0 or size < 1:
                raise _Fail(Err("#NUM!"))
            if name != "CONFIDENCE.T":
                return statistics.NormalDist().inv_cdf(1 - alpha / 2) * sd / math.sqrt(size)
            if size == 1:
                raise _Fail(Err("#DIV/0!"))
            return t_inverse(alpha, size - 1) * sd / math.sqrt(size)
        if name in ("NORM.S.INV", "NORMSINV"):
            p = _num(self.eval(args[0]))
            if not 0 < p < 1:
                raise _Fail(Err("#NUM!"))
            return statistics.NormalDist().inv_cdf(p)
        if name in ("NORM.S.DIST", "NORMSDIST"):
            z = _num(self.eval(args[0]))
            if name == "NORM.S.DIST" and not _truth(self.eval(args[1])):
                return statistics.NormalDist().pdf(z)
            return statistics.NormalDist().cdf(z)
        x, df = _num(self.eval(args[0])), math.trunc(_num(self.eval(args[1])))
        if df < 1:
            raise _Fail(Err("#NUM!"))
        if name in ("T.INV.2T", "TINV"):
            if not 0 < x <= 1:
                raise _Fail(Err("#NUM!"))
            return t_inverse(x, df)
        if name == "T.INV":
            if not 0 < x < 1:
                raise _Fail(Err("#NUM!"))
            return t_inverse(2 * min(x, 1 - x), df) * (1 if x > 0.5 else -1) if x != 0.5 else 0.0
        if name == "T.DIST":
            if not _truth(self.eval(args[2])):
                return math.exp(math.lgamma((df + 1) / 2) - math.lgamma(df / 2)) / math.sqrt(df * math.pi) * \
                    (1 + x * x / df) ** (-(df + 1) / 2)
            return 1 - t_tail(x, df) if x >= 0 else t_tail(-x, df)
        if name == "T.DIST.RT":
            return t_tail(x, df) if x >= 0 else 1 - t_tail(-x, df)
        if x < 0:
            raise _Fail(Err("#NUM!"))
        if name == "T.DIST.2T":
            return 2 * t_tail(x, df)
        tails = math.trunc(_num(self.eval(args[2])))
        if tails not in (1, 2):
            raise _Fail(Err("#NUM!"))
        return tails * t_tail(x, df)

    def _ifs(self, name: str, args: tuple):
        """SUMMEWENNS und Verwandte: alle Bedingungen müssen für dieselbe Zeile gelten."""
        pairs = args if name == "COUNTIFS" else args[1:]
        if not pairs or len(pairs) % 2:
            raise Unknown("Bedingungen ohne Paar")
        areas = [pairs[i] for i in range(0, len(pairs), 2)]
        if any(not self.is_area(a) for a in areas) or (name != "COUNTIFS" and not self.is_area(args[0])):
            raise Unknown("Bedingung ohne Bereich")
        wanted = [self.eval(pairs[i + 1]) for i in range(0, len(pairs), 2)]
        tests = [_criterion(w) for w in wanted]
        sizes = {self._size(a) for a in areas + ([] if name == "COUNTIFS" else [args[0]])}
        if len(sizes) != 1:
            raise _Fail(Err("#VALUE!"))
        fast = self._conditional(name, areas, wanted, None if name == "COUNTIFS" else args[0])
        if fast is not None:
            return fast
        columns = [self.values_of(a) for a in areas]
        if name != "COUNTIFS":
            columns.insert(0, self.values_of(args[0]))
        offset = 0 if name == "COUNTIFS" else 1
        hits = [all(t(col[k]) for t, col in zip(tests, columns[offset:])) for k in range(len(columns[0]))]
        found = [columns[0][k] for k, hit in enumerate(hits) if hit]
        if name == "COUNTIFS":
            return float(len(found))
        picked = [v for v in found if isinstance(v, float) and not isinstance(v, bool)]
        if name == "SUMIFS":
            return math.fsum(picked)
        if name == "AVERAGEIFS":
            if not picked:
                raise _Fail(Err("#DIV/0!"))
            return math.fsum(picked) / len(picked)
        if not picked:
            return 0.0
        return min(picked) if name == "MINIFS" else max(picked)

    # --- Nachschlagen: SVERWEIS, WVERWEIS, XVERWEIS, VERGLEICH, XVERGLEICH, INDEX ---

    def _key(self, node: Node):
        key = self.eval(node)
        if isinstance(key, Err):
            raise _Fail(key)
        return "" if key is None else key

    def _line(self, node: Node) -> list:
        """Ein Bereich mit nur einer Zeile oder Spalte als Liste."""
        rows = self.grid(node)
        if len(rows) == 1:
            return rows[0]
        if all(len(r) == 1 for r in rows):
            return [r[0] for r in rows]
        raise Unknown("Bereich mit mehreren Zeilen und Spalten")

    def _vlookup(self, args: tuple):
        return self._table_lookup(args, by_rows=True)

    def _hlookup(self, args: tuple):
        return self._table_lookup(args, by_rows=False)

    def _table_lookup(self, args: tuple, by_rows: bool):
        key = self._key(args[0])
        rows = self.grid(args[1])
        table = rows if by_rows else [list(col) for col in zip(*rows)]
        k = int(_num(self.eval(args[2])))
        if k < 1:
            raise _Fail(Err("#VALUE!"))
        if not table or k > len(table[0]):
            raise _Fail(Err("#REF!"))
        approx = _truth(self.arg(args, 3, True))
        first = [r[0] for r in table]
        at = _find(key, first, 1) if approx else self._exact(key, first, ("first", by_rows, self.area_key(args[1])))
        if self.hits is not None:
            sheet, (r1, c1, _, _) = self.area_of(args[1])
            self.hits.append(("VLOOKUP" if by_rows else "HLOOKUP", sheet, (r1 + at, c1 + k - 1) if by_rows else (r1 + k - 1, c1 + at)))
        return table[at][k - 1]

    def _match(self, args: tuple):
        key = self._key(args[0])
        mode = max(-1, min(1, int(_num(self.arg(args, 2, 1.0)))))
        line = self._line(args[1])
        return float((self._exact(key, line, ("line", self.area_key(args[1]))) if mode == 0 else _find(key, line, mode)) + 1)

    def _xmatch(self, args: tuple):
        key = self._key(args[0])
        found = self._xsearch(key, args[1], int(_num(self.arg(args, 2, 0.0))), int(_num(self.arg(args, 3, 1.0))))
        if found is None:
            raise _Fail(Err("#N/A"))
        return float(found + 1)

    def _xsearch(self, key, node: Node, match: int, search: int) -> int | None:
        line = self._line(node)
        if match == 0 and search == 1:
            try:
                return self._exact(key, line, ("line", self.area_key(node)))
            except _Fail:
                return None
        return _xfind(key, line, match, search)

    def _xlookup(self, args: tuple):
        key = self._key(args[0])
        where = self._line(args[1])
        source = self._line(args[2])
        if len(source) != len(where):
            raise _Fail(Err("#VALUE!"))
        found = self._xsearch(key, args[1], int(_num(self.arg(args, 4, 0.0))), int(_num(self.arg(args, 5, 1.0))))
        if found is None:
            if len(args) > 3 and args[3].kind != "empty":
                return self.eval(args[3])
            raise _Fail(Err("#N/A"))
        if self.hits is not None:
            sheet, (r1, c1, r2, _) = self.area_of(args[2])
            self.hits.append(("XLOOKUP", sheet, (r1 + found, c1) if r2 > r1 else (r1, c1 + found)))
        return source[found]

    def _index(self, args: tuple):
        rows = self.grid(args[0]) if self.is_area(args[0]) or args[0].kind in ("ref", "xref") else None
        if rows is None:
            raise Unknown("INDEX ohne Bereich")
        r = int(_num(self.arg(args, 1, 0.0)))
        c = int(_num(self.arg(args, 2, 0.0)))
        if len(args) < 3 and len(rows) == 1:
            r, c = 1, r  # INDEX(A1:F1; 3): in einer Zeile zählt die Zahl die Spalten.
        if len(args) < 3 and len(rows) > 1 and all(len(x) == 1 for x in rows):
            c = 1
        if r == 0 or c == 0:
            if len(rows) == 1 and r == 0:
                r = 1
            elif all(len(x) == 1 for x in rows) and c == 0:
                c = 1
            else:
                raise Unknown("INDEX auf ganze Zeile oder Spalte")
        if not (1 <= r <= len(rows) and 1 <= c <= len(rows[0])):
            raise _Fail(Err("#REF!"))
        if self.hits is not None:
            sheet, (r1, c1, _, _) = self.area_of(args[0])
            self.hits.append(("INDEX", sheet, (r1 + r - 1, c1 + c - 1)))
        return rows[r - 1][c - 1]

    # --- Text und Datum ---

    def _text_call(self, name: str, args: tuple):
        if name in ("CONCAT", "CONCATENATE"):
            return "".join(_text(v) for v in self.flat(args))
        if name == "TEXTJOIN":
            sep = _text(self.eval(args[0]))
            skip = _truth(self.eval(args[1]))
            parts = [_text(v) for v in self.flat(args[2:])]
            return sep.join(p for p in parts if p or not skip)
        text = _text(self.eval(args[0]))
        if name == "LEN":
            return float(len(text))
        if name == "UPPER":
            return text.upper()
        if name == "LOWER":
            return text.lower()
        if name == "TRIM":
            return re.sub(" +", " ", text).strip(" ")
        if name in ("LEFT", "RIGHT"):
            n = int(_num(self.arg(args, 1, 1.0)))
            if n < 0:
                raise _Fail(Err("#VALUE!"))
            return text[:n] if name == "LEFT" else (text[-n:] if n else "")
        if name == "MID":
            start, n = int(_num(self.eval(args[1]))), int(_num(self.eval(args[2])))
            if start < 1 or n < 0:
                raise _Fail(Err("#VALUE!"))
            return text[start - 1:start - 1 + n]
        if name == "REPT":
            n = int(_num(self.eval(args[1])))
            if n < 0:
                raise _Fail(Err("#VALUE!"))
            return text * n
        if name == "EXACT":
            return text == _text(self.eval(args[1]))
        if name in ("FIND", "SEARCH"):
            within = _text(self.eval(args[1]))
            start = int(_num(self.arg(args, 2, 1.0)))
            if name == "FIND":
                at = within.find(text, start - 1) if start >= 1 else -1
            else:
                at = within.lower().find(text.lower(), start - 1) if start >= 1 and "*" not in text and "?" not in text else None
                if at is None:
                    raise Unknown("SUCHEN mit Platzhalter")
            if at < 0:
                raise _Fail(Err("#VALUE!"))
            return float(at + 1)
        if name == "SUBSTITUTE":
            old, new = _text(self.eval(args[1])), _text(self.eval(args[2]))
            if len(args) > 3:
                raise Unknown("WECHSELN mit Nummer")
            return text.replace(old, new) if old else text
        if name == "VALUE":
            value = _as_float(text)
            if value is None:
                raise _Fail(Err("#VALUE!"))
            return value
        raise Unknown(f"Funktion {name}")

    def _date_call(self, name: str, args: tuple):
        if name in ("TODAY", "NOW"):
            raise Unknown("ändert sich jeden Tag")
        serial = _num(self.eval(args[0]))
        day = _day(serial)
        if name == "YEAR":
            return float(day.year)
        if name == "MONTH":
            return float(day.month)
        if name == "DAY":
            return float(day.day)
        if name == "WEEKDAY":
            kind = int(_num(self.arg(args, 1, 1.0)))
            monday = day.weekday()  # 0 = Montag
            if kind == 1:
                return float((monday + 1) % 7 + 1)
            if kind == 2:
                return float(monday + 1)
            if kind == 3:
                return float(monday)
            raise Unknown("WOCHENTAG mit dieser Art")
        if name in ("EDATE", "EOMONTH"):
            months = int(_num(self.eval(args[1])))
            total = day.year * 12 + day.month - 1 + months
            year, month = divmod(total, 12)
            month += 1
            last = calendar.monthrange(year, month)[1]
            if name == "EOMONTH":
                return _serial(year, month, last)
            return _serial(year, month, min(day.day, last))
        if name == "DATEDIF":
            end = _day(_num(self.eval(args[1])))
            unit = _text(self.eval(args[2])).upper()
            if end < day:
                raise _Fail(Err("#NUM!"))
            if unit == "D":
                return float((end - day).days)
            months = (end.year - day.year) * 12 + end.month - day.month - (end.day < day.day)
            if unit == "M":
                return float(months)
            if unit == "Y":
                return float(months // 12)
            raise Unknown("DATEDIF mit dieser Einheit")
        if name in ("NETWORKDAYS", "WORKDAY"):
            if len(args) not in (2, 3):
                raise Unknown(f"{'NETTOARBEITSTAGE' if name == 'NETWORKDAYS' else 'ARBEITSTAG'} mit dieser Anzahl Argumente")
            free = set()
            if len(args) == 3:
                values = self.values_of(args[2]) if self.is_area(args[2]) else [self.eval(args[2])]
                for v in values:
                    if v is None:
                        continue
                    if isinstance(v, Err):
                        raise _Fail(v)
                    if not isinstance(v, float) or isinstance(v, bool):
                        raise _Fail(Err("#VALUE!"))
                    free.add(math.floor(v))
            if name == "WORKDAY":
                return float(workday(math.floor(serial), math.trunc(_num(self.eval(args[1]))), free))
            return float(workdays(math.floor(serial), math.floor(_num(self.eval(args[1]))), free))
        raise Unknown(f"Funktion {name}")


_LOOKUPS = {"VLOOKUP", "HLOOKUP", "MATCH", "XMATCH", "XLOOKUP", "INDEX"}
def workday(start: int, days: int, free: set[int] = frozenset()) -> int:
    """ARBEITSTAG: der Tag, days Arbeitstage (Montag bis Freitag, ohne die freien Tage) nach start, rückwärts bei
    negativem days; bei 0 der Tag selbst, wie in Excel."""
    step, day = (1 if days > 0 else -1), start
    if abs(days) > 100_000:
        raise Unknown("ARBEITSTAG über so viele Tage")
    left = abs(days)
    while left:
        day += step
        if (day + 5) % 7 < 5 and day not in free:
            left -= 1
    return day


def workdays(start: int, end: int, free: set[int] = frozenset()) -> int:
    """NETTOARBEITSTAGE: Montag bis Freitag von start bis end, beide mitgezählt, ohne die freien Tage; rückwärts
    negativ, wie in Excel. Die Zahl 0 ist ein Samstag (30.12.1899)."""
    sign = 1
    if start > end:
        start, end, sign = end, start, -1
    full, rest = divmod(end - start + 1, 7)
    count = full * 5 + sum(1 for d in range(start + full * 7, end + 1) if (d + 5) % 7 < 5)
    count -= sum(1 for d in free if start <= d <= end and (d + 5) % 7 < 5)
    return sign * count


# Die Funktionen für Tests und Intervalle, je mit der Anzahl ihrer Argumente.
_ELEMENTWISE = {"ABS", "SQRT", "LN", "EXP", "INT", "SIGN", "ROUND", "ROUNDUP", "ROUNDDOWN", "LOG10"}
_PAIRED = {"CORREL", "PEARSON", "RSQ", "SLOPE", "INTERCEPT", "STEYX", "COVARIANCE.S", "COVARIANCE.P", "COVAR",
           "FORECAST", "FORECAST.LINEAR"}
_SHAPE = {"GEOMEAN", "HARMEAN", "DEVSQ", "AVEDEV", "SKEW", "KURT", "MODE", "MODE.SNGL"}


def _shape(name: str, values: list[float]) -> float:
    """Mittel, Schiefe und Wölbung einer Zahlenreihe, wie Excel sie rechnet."""
    n = len(values)
    if not n:
        raise _Fail(Err("#NUM!" if name in ("GEOMEAN", "HARMEAN", "DEVSQ") else "#N/A" if name.startswith("MODE")
                        else "#DIV/0!"))
    if name in ("GEOMEAN", "HARMEAN"):
        if any(v <= 0 for v in values):
            raise _Fail(Err("#NUM!"))
        if name == "GEOMEAN":
            return math.exp(math.fsum(math.log(v) for v in values) / n)
        return n / math.fsum(1 / v for v in values)
    if name.startswith("MODE"):
        counts: dict = {}
        for v in values:
            counts[v] = counts.get(v, 0) + 1
        best = max(counts.values())
        if best < 2:
            raise _Fail(Err("#N/A"))
        # Bei Gleichstand nimmt Excel den Wert, der zuerst vorkommt.
        return next(v for v in values if counts[v] == best)
    mean = math.fsum(values) / n
    dev = [v - mean for v in values]
    if name == "DEVSQ":
        return math.fsum(d * d for d in dev)
    if name == "AVEDEV":
        return math.fsum(abs(d) for d in dev) / n
    if n < (3 if name == "SKEW" else 4):
        raise _Fail(Err("#DIV/0!"))
    sd = math.sqrt(math.fsum(d * d for d in dev) / (n - 1))
    if not sd:
        raise _Fail(Err("#DIV/0!"))
    if name == "SKEW":
        return n / ((n - 1) * (n - 2)) * math.fsum((d / sd) ** 3 for d in dev)
    return (n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * math.fsum((d / sd) ** 4 for d in dev)
            - 3 * (n - 1) ** 2 / ((n - 2) * (n - 3)))


_TESTS = {"T.TEST": 4, "TTEST": 4, "CONFIDENCE": 3, "CONFIDENCE.NORM": 3, "CONFIDENCE.T": 3, "NORM.S.INV": 1,
          "NORMSINV": 1, "NORM.S.DIST": 2, "NORMSDIST": 1, "T.INV.2T": 2, "TINV": 2, "T.INV": 2, "T.DIST": 3,
          "T.DIST.RT": 2, "T.DIST.2T": 2, "TDIST": 3}


def t_statistic(one: list, two: list, kind: int) -> tuple[float, float]:
    """t und Freiheitsgrade wie T.TEST: 1 gepaart (Zeile für Zeile), 2 gleiche Varianzen (gepoolt), 3 Welch mit den
    Freiheitsgraden nach Welch-Satterthwaite, ungerundet wie in Excel."""
    if kind == 1:
        if len(one) != len(two):
            raise _Fail(Err("#N/A"))
        if any(not isinstance(v, float) or isinstance(v, bool) for v in one + two):
            # Wie Excel Paare mit einer leeren Zelle zählt, rechnet der Prüfer nicht nach.
            raise Unknown("gepaart mit Lücken")
        diffs = [a - b for a, b in zip(one, two)]
        if len(diffs) < 2 or statistics.stdev(diffs) == 0:
            raise _Fail(Err("#DIV/0!"))
        return statistics.fmean(diffs) / (statistics.stdev(diffs) / math.sqrt(len(diffs))), len(diffs) - 1.0
    x, y = _numbers(one), _numbers(two)
    n1, n2 = len(x), len(y)
    if n1 < 2 or n2 < 2:
        raise _Fail(Err("#DIV/0!"))
    v1, v2 = statistics.variance(x), statistics.variance(y)
    if kind == 2:
        df = n1 + n2 - 2.0
        se = math.sqrt(((n1 - 1) * v1 + (n2 - 1) * v2) / df * (1 / n1 + 1 / n2))
    else:
        a, b = v1 / n1, v2 / n2
        se = math.sqrt(a + b)
        df = (a + b) ** 2 / (a * a / (n1 - 1) + b * b / (n2 - 1)) if a + b > 0 else 1.0
    if se == 0:
        raise _Fail(Err("#DIV/0!"))
    return (statistics.fmean(x) - statistics.fmean(y)) / se, df


def t_tail(t: float, df: float) -> float:
    """P(T > t) für t ≥ 0 bei df Freiheitsgraden (auch gebrochen, wie bei Welch)."""
    return 0.5 * _beta(df / 2, 0.5, df / (df + t * t))


def t_inverse(p: float, df: float) -> float:
    """Das t, über dem (in beiden Richtungen zusammen) die Wahrscheinlichkeit p liegt: T.INV.2T."""
    if p >= 1:
        return 0.0
    low, high = 0.0, 1.0
    while 2 * t_tail(high, df) > p:
        low, high = high, high * 2
        if high > 1e12:
            raise _Fail(Err("#NUM!"))
    for _ in range(200):
        mid = (low + high) / 2
        if mid in (low, high):
            break
        if 2 * t_tail(mid, df) > p:
            low = mid
        else:
            high = mid
    return (low + high) / 2


def _beta(a: float, b: float, x: float) -> float:
    """Die regularisierte unvollständige Betafunktion I_x(a, b), als Kettenbruch (Lentz)."""
    if x <= 0:
        return 0.0
    if x >= 1:
        return 1.0
    front = math.exp(math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b) + a * math.log(x) + b * math.log1p(-x))
    if x < (a + 1) / (a + b + 2):
        return front * _fraction(a, b, x) / a
    return 1 - front * _fraction(b, a, 1 - x) / b


def _fraction(a: float, b: float, x: float) -> float:
    tiny = 1e-300
    c, d = 1.0, 1 - (a + b) * x / (a + 1)
    d = 1 / (d if abs(d) > tiny else tiny)
    h = d
    for m in range(1, 500):
        for aa in (m * (b - m) * x / ((a + 2 * m - 1) * (a + 2 * m)), -(a + m) * (a + b + m) * x / ((a + 2 * m) * (a + 2 * m + 1))):
            d = 1 + aa * d
            d = 1 / (d if abs(d) > tiny else tiny)
            c = 1 + aa / c
            c = c if abs(c) > tiny else tiny
            h *= d * c
        if abs(d * c - 1) < 1e-16:
            break
    return h
_TEXT_FUNCTIONS = {"CONCAT", "CONCATENATE", "TEXTJOIN", "LEN", "UPPER", "LOWER", "TRIM", "LEFT", "RIGHT", "MID", "REPT",
                   "EXACT", "FIND", "SEARCH", "SUBSTITUTE", "VALUE"}
_ANGLES = {"SIN": math.sin, "COS": math.cos, "TAN": math.tan, "ASIN": math.asin, "ACOS": math.acos, "ATAN": math.atan,
           "DEGREES": math.degrees, "RADIANS": math.radians}
_DATE_FUNCTIONS = {"TODAY", "NOW", "YEAR", "MONTH", "DAY", "WEEKDAY", "EDATE", "EOMONTH", "DATEDIF", "NETWORKDAYS",
                   "WORKDAY"}


def _same(a, b) -> bool:
    """Gleich, wie SVERWEIS vergleicht: Zahl mit Zahl, Text mit Text ohne Groß- und Kleinschreibung."""
    if isinstance(a, bool) or isinstance(b, bool):
        return isinstance(a, bool) and isinstance(b, bool) and a == b
    if isinstance(a, float) and isinstance(b, float):
        return a == b
    if isinstance(a, str) and isinstance(b, str):
        return a.casefold() == b.casefold()
    return False


def _norm(value):
    """Schlüssel für das Verzeichnis: so gleich, wie _same vergleicht."""
    if isinstance(value, bool):
        return ("b", value)
    if isinstance(value, float):
        return ("n", value)
    if isinstance(value, str) and value != "":
        return ("s", value.casefold())
    return None


def _kind(value) -> str:
    if isinstance(value, bool):
        return "b"
    if isinstance(value, float):
        return "n"
    if isinstance(value, str) and value != "":
        return "s"
    return ""


def _order_key(value):
    return value.casefold() if isinstance(value, str) else value


def _find(key, line: list, mode: int) -> int:
    """VERGLEICH: 0 genau (mit * und ? bei Text), 1 größter Wert ≤ key, -1 kleinster Wert ≥ key.

    Ungefähr sucht Excel binär und verlässt sich darauf, dass die Werte sortiert sind. Sind sie es nicht,
    ist das Ergebnis Glückssache: Dann rechnet der Prüfer nicht nach (und meldet es in analyse.py)."""
    if mode == 0:
        if isinstance(key, str) and re.search(r"(?<!~)[*?]", key):
            test = _criterion(key)
            hits = [i for i, v in enumerate(line) if isinstance(v, str) and test(v)]
        else:
            hits = [i for i, v in enumerate(line) if _same(v, key)]
        if not hits:
            raise _Fail(Err("#N/A"))
        return hits[0]
    kind = _kind(key)
    same = [(i, v) for i, v in enumerate(line) if _kind(v) == kind and kind]
    keys = [_order_key(v) for _, v in same]
    ordered = keys == sorted(keys) if mode == 1 else keys == sorted(keys, reverse=True)
    if not ordered:
        raise Unknown("ungefähre Suche in unsortierten Werten")
    k = _order_key(key)
    best = None
    for i, v in same:
        if (mode == 1 and _order_key(v) <= k) or (mode == -1 and _order_key(v) >= k):
            best = i
        else:
            break
    if best is None:
        raise _Fail(Err("#N/A"))
    return best


def _xfind(key, line: list, match: int, search: int) -> int | None:
    """XVERGLEICH: 0 genau, -1 genau oder nächstkleiner, 1 genau oder nächstgrößer, 2 mit Platzhaltern."""
    if search not in (1, -1):
        raise Unknown("binäre Suche")
    order = list(range(len(line))) if search == 1 else list(range(len(line) - 1, -1, -1))
    if match == 2:
        test = _criterion(key) if isinstance(key, str) else (lambda v: _same(v, key))
        return next((i for i in order if test(line[i])), None)
    exact = next((i for i in order if _same(line[i], key)), None)
    if exact is not None or match == 0:
        return exact
    if match not in (-1, 1):
        raise Unknown("XVERGLEICH mit dieser Art")
    kind = _kind(key)
    k = _order_key(key)
    candidates = [i for i in order if _kind(line[i]) == kind and kind and ((_order_key(line[i]) < k) if match == -1 else (_order_key(line[i]) > k))]
    if not candidates:
        return None
    pick = max if match == -1 else min
    best = pick(_order_key(line[i]) for i in candidates)
    return next(i for i in candidates if _order_key(line[i]) == best)


def _day(serial: float) -> date:
    if serial < 0:
        raise _Fail(Err("#NUM!"))
    return date(1899, 12, 30) + timedelta(days=int(math.floor(serial)))


def _excel_date(year: int, month: int, day: int) -> float:
    """DATUM wie in Excel: Jahre 0 bis 1899 zählen ab 1900, Monate über 12 und Tage über das Monatsende laufen weiter
    (DATUM(2009;1;400) ist der 4.2.2010), und Excel zählt den 29.2.1900 mit, den es nie gab."""
    if 0 <= year < 1900:
        year += 1900
    year, month = year + (month - 1) // 12, (month - 1) % 12 + 1
    if not 1900 <= year <= 9999:
        raise _Fail(Err("#NUM!"))
    first = date(year, month, 1)
    serial = (first - date(1899, 12, 31)).days + (first >= date(1900, 3, 1)) + day - 1
    if not 0 <= serial <= 2958465:
        raise _Fail(Err("#NUM!"))
    return float(serial)


def _serial(year: int, month: int, day: int) -> float:
    """Ein Datum als Excel-Zahl (Tage seit dem 30.12.1899)."""
    try:
        return float((date(year, month, day) - date(1899, 12, 30)).days)
    except ValueError as err:
        raise _Fail(Err("#NUM!")) from err


def _truth(value) -> bool:
    if isinstance(value, str):
        if value.upper() in ("TRUE", "WAHR"):
            return True
        if value.upper() in ("FALSE", "FALSCH"):
            return False
        raise _Fail(Err("#VALUE!"))
    return bool(_num(value))


def _comparable(value):
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, str):
        return value.lower()
    return value


def _text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, Err):
        raise _Fail(value)
    if isinstance(value, bool):
        return "WAHR" if value else "FALSCH"
    if isinstance(value, float):
        return str(int(value)) if value == int(value) else repr(value)
    return str(value)
