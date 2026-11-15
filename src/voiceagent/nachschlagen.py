"""Ursachen beim Nachschlagen, jede mit einer nachgerechneten Korrektur.

„Satz“ in C2 ist =SVERWEIS(A2;Tarife!$A$2:$C$4;3;FALSCH), aber Spalte 3 der Tabelle heißt „Gültig ab“, „Satz“
steht in Spalte 2. Oder SVERWEIS sucht „M“ in Tarife!$A$2:$B$4, und „M“ steht direkt darunter in A5: Der Bereich
hört eine Zeile zu früh auf, Excel zeigt #NV (oder WENNFEHLER macht still 0 daraus). Dazu: das fehlende FALSCH bei
einer unsortierten Tabelle, eine Tabelle ohne $, die beim Herunterziehen mitrutscht, und ein Schlüssel, der nur
wegen eines Leerzeichens oder einer Zahl als Text nicht gefunden wird.

Der Prüfer liest dafür die Spaltenköpfe beider Tabellen und rechnet jede Korrektur auf dem Blatt nach, bevor er sie
vorschlägt. Die Korrektur steht so da, wie man sie in Excel einfügt: deutsche Funktionsnamen, Semikolon, die
$-Zeichen wie im Original. Alles ohne KI.
"""

from __future__ import annotations

import re
from dataclasses import replace

from . import formel
from .analyse import (
    Finding,
    Sheet,
    _approximate,
    _as_number,
    _fmt,
    _german,
    _keys_of,
    _match_type,
    _sorted,
    address,
    calc_value,
)
from .bereiche import _FILLER, _german_formula, _kept, _paths, _put, _text, _words
from .formel import Err, Node
from .sprache import T, english, using
from .workbook import _is_word, _left_label
from .xlsx import column_letters

Pos = tuple[int, int]
Area = tuple[int, int, int, int]

MAX_FINDINGS = 8
# Um die Nachschlageformel darf nur stehen, was den Wert nicht ändert: dann muss ihr Ergebnis zum Spaltenkopf passen.
_WRAPPERS = {"IFERROR", "IFNA", "ROUND", "TRIM", "VALUE"}
_LOOKS = ("VLOOKUP", "HLOOKUP", "XLOOKUP", "INDEX", "MATCH")
_GERMAN_ERRORS = {"#N/A": "#NV", "#VALUE!": "#WERT!", "#REF!": "#BEZUG!", "#NUM!": "#ZAHL!", "#DIV/0!": "#DIV/0!"}
_FAILS = (formel.Unknown, formel._Fail, ValueError, TypeError, ZeroDivisionError, IndexError, RecursionError)  # noqa: SLF001


def _shown(value) -> str:
    if isinstance(value, Err):
        return T(_GERMAN_ERRORS.get(value.code, value.code), value.code)
    if isinstance(value, str):
        return T(f"„{value}“", f"“{value}”") if value else T("nichts (leer)", "nothing (empty)")
    if isinstance(value, bool):
        return T("WAHR" if value else "FALSCH", "TRUE" if value else "FALSE")
    return _fmt(value)


def _name(function: str) -> str:
    """Die Funktion, wie Excel sie in der Sprache der Anfrage nennt: SVERWEIS oder VLOOKUP."""
    return T(_german(function), function)


def _cells(count: int) -> str:
    return T(f"{count} weitere {'Zelle' if count == 1 else 'Zellen'}", f"{count} more {'cell' if count == 1 else 'cells'}")


def _with_fix(sheet: Sheet, finding: Finding, german_tip) -> Finding:
    """Die Korrektur liest einsetzen.fix aus dem deutschen Tipp („Korrigiert: =…“). Auf Englisch steht im Tipp ein
    anderer Satz: Dann baut sie hier der deutsche Tipp, der englische bleibt stehen."""
    if english() and finding.tip and "=" in finding.tip and not finding.fix:
        from . import einsetzen  # noqa: PLC0415 - einsetzen braucht analyse, analyse braucht diese Datei

        with using("de"):
            tip = german_tip()
        finding.fix = einsetzen.fix(sheet, replace(finding, tip=tip))
    return finding


def _same(a, b) -> bool:
    if isinstance(a, Err) or isinstance(b, Err):
        return isinstance(a, Err) and isinstance(b, Err) and a.code == b.code
    return formel._norm(a) == formel._norm(b)  # noqa: SLF001


def _missing(value) -> bool:
    return isinstance(value, Err) and value.code == "#N/A"


def _cell(sheet: Sheet, where: str | None, pos: Pos):
    if where is None or (sheet.name and where.casefold() == sheet.name.casefold()):
        return sheet.cells.get(pos)
    return sheet.book.get(where.casefold(), {}).get(pos)


def _value(sheet: Sheet, where: str | None, pos: Pos):
    if where is None or (sheet.name and where.casefold() == sheet.name.casefold()):
        return sheet.inputs.get(pos)
    try:
        return sheet.other(where, pos)
    except formel.Unknown:
        return None


def _word_at(sheet: Sheet, where: str | None, pos: Pos) -> str:
    cell = _cell(sheet, where, pos)
    if cell is None or cell.formula is not None or not _is_word(cell.value):
        return ""
    return cell.value.strip()


def _header_row(sheet: Sheet, where: str | None, area: Area) -> int | None:
    """Die Zeile mit den Spaltenköpfen der Nachschlagetabelle: direkt darüber, oder die erste Zeile des Bereichs."""
    r1, c1, r2, c2 = area
    for row in (r1 - 1, r1):
        if row < 1 or (row == r1 and r2 == r1):
            continue
        heads = [c for c in range(c1, c2 + 1) if _word_at(sheet, where, (row, c))]
        if len(heads) * 2 >= c2 - c1 + 1 and len(heads) >= 2:
            return row
    return None


def _tokens(text: str) -> frozenset[str]:
    return frozenset(w for w in _words(text) if w not in _FILLER and not w.isdigit())


def _fits(name: str, head: str) -> int:
    """2: gleicher Name, 1: der eine ist im anderen enthalten („Satz“ und „Satz %“), 0: passt nicht."""
    a, b = _tokens(name), _tokens(head)
    if not a or not b:
        return 0
    if a == b:
        return 2
    return 1 if a <= b or b <= a else 0


def _target(sheet: Sheet, pos: Pos) -> str:
    """Wie die Zelle heißt: ihr Spaltenkopf in der Tabelle, sonst das erste Wort über den Formeln und Zahlen
    darüber (zwei Tabellen mit verschiedenen Kopfzeilen auf einem Blatt), sonst das Wort links daneben."""
    r, c = pos
    if sheet.plan.in_table(r, c):
        return sheet.plan.headers.get(c, "")
    row = r - 1
    while row >= 1 and r - row <= 2000:
        cell = sheet.cells.get((row, c))
        if cell is None or cell.value == "":
            break
        if cell.formula is None and _is_word(cell.value):
            return cell.value.strip()
        if cell.formula is None and not isinstance(sheet.inputs.get((row, c)), float):
            break
        row -= 1
    return _left_label(sheet.cells, r, c)


def _bare(node: Node) -> Node | None:
    """Die Nachschlageformel selbst, wenn um sie herum nur WENNFEHLER, RUNDEN und Ähnliches steht."""
    while node.kind == "call" and node.value in _WRAPPERS and node.args:
        node = node.args[0]
    return node if node.kind == "call" and node.value in ("VLOOKUP", "HLOOKUP", "XLOOKUP", "INDEX") else None


def _corners(ref: str, area: Area, fixed: bool) -> str:
    """A2:B4 mit einem neuen Bereich, die $-Zeichen wie vorher (fixed: überall $)."""
    parts = ref.split(":")
    corners = [(area[0], area[1]), (area[2], area[3])]
    out = []
    for part, (row, col) in zip(parts if len(parts) == 2 else [parts[0], parts[0]], corners, strict=True):
        match = re.fullmatch(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)", part)
        if match is None:
            return ""
        dollar_col, dollar_row = ("$", "$") if fixed else (match[1], match[3])
        out.append(f"{dollar_col}{column_letters(col)}{dollar_row}{row}")
    return ":".join(out)


def _restyled(text: str, area: Area, fixed: bool = False) -> str:
    """Tarife!$A$2:$B$4 mit einem neuen Bereich, die $-Zeichen wie vorher: Tarife!$A$2:$B$5."""
    prefix, bang, ref = text.rpartition("!")
    corners = _corners(ref, area, fixed)
    return f"{prefix}{bang}{corners}" if corners else ""


def _area_node(old: Node, area: Area) -> Node:
    if old.kind in ("xrange", "xref"):
        return Node("xrange", (old.value[0], area))
    return Node("range", area)


def _formula(node: Node, kept: dict) -> str:
    english = _text(node, kept)
    return _german_formula(english) if english else ""


def _place(sheet: Sheet, where: str | None) -> str:
    return f"{where}!" if where and not (sheet.name and where.casefold() == sheet.name.casefold()) else ""


def _span(sheet: Sheet, where: str | None, area: Area) -> str:
    inner = f"{address(area[:2])}:{address(area[2:])}" if area[:2] != area[2:] else address(area[:2])
    return f"{_place(sheet, where)}{inner}"


class _Look:
    """Eine Nachschlageformel in einer Zelle: wo sie steht, ihr Aufruf und die Tabelle, in der sie sucht."""

    def __init__(self, sheet: Sheet, calc, pos: Pos, path: tuple, call: Node) -> None:
        self.sheet, self.calc, self.pos, self.path, self.call = sheet, calc, pos, path, call
        self.node = sheet.nodes[pos]

    @property
    def kept(self) -> dict:
        return _kept(self.sheet.cells[self.pos].formula or "")

    def area(self, i: int) -> tuple[str | None, Area]:
        return self.calc.area_of(self.call.args[i])

    def with_args(self, changes: dict[int, Node]) -> Node:
        args = list(self.call.args)
        for i, new in changes.items():
            while i >= len(args):
                args.append(Node("empty", None))
            args[i] = new
        return _put(self.node, self.path, Node("call", self.call.value, tuple(args)))

    def kept_with(self, changes: dict[int, Node]) -> dict:
        """Die $-Zeichen der geänderten Bereiche so, wie sie im Original standen."""
        return self.restyle([(self.call.args[i], new) for i, new in changes.items() if i < len(self.call.args)])

    def restyle(self, pairs: list[tuple[Node, Node]], fixed: bool = False) -> dict:
        kept = dict(self.kept)
        for old, new in pairs:
            if new.kind in ("xrange", "range") and old.kind in ("xrange", "range", "xref", "ref"):
                before = kept.get((old.kind, old.value)) or _text(old, {})
                if before:
                    styled = _restyled(before, new.value[1] if new.kind == "xrange" else new.value, fixed)
                    if styled:
                        kept[(new.kind, new.value)] = styled
        return kept

    def run(self, node: Node):
        try:
            return self.calc.run(node)
        except _FAILS:
            return None

    def key(self):
        try:
            return calc_value(self.calc.eval(self.call.args[0]))
        except _FAILS:
            return None

    def key_place(self) -> str:
        """Wo der gesuchte Wert steht: „A5“, „Daten!A5“ oder „dem Suchwert“."""
        arg = self.call.args[0]
        if arg.kind == "ref":
            return address(arg.value)
        if arg.kind == "xref":
            return f"{_place(self.sheet, arg.value[0])}{address(arg.value[1])}"
        return T("dem Suchwert", "the lookup value")


def _looks(sheet: Sheet) -> list[_Look]:
    calc, _ = sheet.evaluator()
    found = []
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown:
            continue
        for path, part in _paths(node):
            if part.kind == "call" and part.value in _LOOKS and len(part.args) >= 2:
                found.append(_Look(sheet, calc, pos, path, part))
    return found


def _line(look: _Look) -> tuple[str | None, list[Pos]] | None:
    """Die Zellen, in denen die Suche den Schlüssel sucht: die erste Spalte bei SVERWEIS, die erste Zeile bei WVERWEIS."""
    call = look.call
    if call.value == "INDEX":
        return None
    try:
        where, (r1, c1, r2, c2) = look.area(1)
    except formel.Unknown:
        return None
    if call.value == "VLOOKUP" or (call.value != "HLOOKUP" and c1 == c2):
        return where, [(r, c1) for r in range(r1, r2 + 1)]
    if call.value == "HLOOKUP" or r1 == r2:
        return where, [(r1, c) for c in range(c1, c2 + 1)]
    return None


def _exact(look: _Look) -> bool:
    call = look.call
    try:
        if call.value in ("VLOOKUP", "HLOOKUP"):
            return len(call.args) >= 4 and call.args[3].kind != "empty" and not look.calc.eval(call.args[3])
        if call.value == "MATCH":
            return len(call.args) >= 3 and look.calc.eval(call.args[2]) == 0
        if call.value == "XLOOKUP":
            return len(call.args) < 5 or call.args[4].kind == "empty" or look.calc.eval(call.args[4]) == 0
    except _FAILS:
        return False
    return False


# --- Die falsche Spalte ---


def _column_fix(look: _Look) -> tuple[str, str, str, Node, dict] | None:  # noqa: C901 - SVERWEIS, INDEX und XVERWEIS je für sich
    """Holt die Formel eine Spalte, deren Kopf nicht zum Namen der Zelle passt, und gibt es genau eine, die passt?
    Oder zeigt die Spaltennummer hinter den Bereich (#BEZUG!), und dort steht die gemeinte Spalte?
    (Art, Kopf, der genommen wird, Kopf, der gemeint ist, neue Formel, $-Zeichen)."""
    sheet, call = look.sheet, look.call
    if _bare(look.node) is not call:
        return None
    name = _target(sheet, look.pos)
    if not name:
        return None
    single = call.value == "XLOOKUP" or (call.value == "INDEX" and len(call.args) == 2)
    try:
        if single:
            table = 2 if call.value == "XLOOKUP" else 0
            if len(call.args) <= table:
                return None
            where, area = look.area(table)
            r1, c1, r2, c2 = area
            if c1 != c2 or r1 == r2:
                return None
            taken = c1
        elif call.value in ("VLOOKUP", "INDEX"):
            if len(call.args) < 3 or call.args[2].kind != "num" or call.args[2].value < 1:
                return None
            table = 0 if call.value == "INDEX" else 1
            where, area = look.area(table)
            r1, c1, r2, c2 = area
            taken = c1 + int(call.args[2].value) - 1
        else:
            return None
    except formel.Unknown:
        return None
    head_row = _header_row(sheet, where, (r1, min(c1, taken), r2, max(c2, taken)))
    if head_row is None:
        # Eine Spalte für XVERWEIS: die Köpfe stehen über der ganzen Tabelle, nicht nur über dieser Spalte.
        head_row = r1 - 1 if r1 > 1 and _word_at(sheet, where, (r1 - 1, taken)) else None
    if head_row is None:
        return None
    taken_head = _word_at(sheet, where, (head_row, taken))
    # Nur Spalten derselben Tabelle: die Köpfe ohne Lücke neben der Suchspalte, nie die Spalte der Formel selbst.
    columns = _block(sheet, where, head_row, taken if single else c1)
    if taken not in columns:
        return None
    if _place(sheet, where) == "":
        columns = [c for c in columns if c != look.pos[1]]
    if call.value == "VLOOKUP":
        columns = [c for c in columns if c > c1]
    scored = {c: _fits(name, _word_at(sheet, where, (head_row, c))) for c in columns}
    best = max(scored.values(), default=0)
    if not best:
        return None
    if taken > c2 and taken_head and scored.get(taken) == best and list(scored.values()).count(best) == 1:
        # Die Nummer stimmt, nur der Bereich ist zu schmal: Excel zeigt #BEZUG!.
        changes = {table: _area_node(call.args[table], (r1, c1, r2, taken))}
        return "narrow", taken_head, taken_head, look.with_args(changes), look.kept_with(changes)
    if not taken_head or _fits(name, taken_head):
        return None
    meant = [c for c, s in scored.items() if s == best and c != taken]
    if len(meant) != 1:
        return None
    col = meant[0]
    if single:
        changes = {table: _area_node(call.args[table], (r1, col, r2, col))}
    else:
        changes = {2: Node("num", float(col - c1 + 1))}
        if col > c2:
            changes[table] = _area_node(call.args[table], (r1, c1, r2, col))
    return "col", taken_head, _word_at(sheet, where, (head_row, col)), look.with_args(changes), look.kept_with(changes)


def _block(sheet: Sheet, where: str | None, row: int, col: int) -> list[int]:
    """Die Spalten mit Köpfen links und rechts von col, bis zur ersten Lücke."""
    if not _word_at(sheet, where, (row, col)):
        return []
    left = right = col
    while left > 1 and _word_at(sheet, where, (row, left - 1)):
        left -= 1
    while right - left < 200 and _word_at(sheet, where, (row, right + 1)):
        right += 1
    return list(range(left, right + 1))


# --- Der Bereich hört zu früh auf ---


def _block_end(sheet: Sheet, where: str | None, col: int, start: int) -> int:
    """Bis wohin die Einträge in einer Spalte ab start ohne Lücke weitergehen."""
    row = start
    while _value(sheet, where, (row + 1, col)) not in (None, ""):
        row += 1
    return row


def _grown(look: _Look, where: str | None, rows: tuple[int, int], end: int) -> tuple[Node, dict]:
    """Alle Bereiche mit denselben Zeilen wachsen bis end: die Rückgabespalte von XVERWEIS, der Bereich von INDEX."""
    node, pairs = look.node, []
    for path, part in list(_paths(look.node)):
        if part.kind not in ("range", "xrange"):
            continue
        try:
            other, (q1, d1, q2, d2) = look.calc.area_of(part)
        except formel.Unknown:
            continue
        if (other or "").casefold() == (where or "").casefold() and (q1, q2) == rows:
            new = _area_node(part, (q1, d1, end, d2))
            node = _put(node, path, new)
            pairs.append((part, new))
    return node, look.restyle(pairs)


def _range_fix(look: _Look, key) -> tuple[str, Node, dict] | None:
    """Findet die genaue Suche den Schlüssel nicht, obwohl er direkt unter dem Bereich in derselben Spalte steht?
    (wo er steht, neue Formel, $-Zeichen)."""
    sheet, call = look.sheet, look.call
    if call.value not in ("VLOOKUP", "MATCH", "XLOOKUP"):
        return None
    try:
        where, (r1, c1, r2, c2) = look.area(1)
    except formel.Unknown:
        return None
    if (call.value != "VLOOKUP" and c1 != c2) or (r1 == 1 and r2 >= formel.MAX_ROW - 1):
        return None
    # Der Bereich endet mitten in einem Block von Einträgen, und der Schlüssel steht im Rest des Blocks.
    end = _block_end(sheet, where, c1, r2)
    if end == r2 or end - r2 > max(50, r2 - r1 + 1):
        return None
    norm = formel._norm(key)  # noqa: SLF001
    below = [r for r in range(r2 + 1, end + 1) if formel._norm(_value(sheet, where, (r, c1))) == norm]  # noqa: SLF001
    if not below:
        return None
    node, kept = _grown(look, where, (r1, r2), end)
    return address((below[0], c1)), node, kept


# --- Ungefähr statt genau ---


def _flag_text(look: _Look, written: bool) -> str:
    call = look.call
    if call.value == "MATCH":
        if not written:
            return T("das dritte Argument 0 fehlt", "the third argument 0 is missing")
        third = _shown(look.calc.run(call.args[2]))
        return T(f"als drittes Argument {third} steht", f"the third argument is {third}")
    if not written:
        return T("das vierte Argument FALSCH fehlt", "the fourth argument FALSE is missing")
    flag = look.calc.run(call.args[3])
    fourth = T("WAHR", "TRUE") if flag is True else _shown(flag)
    return T(f"als viertes Argument {fourth} steht", f"the fourth argument is {fourth}")


def _exact_fix(look: _Look, cache: dict) -> tuple[str, bool, bool, Node, dict] | None:
    """Sucht die Formel ungefähr, wo sie genau suchen müsste? Bei unsortierter Tabelle immer, bei Text, der fehlt,
    wenn das Argument nur vergessen ist. (Grund, ausdrücklich geschrieben, absteigend, neue Formel, $-Zeichen)."""
    call, calc = look.call, look.calc
    if call.value not in ("VLOOKUP", "HLOOKUP", "MATCH"):
        return None
    rough, written = _approximate(calc, call)
    if not rough:
        return None
    try:
        area = calc.area_key(call.args[1])
    except formel.Unknown:
        return None
    if (call.value, area) not in cache:
        keys = _keys_of(calc, call)
        cache[(call.value, area)] = (keys, {formel._norm(k) for k in keys or []}, {})  # noqa: SLF001
    keys, norms, order = cache[(call.value, area)]
    key = look.key()
    kind = formel._kind(key) if key is not None else None  # noqa: SLF001
    if not keys or not kind:
        return None
    descending = call.value == "MATCH" and _match_type(calc, call) < 0
    if (kind, descending) not in order:
        order[(kind, descending)] = _sorted(keys, kind, descending)
    if not order[(kind, descending)]:
        reason = "unsorted"
    elif not written and kind == "s" and formel._norm(key) not in norms:  # noqa: SLF001
        reason = "missing"
    else:
        return None
    flag = {2: Node("num", 0.0)} if call.value == "MATCH" else {3: Node("bool", False)}
    return reason, written, descending, look.with_args(flag), look.kept


def _taken(look: _Look, now) -> tuple[str | None, Pos] | None:
    """Welche Zelle der Suchspalte die ungefähre Suche getroffen hat. Bei unsortierter Tabelle rechnet der Prüfer
    Excels binäre Suche nicht nach: dann die einzige Zeile, deren Wert die Zelle zeigt."""
    call, calc = look.call, look.calc
    try:
        where, (r1, c1, r2, c2) = look.area(1)
    except formel.Unknown:
        return None
    try:
        if call.value == "MATCH":
            at = int(calc.eval(call)) - 1
            return where, ((r1 + at, c1) if c1 == c2 else (r1, c1 + at))
        calc.hits = []
        calc.eval(call)
        hits = calc.hits
    except _FAILS:
        hits = None
    finally:
        calc.hits = None
    if hits:
        _, sheet, (row, col) = hits[-1]
        return sheet, ((row, c1) if call.value == "VLOOKUP" else (r1, col))
    if call.value == "MATCH" or _bare(look.node) is not call or now is None or isinstance(now, Err):
        return None
    try:
        k = int(calc.eval(call.args[2]))
    except _FAILS:
        return None
    if call.value == "VLOOKUP":
        rows = [r for r in range(r1, r2 + 1) if _same(calc_value(_value(look.sheet, where, (r, c1 + k - 1))), now)]
        return (where, (rows[0], c1)) if len(rows) == 1 else None
    cols = [c for c in range(c1, c2 + 1) if _same(calc_value(_value(look.sheet, where, (r1 + k - 1, c))), now)]
    return (where, (r1, cols[0])) if len(cols) == 1 else None


# --- Leerzeichen, Zahl als Text ---


def _near_fix(look: _Look, values: list) -> tuple[str, str, Node | None, dict, object] | None:
    """Fehlt der Schlüssel nur wegen eines Leerzeichens oder weil er einmal Zahl, einmal Text ist?
    (Art, Zelle in der Tabelle, neue Formel oder None, $-Zeichen, Wert nach dem Bereinigen der Tabelle)."""
    call = look.call
    key = look.key()
    _, cells = _line(look) or (None, [])
    if key is None or isinstance(key, bool) or not cells:
        return None
    arg = call.args[0]
    if isinstance(key, str):
        plain = key.strip().casefold()
        same = [i for i, v in enumerate(values) if isinstance(v, str) and v.strip().casefold() == plain]
        if key != key.strip() and any(values[i].casefold() == plain for i in same):
            i = next(i for i in same if values[i].casefold() == plain)
            return "key-space", address(cells[i]), look.with_args({0: Node("call", "TRIM", (arg,))}), look.kept, None
        if same and key == key.strip():
            i = same[0]
            return "table-space", address(cells[i]), None, look.kept, _row_value(look, i)
        number = _as_number(key)
        hit = [i for i, v in enumerate(values) if isinstance(v, float) and not isinstance(v, bool) and v == number]
        if number is not None and hit:
            return "key-text", address(cells[hit[0]]), look.with_args({0: Node("call", "VALUE", (arg,))}), look.kept, None
        return None
    if isinstance(key, float):
        hit = [i for i, v in enumerate(values) if isinstance(v, str) and _as_number(v) == key]
        if not hit:
            return None
        if any(isinstance(v, float) and not isinstance(v, bool) for v in values):
            # Nur einzelne Schlüssel der Tabelle sind Text: dort die Zahl eintragen, sonst fehlen die übrigen.
            return "table-text-cell", address(cells[hit[0]]), None, look.kept, _row_value(look, hit[0])
        glued = Node("op", "&", (arg, Node("str", "")))
        return "table-text", address(cells[hit[0]]), look.with_args({0: glued}), look.kept, None
    return None


def _row_value(look: _Look, i: int):
    """Was die Suche liefern würde, wenn sie den i-ten Eintrag fände (nur ohne weitere Rechnung um sie herum)."""
    call = look.call
    if look.node is not call and not (look.node.kind == "call" and look.node.value in ("IFERROR", "IFNA")
                                       and look.path == (0,)):
        return None
    try:
        if call.value in ("VLOOKUP", "HLOOKUP"):
            where, (r1, c1, _, _) = look.area(1)
            k = int(look.calc.eval(call.args[2]))
            pos = (r1 + i, c1 + k - 1) if call.value == "VLOOKUP" else (r1 + k - 1, c1 + i)
            return calc_value(_value(look.sheet, where, pos))
        if call.value == "XLOOKUP":
            return look.calc._line(call.args[2])[i]  # noqa: SLF001
        if call.value == "MATCH":
            return float(i + 1)
    except _FAILS:
        return None
    return None


# --- Die Tabelle rutscht mit ---


def _slides(looks: list[_Look]) -> list[tuple]:
    """Gruppen gleicher Formeln untereinander, deren Tabelle ohne $ beim Herunterziehen mitrutscht."""
    groups: dict[tuple, list[_Look]] = {}
    for look in looks:
        groups.setdefault((look.pos[1], look.call.value, look.path), []).append(look)
    found = []
    for (_, _, path), items in groups.items():
        if len(items) < 3:
            continue
        first = items[0]
        table = 0 if first.call.value == "INDEX" else 1
        moving: list[tuple] = []
        for sub, part in _paths(first.node):
            if part.kind not in ("range", "xrange"):
                continue
            try:
                where, area = first.calc.area_of(part)
            except formel.Unknown:
                continue
            if area[0] == area[2]:
                continue  # Eine Zeile rutscht mit Absicht mit: INDEX(B2:M2; …) sucht in der eigenen Zeile.
            ok = True
            for look in items[1:]:
                shift = look.pos[0] - first.pos[0]
                try:
                    there = _at(look.node, sub)
                    other, (q1, d1, q2, d2) = look.calc.area_of(there)
                except (formel.Unknown, IndexError, AttributeError):
                    ok = False
                    break
                if (other or "").casefold() != (where or "").casefold() or (q1, d1, q2, d2) != (
                        area[0] + shift, area[1], area[2] + shift, area[3]):
                    ok = False
                    break
            if ok:
                moving.append((sub, part, where, area))
        table_path = path + (table,)
        if any(sub == table_path for sub, *_ in moving):
            found.append((items, moving))
    return found


def _at(node: Node, path: tuple) -> Node:
    for i in path:
        node = node.args[i]
    return node


def _anchored(items: list[_Look], moving: list[tuple]) -> tuple[list, list, int | None] | None:
    """Jede Formel mit der festen Tabelle der ersten Zelle nachgerechnet: [(Zelle, jetzt, danach)], die neuen
    Bereiche und, falls die feste Tabelle zu kurz ist, bis wohin sie reichen muss."""
    first = items[0]

    def fixed(areas: dict) -> list:
        out = []
        for look in items:
            node = look.node
            for sub, part, _, _ in moving:
                node = _put(node, sub, _area_node(part, areas[sub]))
            out.append((look, look.run(look.node), look.run(node)))
        return out

    def worse(results: list) -> bool:
        return any(after is None or (isinstance(after, Err) and not isinstance(now, Err)) for _, now, after in results)

    areas = {sub: area for sub, _, _, area in moving}
    results = fixed(areas)
    end = None
    if worse(results):
        # Die oberen Einträge sind wieder da, dafür fehlen unten welche: die feste Tabelle muss bis zum Ende reichen.
        table = 0 if first.call.value == "INDEX" else 1
        _, _, where, (r1, c1, r2, _) = next(m for m in moving if m[0] == first.path + (table,))
        end = _block_end(first.sheet, where, c1, r2)
        if end == r2:
            return None
        areas = {sub: (a[0], a[1], end, a[3]) if (a[0], a[2]) == (r1, r2) else a for sub, _, _, a in moving}
        results = fixed(areas)
        if worse(results):
            return None
    return results, [(part, _area_node(part, areas[sub])) for sub, part, _, _ in moving], end


# --- Hinweise ---


def _now_text(look: _Look, now) -> str:
    """Was Excel in der Zelle zeigt (ein Datum als Datum), nicht die Zahl dahinter."""
    cell = look.sheet.cells[look.pos]
    shown = (cell.value or "").strip()
    if shown and isinstance(now, str):
        return T(f"„{cell.value}“", f"“{cell.value}”")
    if shown and isinstance(now, float) and shown != _fmt(now) and not re.fullmatch(r"-?[\d.,]+", shown):
        return shown
    return _shown(now)


def _more(count: int) -> str:
    return T(f" Genauso in {count} weiteren {'Zelle' if count == 1 else 'Zellen'} darunter.",
             f" The same in {_cells(count)} below.") if count else ""


def _drag(count: int) -> str:
    return T(" Danach wieder bis unten herunterziehen.", " Then fill down to the bottom again.") if count else ""


def _put_all(look: _Look, pairs: list[tuple[Node, Node]]) -> Node:
    """Die Formel der Zelle mit den neuen Bereichen an den Stellen der alten."""
    node = look.node
    for old, new in pairs:
        for path, part in list(_paths(node)):
            if part is old:
                node = _put(node, path, new)
                break
    return node


def _sliding(sheet: Sheet, looks: list[_Look]) -> tuple[list[Finding], set[Pos], set[str]]:
    """Die mitrutschende Tabelle: Sie erklärt, warum weiter unten Schlüssel fehlen."""
    found, done, covered = [], set(), set()
    for items, moving in _slides(looks):
        fixed = _anchored(items, moving)
        if fixed is None:
            continue
        results, pairs, end = fixed
        changed = [(look, now, after) for look, now, after in results if not _same(now, after)]
        if not changed:
            continue
        first, (look, now, after) = items[0], changed[0]
        text = _formula(_put_all(first, pairs), first.restyle(pairs, fixed=True))
        if not text:
            continue
        name = _name(first.call.value)
        table = 0 if first.call.value == "INDEX" else 1
        where, start = first.area(table)
        _, there = look.area(table)
        longer = T(f" Außerdem muss sie bis Zeile {end} gehen, so weit reichen die Einträge.",
                   f" It also has to go down to row {end}, that is how far the entries reach.") if end else ""
        rest = len(changed) - 1
        others = T(f" Genauso falsch: {rest} weitere {'Zelle' if rest == 1 else 'Zellen'}.",
                   f" Also wrong: {_cells(rest)}.") if rest else ""

        def tip(first=first, text=text, items=items, look=look, after=after) -> str:
            return T(f"Korrigiert in {address(first.pos)}: ={text}, dann bis {address(items[-1].pos)} herunterziehen. "
                     f"Nachgerechnet zeigt {address(look.pos)} dann {_shown(after)}.",
                     f"Corrected in {address(first.pos)}: ={text}, then fill down to {address(items[-1].pos)}. "
                     f"Recalculated, {address(look.pos)} then shows {_shown(after)}.")

        found.append(_with_fix(sheet, Finding(
            f"lookslide:{address(first.pos)}",
            T(f"Die Tabelle, in der {name} in {address(first.pos)} sucht, hat keine $-Zeichen und rutscht beim "
              f"Herunterziehen mit: {address(look.pos)} sucht schon in {_span(sheet, where, there)} statt in "
              f"{_span(sheet, where, start)}. Darum zeigt {address(look.pos)} {_now_text(look, now)} statt "
              f"{_shown(after)}.{longer}{others}",
              f"The table that {name} in {address(first.pos)} searches has no $ signs and shifts when filled "
              f"down: {address(look.pos)} already searches {_span(sheet, where, there)} instead of "
              f"{_span(sheet, where, start)}. That is why {address(look.pos)} shows {_now_text(look, now)} instead of "
              f"{_shown(after)}.{longer}{others}"),
            "fehler",
            tip(),
        ), tip))
        done |= {item.pos for item in items}
        covered |= {f"{kind}:{address(item.pos)}" for item in items for kind in ("sliding", "hidden", "approx", "nearest")}
    return found, done, covered


def _collect(sheet: Sheet, looks: list[_Look], done: set[Pos]) -> dict[tuple, list]:  # noqa: C901 - eine Ursache je Abschnitt
    """Jede Nachschlageformel mit einer nachgerechneten Ursache, gruppiert nach Spalte und Art."""
    groups: dict[tuple, list] = {}
    seen: dict = {}
    orders: dict = {}
    for look in looks:
        if look.pos in done:
            continue
        column = _column_fix(look)
        if column is not None:
            kind, taken, meant, node, kept = column
            now, after = look.run(look.node), look.run(node)
            if after is not None and not isinstance(after, Err) and not _same(after, now):
                groups.setdefault((kind, look.pos[1], taken, meant), []).append((look, now, after, node, kept))
                continue
        exact = _exact_fix(look, orders)
        if exact is not None:
            reason, written, descending, node, kept = exact
            # Bei unsortierter Tabelle rechnet der Prüfer die ungefähre Suche nicht nach: Es zählt, was Excel zeigt.
            now, after = look.run(look.node), look.run(node)
            if now is None:
                now = calc_value(sheet.shown.get(look.pos))
            if after is not None and now not in (None, ""):
                area = look.calc.area_key(look.call.args[1])
                groups.setdefault(("exact", look.pos[1], look.call.value, area), []).append(
                    (look, now, after, node, kept, reason, written, descending))
            continue
        if not _exact(look):
            continue
        key = look.key()
        line = _line(look)
        if key is None or key == "" or isinstance(key, bool) or line is None:
            continue
        where, cells = line
        column_key = ((where or "").casefold(), cells[0], cells[-1])
        if column_key not in seen:
            values = [calc_value(_value(sheet, where, pos)) for pos in cells]
            seen[column_key] = (values, {formel._norm(v) for v in values})  # noqa: SLF001
        values, norms = seen[column_key]
        if formel._norm(key) in norms or not _missing(look.run(look.call)):  # noqa: SLF001
            continue
        short = _range_fix(look, key)
        if short is not None:
            at, node, kept = short
            now, after = look.run(look.node), look.run(node)
            if after is not None and not isinstance(after, Err):
                groups.setdefault(("short", look.pos[1], look.area(1)), []).append((look, now, after, node, kept, key, at))
                continue
        near = _near_fix(look, values)
        if near is not None:
            kind, at, node, kept, value = near
            now = look.run(look.node)
            after = look.run(node) if node is not None else value
            if after is not None and not isinstance(after, Err) and not _same(after, now):
                groups.setdefault(("near", look.pos[1], look.call.value, kind), []).append(
                    (look, now, after, node, kept, key, at, kind))
    return groups


def _column_text(sheet: Sheet, group: tuple, items: list) -> tuple[str, str, str]:
    kind, _, taken, meant = group
    look, now, after = items[0][:3]
    name, cell, rest = _name(look.call.value), address(look.pos), len(items) - 1
    index = {"INDEX": 0, "XLOOKUP": 2}.get(look.call.value, 1)
    where, area = look.area(index)
    table = _span(sheet, where, area)
    label = _target(sheet, look.pos)
    if kind == "narrow":
        width = area[3] - area[1] + 1
        text = T(f"„{label}“ in {cell} holt mit {name} Spalte {int(look.call.args[2].value)} aus {table}, aber der Bereich "
                 f"hat nur {width} {'Spalte' if area[3] == area[1] else 'Spalten'}. Die Spalte „{meant}“ liegt dahinter, darum ergibt {cell} "
                 f"{_now_text(look, now)} statt {_shown(after)}.{_more(rest)}",
                 f"“{label}” in {cell} uses {name} to get column {int(look.call.args[2].value)} from {table}, but the range "
                 f"has only {width} {'column' if area[3] == area[1] else 'columns'}. The column “{meant}” lies beyond it, so {cell} "
                 f"gives {_now_text(look, now)} instead of {_shown(after)}.{_more(rest)}")
    else:
        numbered = look.call.value == "VLOOKUP" or (look.call.value == "INDEX" and len(look.call.args) > 2)
        if numbered:
            how = T(f"holt mit {name} Spalte {int(look.call.args[2].value)} aus {table}",
                    f"uses {name} to get column {int(look.call.args[2].value)} from {table}")
        else:
            how = T(f"holt mit {name} die Werte aus {table}", f"uses {name} to get the values from {table}")
        text = T(f"„{label}“ in {cell} {how}, und diese Spalte heißt „{taken}“. Gemeint ist die Spalte „{meant}“. "
                 f"So zeigt {cell} {_now_text(look, now)} statt {_shown(after)}.{_more(rest)}",
                 f"“{label}” in {cell} {how}, and that column is called “{taken}”. The column meant is “{meant}”. "
                 f"So {cell} shows {_now_text(look, now)} instead of {_shown(after)}.{_more(rest)}")
    return f"lookcol:{cell}", text, ""


def _short_text(sheet: Sheet, items: list) -> tuple[str, str, str]:
    look, now, after = items[0][:3]
    *_, key, at = items[0]
    name, cell = _name(look.call.value), address(look.pos)
    where, area = look.area(1)
    text = T(f"{name} in {cell} sucht {_shown(key)} in {_span(sheet, where, area)}, aber {_shown(key)} steht in "
             f"{_place(sheet, where)}{at}, unter dem Bereich. Der Bereich hört zu früh auf, darum ergibt "
             f"{cell} {_now_text(look, now)} statt {_shown(after)}.{_more(len(items) - 1)}",
             f"{name} in {cell} looks for {_shown(key)} in {_span(sheet, where, area)}, but {_shown(key)} is in "
             f"{_place(sheet, where)}{at}, below the range. The range ends too early, so "
             f"{cell} gives {_now_text(look, now)} instead of {_shown(after)}.{_more(len(items) - 1)}")
    return f"lookshort:{cell}", text, ""


def _exact_text(sheet: Sheet, item: tuple, others: int) -> tuple[str, str, str]:
    look, now, after, _, _, reason, written, descending = item
    name, cell = _name(look.call.value), address(look.pos)
    where, area = look.area(1)
    table = _span(sheet, where, area)
    key = look.key()
    hit = _taken(look, now)
    took = ""
    if hit is not None:
        row = f"{_shown(calc_value(_value(sheet, hit[0], hit[1])))} ({_place(sheet, hit[0])}{address(hit[1])})"
        took = T(f" die Zeile von {row}", f" the row of {row}")
    flag = _flag_text(look, written)
    if reason == "unsorted":
        order = T("absteigend" if descending else "aufsteigend", "descending" if descending else "ascending")
        text = T(f"{name} in {cell} sucht ungefähr, weil {flag}, und {table} ist nicht {order} sortiert. Darum nimmt "
                 f"{cell}{took or ' einen falschen Eintrag'} und zeigt {_now_text(look, now)}, richtig ist {_shown(after)}.",
                 f"{name} in {cell} looks for an approximate match because {flag}, and {table} is not sorted in {order} "
                 f"order. So {cell} takes{took or ' a wrong entry'} and shows {_now_text(look, now)}; correct is "
                 f"{_shown(after)}.")
    else:
        text = T(f"{name} in {cell} findet {_shown(key)} nicht in {table} und nimmt still{took or ' einen anderen Eintrag'}, "
                 f"weil {flag}. So zeigt {cell} {_now_text(look, now)}",
                 f"{name} in {cell} does not find {_shown(key)} in {table} and silently takes{took or ' another entry'} "
                 f"because {flag}. So {cell} shows {_now_text(look, now)}")
        if isinstance(after, Err):
            text += T(f". Mit genauer Suche zeigt Excel {_shown(after)}, solange {_shown(key)} in der Tabelle fehlt.",
                      f". With an exact match, Excel shows {_shown(after)} as long as {_shown(key)} is missing from the table.")
        else:
            text += T(f" statt {_shown(after)}.", f" instead of {_shown(after)}.")
    if others:
        text += T(f" Genauso falsch: {others} weitere {'Zelle' if others == 1 else 'Zellen'} in dieser Spalte.",
                  f" Also wrong: {_cells(others)} in this column.")
    return f"lookexact:{cell}", text, ""


def _near_text(sheet: Sheet, items: list, fixed: str) -> tuple[str, str, str]:
    look, now, after = items[0][:3]
    *_, key, at, near = items[0]
    name, cell, rest = _name(look.call.value), address(look.pos), len(items) - 1
    where, area = look.area(1)
    table, place, source = _span(sheet, where, area), f"{_place(sheet, where)}{at}", look.key_place()
    because = T(f"Darum zeigt {cell} {_now_text(look, now)} statt {_shown(after)}.{_more(rest)}",
                f"That is why {cell} shows {_now_text(look, now)} instead of {_shown(after)}.{_more(rest)}")
    corrected = T(f"Korrigiert: ={fixed}, nachgerechnet {_shown(after)}.{_drag(rest)}",
                  f"Corrected: ={fixed}, recalculated {_shown(after)}.{_drag(rest)}")
    missing = T(f"{name} in {cell} findet {_shown(key)} nicht in {table}", f"{name} in {cell} does not find {_shown(key)} in {table}")
    if near == "key-space":
        start = key != key.lstrip()
        side = T("am Anfang" if start else "am Ende", "at the start" if start else "at the end")
        text = T(f"{missing}, weil in {source} ein Leerzeichen {side} steht. "
                 f"In der Tabelle steht {_shown(key.strip())} ({place}). {because}",
                 f"{missing} because {source} has a space {side}. "
                 f"The table has {_shown(key.strip())} ({place}). {because}")
        tip = T(f"{corrected} Oder lösch das Leerzeichen in {source}.", f"{corrected} Or delete the space in {source}.")
    elif near == "table-space":
        raw = _value(sheet, where, formel.position(at))
        start = isinstance(raw, str) and raw != raw.lstrip()
        side = T("am Anfang" if start else "am Ende", "at the start" if start else "at the end")
        text = T(f"{missing}, weil in {place} ein Leerzeichen {side} steht ({_shown(raw)}). {because}",
                 f"{missing} because {place} has a space {side} ({_shown(raw)}). {because}")
        tip = T(f"Lösch das Leerzeichen in {_places(sheet, where, items)}, dann zeigt {cell} {_shown(after)} (nachgerechnet).",
                f"Delete the space in {_places(sheet, where, items)}, then {cell} shows {_shown(after)} (recalculated).")
    elif near == "key-text":
        text = T(f"{missing}, weil in {source} die Zahl als Text steht, in der Tabelle aber als Zahl ({place}). {because}",
                 f"{missing} because {source} has the number as text, but the table has it as a number ({place}). {because}")
        tip = T(f"{corrected} Oder trag in {source} die Zahl ohne Apostroph ein.",
                f"{corrected} Or enter the number in {source} without an apostrophe.")
    elif near == "table-text-cell":
        text = T(f"{name} in {cell} findet {_fmt(key)} nicht in {table}, weil in {place} die Zahl als Text steht, in "
                 f"{source} aber als Zahl. {because}",
                 f"{name} in {cell} does not find {_fmt(key)} in {table} because {place} has the number as text, but "
                 f"{source} has it as a number. {because}")
        tip = T(f"Trag in {_places(sheet, where, items)} die Zahl ohne Apostroph ein (oder auf das grüne Dreieck klicken "
                f"→ In eine Zahl umwandeln), dann zeigt {cell} {_shown(after)} (nachgerechnet).",
                f"Enter the number in {_places(sheet, where, items)} without an apostrophe (or click the green triangle "
                f"→ Convert to Number), then {cell} shows {_shown(after)} (recalculated).")
    else:
        text = T(f"{name} in {cell} findet {_fmt(key)} nicht in {table}, weil die Tabelle die Zahlen als Text enthält "
                 f"({place}), {source} aber als Zahl. {because}",
                 f"{name} in {cell} does not find {_fmt(key)} in {table} because the table holds the numbers as text "
                 f"({place}), but {source} has a number. {because}")
        tip = T(f"{corrected} Oder wandle die Spalte der Tabelle in Zahlen um (Daten → Text in Spalten → Fertig stellen).",
                f"{corrected} Or convert the table column to numbers (Data → Text to Columns → Finish).")
    return f"looknear:{cell}", text, tip


def _places(sheet: Sheet, where: str | None, items: list) -> str:
    """Die Zellen der Tabelle, die alle Formeln der Gruppe stören: „Tarife!A3“ oder „Tarife!A3 und Tarife!A7“."""
    places = list(dict.fromkeys(f"{_place(sheet, where)}{item[6]}" for item in items))
    return places[0] if len(places) == 1 else f"{', '.join(places[:-1])}{T(' und ', ' and ')}{places[-1]}"


def _said(sheet: Sheet, group: tuple, items: list, every: list, fixed: str) -> tuple[str, str, str]:
    """Schlüssel, Text und Tipp eines Hinweises."""
    kind = group[0]
    look, _, after = items[0][:3]
    tip = T(f"Korrigiert: ={fixed}, nachgerechnet {_shown(after)}.{_drag(len(every) - 1)}",
            f"Corrected: ={fixed}, recalculated {_shown(after)}.{_drag(len(every) - 1)}")
    if kind in ("col", "narrow"):
        key, text, _ = _column_text(sheet, group, items)
    elif kind == "short":
        key, text, _ = _short_text(sheet, items)
    elif kind == "exact":
        key, text, _ = _exact_text(sheet, items[0], len(items) - 1)
        if isinstance(after, Err):
            tip = T(f"Korrigiert: ={fixed}. Nachgerechnet zeigt {address(look.pos)} dann {_shown(after)}, bis "
                    f"{_shown(look.key())} in der Tabelle steht.{_drag(len(every) - 1)}",
                    f"Corrected: ={fixed}. Recalculated, {address(look.pos)} then shows {_shown(after)} until "
                    f"{_shown(look.key())} is in the table.{_drag(len(every) - 1)}")
    else:
        key, text, tip = _near_text(sheet, items, fixed)
    return key, text, tip


def checks(sheet: Sheet) -> tuple[list[Finding], set[str], set[Pos]]:
    """Die Hinweise, die Schlüssel älterer Hinweise (analyse.py), die sie genauer sagen, und die Zellen, deren
    Ursache sie nennen."""
    looks = _looks(sheet)
    found, done, covered = _sliding(sheet, looks)
    explained = set(done)
    for group, items in _collect(sheet, looks, done).items():
        if len(found) >= MAX_FINDINGS:
            break
        kind, every = group[0], items
        if kind == "exact":
            items = [item for item in items if not _same(item[1], item[2])]
            if not items:
                continue
        look, _, after, node, kept = items[0][:5]
        fixed = _formula(node, kept) if node is not None else ""
        if node is not None and not fixed:
            continue
        if kind == "narrow":
            covered.add(f"error:{address(look.pos)}")
        elif kind in ("short", "near"):
            covered |= {f"hidden:{address(item[0].pos)}" for item in items}
        elif kind == "exact":
            covered |= {f"{old}:{address(item[0].pos)}" for item in every for old in ("approx", "nearest")}
        key, text, tip = _said(sheet, group, items, every, fixed)
        found.append(_with_fix(sheet, Finding(key, text, "fehler", tip),
                               lambda group=group, items=items, every=every, fixed=fixed:
                               _said(sheet, group, items, every, fixed)[2]))
        explained |= {item[0].pos for item in every}
    return found, covered, explained
