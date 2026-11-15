"""Excel-Tabellen und Namen: Gegenprobe rechnet =[@Menge]*[@Preis] wie =C5*D5 und =(B5-C5)*MwSt wie =(B5-C5)*$I$1.

Excel schreibt Formeln in Tabellen (Einfügen → Tabelle) mit strukturierten Bezügen: [@Menge] ist die Zelle der Spalte
Menge in derselben Zeile, Verkauf[Umsatz] die ganze Spalte, Verkauf[[#Totals],[Umsatz]] ihre Ergebniszeile. Dazu
kommen Namen für Zellen, Bereiche und Werte (Formeln → Namensmanager: MwSt → Daten!$I$1). Gegenprobe prüft Formeln
Zelle für Zelle; dafür löst resolve() beim Lesen jeden solchen Bezug in die Zellen auf, die er meint. Die Formel, wie
sie in Excel steht, bleibt in Cell.typed. Hinweise zeigen sie so, und Korrekturen schreibt restyle() wieder im Stil der
Mappe (=[@Menge]*[@Preis] statt =C5*D5): Sonst markiert Excel die berechnete Spalte als uneinheitlich.

Was sich nicht sicher auflösen lässt (ein Name mit relativen Bezügen, eine andere Mappe, eine Formel mit LET), bleibt,
wie es ist: Dann rechnet die Prüfung diese Zelle nicht nach, wie zuvor auch.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field, replace

from .sprache import T
from .xlsx import MAX_FORMULA, Area, Cell, Sheet, column_letters, column_number

Pos = tuple[int, int]

# So tief dürfen Namen auf Namen verweisen (Brutto → Netto*Satz → Satz → Daten!$I$1).
MAX_DEPTH = 4
MAX_NAMES = 2000

_CELL = re.compile(r"\$?([A-Za-z]{1,3})\$?(\d{1,7})(?![\w.(\[!])")
_WHOLE = re.compile(r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![\w.(\[!])|\$?\d{1,7}:\$?\d{1,7}(?![\w.])")
_NUMBER = re.compile(r"\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?")
_ERROR = re.compile(r"#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|"
                    r"CONNECT!|BUSY!|UNKNOWN!)", re.I)
_IDENT = re.compile(r"(?:[^\W\d]|[_\\])[\w.\\?]*")
_SHEET_REF = re.compile(r"((?:'(?:[^']|'')+'|[^\W\d][\w.]*)!)?(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?)")
# Eine Zelle oder ein Bereich in einer fertigen Formel (für restyle): nicht mitten in einem Namen oder Tabellenbezug.
_A1 = re.compile(r"(?<![\w.$!'\]\[@])((?:'(?:[^']|'')+'|[^\W\d][\w.]*)!)?(\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?)"
                 r"(?![\w(!\[])")
_SPECIALS = {"#all": "all", "#data": "data", "#headers": "headers", "#totals": "totals", "#this row": "row"}
_ROW = "\x00"  # Platzhalter für die Zeile der Zelle in einer gemerkten Formel


def _key(name: str) -> str:
    return " ".join(name.split()).casefold()


@dataclass(frozen=True)
class Table:
    name: str
    sheet: str
    area: Area
    columns: tuple[str, ...]
    header: bool = True
    totals: int = 0

    @property
    def top(self) -> int:
        return self.area[0] + (1 if self.header else 0)

    @property
    def bottom(self) -> int:
        return self.area[2] - self.totals

    def column(self, name: str) -> int | None:
        wanted = _key(name)
        for i, col in enumerate(self.columns):
            if _key(col) == wanted:
                return self.area[1] + i
        return None

    def holds(self, pos: Pos) -> bool:
        return self.top <= pos[0] <= self.bottom and self.area[1] <= pos[1] <= self.area[3]


@dataclass
class Style:
    """Wie ein Blatt seine Formeln schreibt, für restyle(): Tabellen, deren Zeilen mit [@Spalte] rechnen, Tabellen,
    deren Spalten es als Verkauf[Umsatz] liest, und die Namen, die es benutzt (Name → aufgelöster Bezug)."""

    rows: list[Table] = field(default_factory=list)
    columns: list[Table] = field(default_factory=list)
    names: dict[str, str] = field(default_factory=dict)


def _quote(sheet: str) -> str:
    if re.fullmatch(r"[^\W\d][\w.]*", sheet) and not _CELL.fullmatch(sheet):
        return sheet
    return "'" + sheet.replace("'", "''") + "'"


def _bracket(text: str, at: int) -> int | None:
    """Das Ende (nach „]“) des Klammerausdrucks, der bei text[at] == „[“ beginnt; „'“ hebt das nächste Zeichen auf."""
    depth, i = 0, at
    while i < len(text):
        ch = text[i]
        if ch == "'":
            i += 2
            continue
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                return i + 1
        i += 1
    return None


def _unescape(text: str) -> str:
    return re.sub(r"'(.)", r"\1", text.strip())


def _items(inner: str) -> list[str] | None:
    """„[#This Row],[Menge]“ → [„#This Row“, „Menge“]; „[Von]:[Bis]“ bleibt ein Teil „Von]:[Bis“ (siehe _columns)."""
    parts, i = [], 0
    while i < len(inner):
        while i < len(inner) and inner[i] in " ,":
            i += 1
        if i >= len(inner):
            break
        if inner[i] != "[":
            return None
        end = _bracket(inner, i)
        if end is None:
            return None
        rest = inner[end:].lstrip()
        if rest.startswith(":"):
            # [Von]:[Bis]
            j = inner.index("[", end)
            stop = _bracket(inner, j)
            if stop is None:
                return None
            parts.append((inner[i + 1:end - 1], inner[j + 1:stop - 1]))
            i = stop
        else:
            parts.append(inner[i + 1:end - 1])
            i = end
    return parts


def _spec(inner: str) -> tuple[set[str], tuple[str, str] | None] | None:
    """Was ein strukturierter Bezug meint: welche Zeilen (#Data, #This Row …) und welche Spalten (von, bis)."""
    inner = inner.strip()
    rows: set[str] = set()
    columns: tuple[str, str] | None = None
    if inner.startswith("@"):
        rows.add("row")
        inner = inner[1:].strip()
        if not inner:
            return rows, None
        if not inner.startswith("["):
            name = _unescape(inner)
            return rows, (name, name)
    elif inner.startswith("#"):
        special = _SPECIALS.get(" ".join(inner.split()).casefold())
        return ({special}, None) if special else None
    elif not inner.startswith("["):
        name = _unescape(inner)
        return (rows, (name, name)) if name else None
    items = _items(inner)
    if not items:
        return None
    for item in items:
        if isinstance(item, tuple):
            if columns is not None:
                return None
            columns = (_unescape(item[0]), _unescape(item[1]))
        elif item.strip().startswith("#"):
            special = _SPECIALS.get(" ".join(item.split()).casefold())
            if special is None:
                return None
            rows.add(special)
        else:
            if columns is not None:
                return None
            name = _unescape(item)
            columns = (name, name)
    return rows, columns


def _rows(table: Table, kinds: set[str], row: int | None) -> tuple[int, int] | None:
    top, bottom = table.top, table.bottom
    header, totals = table.area[0], table.area[2]
    if not kinds or kinds == {"data"}:
        return (top, bottom) if top <= bottom else None
    if kinds == {"row"}:
        return (row, row) if row is not None and top <= row <= bottom else None
    if kinds == {"all"}:
        return table.area[0], table.area[2]
    if kinds == {"headers"}:
        return (header, header) if table.header else None
    if kinds == {"totals"}:
        return (totals, totals) if table.totals else None
    if kinds == {"headers", "data"}:
        return (header if table.header else top), bottom
    if kinds == {"data", "totals"}:
        return top, (totals if table.totals else bottom)
    return None


def _display_column(name: str) -> str:
    """So schreibt Excel eine Spalte nach @: [@Menge], aber [@[Preis je Stück]] und Sonderzeichen mit „'“."""
    escaped = re.sub(r"([\[\]#'])", r"'\1", name)
    return escaped if re.fullmatch(r"\w+", name) else f"[{escaped}]"


class _Resolver:
    def __init__(self, sheets: list[Sheet], names: dict[str, str], local: dict[str, dict[str, str]]) -> None:
        self.tables: dict[str, Table] = {}
        for sheet in sheets:
            for t in sheet.tables:
                columns = tuple(t.columns)
                if len(columns) != t.area[3] - t.area[1] + 1:
                    if not t.header:
                        continue
                    # Ohne Spaltennamen (ältere Fenster): Die Kopfzeile einer Tabelle sind ihre Spaltennamen.
                    columns = tuple((sheet.cells.get((t.area[0], c)) or Cell("")).value
                                    for c in range(t.area[1], t.area[3] + 1))
                self.tables[_key(t.name)] = Table(t.name, sheet.name, t.area, columns, t.header, t.totals)
        self.names = {_key(n): (n, text) for n, text in list(names.items())[:MAX_NAMES]}
        self.local = {_key(s): {_key(n): (n, text) for n, text in list(items.items())[:MAX_NAMES]}
                      for s, items in local.items()}
        words = {t.name for t in self.tables.values()} | {n for n, _ in self.names.values()}
        words |= {n for items in self.local.values() for n, _ in items.values()}
        self.words = re.compile("|".join(re.escape(w) for w in sorted(words, key=len, reverse=True)), re.I) if words else None
        self.memo: dict = {}
        self.named: dict = {}

    def table_at(self, sheet: str, pos: Pos) -> Table | None:
        for t in self.tables.values():
            if t.sheet == sheet and t.area[0] <= pos[0] <= t.area[2] and t.area[1] <= pos[1] <= t.area[3]:
                return t
        return None

    def wanted(self, formula: str) -> bool:
        return "[" in formula or (self.words is not None and self.words.search(formula) is not None)

    # --- eine Formel ---

    def formula(self, text: str, sheet: str, pos: Pos | None, depth: int = 0, used: Style | None = None,
                view: str | None = "") -> tuple[str, str] | None:
        """(aufgelöst, wie Excel sie zeigt) oder None, wenn sich etwas darin nicht sicher auflösen lässt. Für [@Spalte]
        steht _ROW statt der Zeile darin. sheet: das Blatt der Formel (für Namen, die nur dort gelten); view: von
        welchem Blatt aus Bezüge ohne Blattnamen geschrieben werden ("" heißt sheet, None: immer mit Blattnamen)."""
        view = sheet if view == "" else view
        names_ok = not re.search(r"\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(", text, re.I)
        own = self.table_at(sheet, pos) if pos is not None else None
        out: list[str] = []
        shown: list[str] = []

        def keep(piece: str) -> None:
            out.append(piece)
            shown.append(piece)

        i, n = 0, len(text)
        changed = False
        while i < n:
            ch = text[i]
            if ch == '"' or ch == "'":
                j = i + 1
                while j < n:
                    if text[j] == ch:
                        if j + 1 < n and text[j + 1] == ch:
                            j += 2
                            continue
                        break
                    j += 1
                if ch == "'" and j + 1 < n and text[j + 1] == "!":
                    done = self._after_sheet(text, j + 2, text[i:j + 2], text[i + 1:j].replace("''", "'"), view, depth,
                                             names_ok, used)
                    if done is None:
                        return None
                    out.append(done[0])
                    shown.append(done[1])
                    i, changed = done[2], changed or done[3]
                    continue
                keep(text[i:j + 1])
                i = j + 1
                continue
            if ch == "[":
                end = _bracket(text, i)
                if end is None:
                    return None
                outer = re.match(r"(?:[^\W\d][\w. ]*)?!", text[end:])
                if outer is not None:
                    # Eine andere Mappe: [1]Blatt!A1 oder [Mappe.xlsx]Blatt!A1. Bleibt, wie es ist.
                    keep(text[i:end + outer.end()])
                    i = end + outer.end()
                    continue
                got = self._structured(own, text[i + 1:end - 1], pos, own, used, view) if own is not None else None
                if got is None:
                    return None
                out.append(got[0])
                shown.append(got[1] or text[i:end])
                changed = True
                i = end
                continue
            if ch == "#":
                match = _ERROR.match(text, i)
                if match is not None:
                    keep(match.group(0))
                    i = match.end()
                    continue
            if ch.isdigit() or (ch == "." and i + 1 < n and text[i + 1].isdigit()):
                match = _WHOLE.match(text, i) or _NUMBER.match(text, i)
                keep(match.group(0))
                i = match.end()
                continue
            if ch == "$" or ch.isalpha() or ch in "_\\":
                match = _WHOLE.match(text, i) or _CELL.match(text, i)
                if match is not None:
                    keep(match.group(0))
                    i = match.end()
                    continue
                match = _IDENT.match(text, i)
                if match is None:
                    keep(ch)
                    i += 1
                    continue
                word, j = match.group(0), match.end()
                rest = text[j:j + 1]
                if rest == "(" or word.upper() in ("TRUE", "FALSE") or word.lower().startswith(("_xlfn.", "_xlpm.", "_xlws.")):
                    keep(word)
                    i = j
                    continue
                if rest == "!":
                    done = self._after_sheet(text, j + 1, word + "!", word, view, depth, names_ok, used)
                    if done is None:
                        return None
                    out.append(done[0])
                    shown.append(done[1])
                    i, changed = done[2], changed or done[3]
                    continue
                table = self.tables.get(_key(word))
                if rest == "[":
                    end = _bracket(text, j)
                    got = self._structured(table, text[j + 1:end - 1], pos, own, used, view) \
                        if table is not None and end is not None else None
                    if got is None:
                        return None
                    out.append(got[0])
                    shown.append(got[1] or text[i:end])
                    changed = True
                    i = end
                    continue
                if table is not None:
                    # Der Tabellenname allein: ihre Daten ohne Kopf- und Ergebniszeile.
                    got = self._structured(table, "#Data", pos, own, used, view)
                    if got is None:
                        return None
                    out.append(got[0])
                    shown.append(word)
                    changed = True
                    i = j
                    continue
                if names_ok and self._known(word, sheet):
                    found = self._name(word, sheet, view, depth, used)
                    if found is None:
                        return None
                    out.append(found)
                    shown.append(word)
                    changed = True
                else:
                    keep(word)
                i = j
                continue
            keep(ch)
            i += 1
        if not changed:
            return text, text
        resolved = "".join(out)
        return (resolved, "".join(shown)) if len(resolved) <= MAX_FORMULA else None

    def _known(self, word: str, sheet: str) -> bool:
        key = _key(word)
        return key in self.names or key in self.local.get(_key(sheet), {})

    def _after_sheet(self, text: str, at: int, prefix: str, target: str, view: str | None, depth: int, names_ok: bool,
                     used: Style | None):
        """Nach „Blatt!“: eine Zelle, ein Bereich oder ein Name, der nur auf diesem Blatt gilt.
        (aufgelöst, wie gezeigt, wo es weitergeht, ob sich etwas geändert hat) oder None."""
        match = _WHOLE.match(text, at) or _CELL.match(text, at)
        if match is not None:
            end = match.end()
            more = re.match(r":\$?[A-Za-z]{1,3}\$?\d{1,7}(?![\w.(\[!])", text[end:])
            if more is not None:
                end += more.end()
            piece = prefix + text[at:end]
            return piece, piece, end, False
        match = _IDENT.match(text, at)
        if match is None:
            return prefix, prefix, at, False
        word = match.group(0)
        if names_ok and _key(word) in self.local.get(_key(target), {}):
            found = self._name(word, target, view, depth, used, only_local=True)
            if found is None:
                return None
            return found, prefix + word, match.end(), True
        return prefix + word, prefix + word, match.end(), False

    def _structured(self, table: Table, inner: str, pos: Pos | None, own: Table | None, used: Style | None,
                    view: str | None) -> tuple[str, str | None] | None:
        """(die Zellen, wie Excel ihn zeigt, oder None: wie geschrieben) oder None, wenn er nichts Sicheres meint."""
        spec = _spec(inner)
        if spec is None:
            return None
        kinds, columns = spec
        if columns is None:
            c1, c2 = table.area[1], table.area[3]
        else:
            c1, c2 = table.column(columns[0]), table.column(columns[1])
            if c1 is None or c2 is None:
                return None
            c1, c2 = min(c1, c2), max(c1, c2)
        this_row = "row" in kinds
        rows = _rows(table, kinds, pos[0] if pos is not None else None)
        if rows is None:
            return None
        r1, r2 = rows
        prefix = "" if view is not None and table.sheet == view else _quote(table.sheet) + "!"
        if this_row:
            if used is not None and own is table and table not in used.rows:
                used.rows.append(table)
            if c1 == c2:
                name = _display_column(table.columns[c1 - table.area[1]])
                return f"{prefix}{column_letters(c1)}{_ROW}", (f"[@{name}]" if own is table else f"{table.name}[@{name}]")
            return f"{prefix}{column_letters(c1)}{_ROW}:{column_letters(c2)}{_ROW}", None
        if used is not None and kinds <= {"data"} and table not in used.columns:
            used.columns.append(table)
        if r1 == r2 and c1 == c2:
            return f"{prefix}${column_letters(c1)}${r1}", None
        return f"{prefix}${column_letters(c1)}${r1}:${column_letters(c2)}${r2}", None

    def _name(self, word: str, scope: str, view: str | None, depth: int, used: Style | None,
              only_local: bool = False) -> str | None:
        """Der Bezug eines Namens, von view aus geschrieben; None, wenn er sich nicht sicher auflösen lässt."""
        key = _key(word)
        local = self.local.get(_key(scope), {}).get(key)
        entry = local if local is not None or only_local else self.names.get(key)
        if entry is None or depth >= MAX_DEPTH:
            return None
        memo = (key, _key(scope) if local is not None else "", _key(view) if view is not None else None)
        if memo not in self.named:
            self.named[memo] = None  # Ein Name, der auf sich selbst verweist, bleibt so ungelöst.
            self.named[memo] = self._target(entry[1], scope if local is not None else (view or scope), view, depth)
        found = self.named[memo]
        if found is not None and used is not None:
            used.names.setdefault(entry[0], found)
        return found

    def _target(self, raw: str, scope: str, view: str | None, depth: int) -> str | None:
        text = raw.strip()
        if text.startswith("="):
            text = text[1:].strip()
        if not text:
            return None
        # Innere Namen und Tabellen immer mit Blattnamen: Der Name gilt auch auf anderen Blättern.
        got = self.formula(text, scope, None, depth + 1, view=None)
        if got is None or _ROW in got[0]:
            return None
        text = got[0]
        plain = re.sub(r'"(?:[^"]|"")*"', "", text)
        if "[" in re.sub(r"'(?:[^']|'')+'", "", plain):
            return None  # eine andere Mappe
        # Nur feste Bezüge ($I$1): Ein relativer Name meint je nach Zelle eine andere.
        plain = re.sub(r"'(?:[^']|'')+'!", "X!", plain)
        for match in re.finditer(r"(?<![\w.])(\$?)[A-Za-z]{1,3}(\$?)\d{1,7}(?![\w.(])", plain):
            if not (match.group(1) and match.group(2)):
                return None
        whole = _SHEET_REF.fullmatch(text)
        if whole is not None:
            prefix = whole.group(1)
            if prefix and view is not None:
                name = prefix[:-1]
                name = name[1:-1].replace("''", "'") if name.startswith("'") else name
                if _key(name) == _key(view):
                    return whole.group(2)
            return text
        if re.fullmatch(r"\d+(?:\.\d+)?(?:[eE][+-]?\d+)?|\"(?:[^\"]|\"\")*\"|TRUE|FALSE|#[A-Z/0!?]+", text, re.I):
            return text
        return f"({text})"

    # --- ein Blatt ---

    def sheet(self, sheet: Sheet) -> None:
        style = Style()
        changes: dict[Pos, Cell] = {}
        for pos, cell in sheet.cells.items():
            formula = cell.formula
            if not formula or not self.wanted(formula):
                continue
            own = self.table_at(sheet.name, pos)
            key = (sheet.name, formula, own.name if own is not None else None)
            done = self.memo.get(key, _MISSING)
            if done is _MISSING or done is _ROWWISE:
                used = Style()
                # In einer Tabelle gilt dieselbe Formel für jede Zeile: einmal mit ihrer ersten Zeile auflösen.
                probe = (own.top, pos[1]) if own is not None else pos
                got = self.formula(formula, sheet.name, probe, used=used)
                found = None if got is None or got[0] == formula else (got[0], got[1], used)
                self.memo[key] = _ROWWISE if own is None and found is not None and _ROW in found[0] else found
                done = found
            if done is None:
                continue
            for t in done[2].rows:
                if t not in style.rows:
                    style.rows.append(t)
            for t in done[2].columns:
                if t not in style.columns:
                    style.columns.append(t)
            for name, target in done[2].names.items():
                style.names.setdefault(name, target)
            resolved = done[0]
            if _ROW in resolved:
                if own is not None and not own.top <= pos[0] <= own.bottom:
                    continue
                resolved = resolved.replace(_ROW, str(pos[0]))
            changes[pos] = replace(cell, formula=resolved, typed=cell.typed if cell.typed is not None else done[1])
        sheet.cells.update(changes)
        if style.rows or style.columns or style.names:
            sheet.refs = style


_MISSING, _ROWWISE = object(), object()


def resolve(sheets: list[Sheet], names: dict[str, str] | None = None, local: dict[str, dict[str, str]] | None = None) -> None:
    """Strukturierte Bezüge und Namen in allen Formeln der Mappe auflösen (die Zellen werden ersetzt)."""
    if not any(s.tables for s in sheets) and not names and not local:
        return
    resolver = _Resolver(sheets, names or {}, local or {})
    if not resolver.tables and resolver.words is None:
        return
    for sheet in sheets:
        resolver.sheet(sheet)


# --- Korrekturen im Stil der Mappe ---


def _area_of(text: str) -> Area:
    a, _, b = text.replace("$", "").partition(":")
    b = b or a
    ma, mb = re.fullmatch(r"([A-Z]+)(\d+)", a), re.fullmatch(r"([A-Z]+)(\d+)", b)
    return int(ma.group(2)), column_number(ma.group(1)), int(mb.group(2)), column_number(mb.group(1))


def styled(formula: str, at: Pos, style: Style | None, sheet: str = "") -> str:
    """Eine englische Formel ohne „=“ für die Zelle at, so geschrieben wie die Mappe: [@Menge] in Tabellenzeilen,
    Verkauf[Umsatz] für ganze Tabellenspalten, MwSt statt $I$1, wenn das Blatt diese Schreibweise benutzt."""
    if style is None or not formula:
        return formula
    own = next((t for t in style.rows if t.holds(at)), None)
    targets = {v: n for n, v in style.names.items()}

    def swap(match: re.Match) -> str:
        whole = match.group(0)
        if whole in targets:
            return targets[whole]
        prefix, ref = match.group(1), match.group(2)
        if prefix:
            return whole
        if own is not None and "$" not in ref and ":" not in ref:
            r, c, _, _ = _area_of(ref)
            if r == at[0] and own.area[1] <= c <= own.area[3]:
                return f"[@{_display_column(own.columns[c - own.area[1]])}]"
        if ":" in ref:
            r1, c1, r2, c2 = _area_of(ref)
            for t in style.columns:
                if t.sheet == sheet and (r1, r2) == (t.top, t.bottom) and t.area[1] <= c1 <= c2 <= t.area[3]:
                    first, last = t.columns[c1 - t.area[1]], t.columns[c2 - t.area[1]]
                    esc = lambda s: re.sub(r"([\[\]#'])", r"'\1", s)  # noqa: E731
                    return f"{t.name}[{esc(first)}]" if c1 == c2 else f"{t.name}[[{esc(first)}]:[{esc(last)}]]"
        return whole

    out = []
    for i, part in enumerate(re.split(r'("(?:[^"]|"")*")', formula)):
        out.append(part if i % 2 else _A1.sub(swap, part))
    return "".join(out)


def restyle(done: dict, cells: dict, style: Style | None, sheet: str, german) -> dict:
    """Die Hinweise eines Blatts für das Fenster: Korrekturen im Stil der Mappe, „was“ so, wie Excel die Zelle zeigt
    (sonst hält das Fenster sie für inzwischen geändert), und Formeln im Text so, wie sie in Excel stehen.
    german: die Funktion, die eine englische Formel deutsch schreibt (bereiche._german_formula)."""
    typed = {p: c for p, c in cells.items() if c.typed is not None} if style is not None or _any_typed(cells) else {}
    if not typed and style is None:
        return done
    from .formel import position  # noqa: PLC0415 - formel braucht bezuege nicht

    for hint in done.get("hints", []):
        if style is not None and hint.get("fix"):
            _as_name(hint, cells, style, sheet)
        swaps: list[tuple[str, str]] = []
        fix = hint.get("fix")
        if fix:
            new, changed = {}, False
            for where, put in fix.get("cells", {}).items():
                if isinstance(put, str) and put.startswith("=") and style is not None:
                    text = styled(put[1:], position(where), style, sheet)
                    if text != put[1:]:
                        changed = True
                        swaps.append(("=" + german(put[1:]), "=" + german(text)))
                    new[where] = "=" + text
                else:
                    new[where] = put
            if changed:
                fix["a1"] = fix["cells"]
                fix["cells"] = new
            for where in list(fix.get("was", {})):
                cell = typed.get(position(where))
                if cell is not None and isinstance(fix["was"][where], str):
                    fix["was"][where] = cell.typed
        for where in set(re.findall(r"(?<![A-Za-z$])([A-Z]{1,3}\d{1,7})(?![\d(])", (hint.get("text") or "") + " " + (hint.get("tip") or ""))):
            cell = typed.get(position(where))
            if cell is not None and cell.formula:
                swaps.append(("=" + german(cell.formula), "=" + german(cell.typed)))
        if style is not None:
            # Eine benannte Zelle heißt im Text, wie sie in der Mappe heißt: „=Umsatz * Kostenquote“.
            for name, target in style.names.items():
                if _SHEET_REF.fullmatch(target) is None:
                    continue
                pattern = re.compile(r"(?<![\w$!.'])" + re.escape(target) + r"(?![\w(:])")
                for key in ("text", "tip"):
                    if hint.get(key) and target in hint[key]:
                        kept = hint[key].replace(f" ({target})", "\x01")
                        hint[key] = pattern.sub(lambda _: name, kept).replace("\x01", f" ({target})")
        for old, new_text in swaps:
            if old == new_text:
                continue
            for key in ("text", "tip"):
                if hint.get(key):
                    hint[key] = _swap(hint[key], old, new_text)
            if fix:
                if isinstance(fix.get("german"), str):
                    fix["german"] = _swap(fix["german"], old, new_text)
                if isinstance(fix.get("parts"), list):
                    fix["parts"] = [_swap(p, old, new_text) if isinstance(p, str) else p for p in fix["parts"]]
    return done


_RANGE_TEXT = re.compile(r"(?<![A-Za-z0-9_$!.'])\$?[A-Z]{1,3}\$?\d{1,7}:\$?[A-Z]{1,3}\$?\d{1,7}(?![\d(])")


def _as_name(hint: dict, cells: dict, style: Style, sheet: str) -> None:
    """Ändert eine Korrektur in allen Formeln nur den Bereich eines Namens (Kurstabelle = $A$6:$B$8 → $A$6:$B$9),
    ist der Fehler der Name selbst: Dann wird der Name korrigiert, nicht jede Formel (die rechneten danach mit
    einem festen Bereich statt mit dem Namen). Aus der Korrektur wird eine Aktion für das Fenster."""
    fix = hint["fix"]
    puts = fix.get("cells", {})
    if not puts or fix.get("format") or fix.get("decimals"):
        return
    from .formel import position  # noqa: PLC0415

    for name, target in style.names.items():
        if ":" not in target or "!" in target:
            continue
        wanted = None
        for where, put in puts.items():
            cell = cells.get(position(where))
            if cell is None or not cell.formula or not isinstance(put, str) or not put.startswith("="):
                return
            if not re.search(r"(?<![\w.])" + re.escape(name) + r"(?![\w(])", cell.typed or "", re.IGNORECASE):
                break
            new = [r for r in _RANGE_TEXT.findall(put[1:]) if r not in _RANGE_TEXT.findall(cell.formula)]
            if len(new) != 1 or (wanted is not None and new[0] != wanted) or \
                    _swap(cell.formula, target, new[0]) != put[1:]:
                break
            wanted = new[0]
        else:
            if wanted is None:
                continue
            plain_old = target.replace("$", "")
            prefix = f"{_quote(sheet)}!"
            hint.pop("fix")
            hint["action"] = {"kind": "name", "name": name, "to": f"={prefix}{wanted}", "was": f"={prefix}{target}"}
            text = hint.get("text") or ""
            hint["text"] = text.replace(plain_old, f"{name} ({plain_old})", 1)
            hint["tip"] = T(f"Der Bereich ist der Name {name}. Formeln → Namens-Manager → {name}: Bezieht sich auf "
                            f"={prefix}{wanted}. Dann rechnen alle Formeln mit {name} richtig.",
                            f"The range is the defined name {name}. Formulas → Name Manager → {name}: Refers to "
                            f"={prefix}{wanted}. Then all formulas with {name} calculate correctly.")
            return


def _swap(text: str, old: str, new: str) -> str:
    """old durch new ersetzen, aber nur als ganze Formel (nicht =C5 in =C55)."""
    return re.sub(re.escape(old) + r"(?![\w$(:\[])", lambda _: new, text)


def _any_typed(cells: dict) -> bool:
    return any(c.typed is not None for c in cells.values())
