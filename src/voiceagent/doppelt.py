"""Eine Spalte holt je Zeile die Summe ihres Schlüssels, und der Schlüssel steht doppelt in der Tabelle.

Ohne KI: In der Kundentabelle rechnet „Umsatz“ je Zeile =SUMMEWENNS(Rechnungen[Umsatz];Rechnungen[Kunde];[@Kunde]).
Steht ein Kunde zweimal in der Tabelle (alte und neue Stammdaten), bekommen beide Zeilen seinen ganzen Umsatz. Wer
die Spalte summiert, zählt ihn doppelt. Das gilt, ob die doppelte Zeile Absicht ist oder nicht: Die Summe über die
Spalte ist zu groß. Korrektur: nur die gültige Zeile (Spalte „Aktuell“ = Ja) oder das erste Vorkommen zählt.
"""

from __future__ import annotations

from . import einsetzen, formel
from .analyse import Finding, Sheet, _fmt, address
from .sprache import T
from .xlsx import column_letters

Pos = tuple[int, int]

_KEYED = {"SUMIFS", "SUMIF", "COUNTIFS", "COUNTIF", "AVERAGEIFS", "AVERAGEIF", "MAXIFS", "MINIFS"}
MAX_FOUND = 2


def _keyed(node, row: int) -> int | None:
    """Die Spalte des Schlüssels in derselben Zeile, nach dem die Formel eine andere Tabelle zusammenfasst."""
    if node is None:
        return None
    for call in formel.calls(node):
        if call.value not in _KEYED:
            continue
        args = list(call.args)
        if call.value in ("SUMIF", "AVERAGEIF", "COUNTIF"):
            pairs = [(args[0], args[1])] if len(args) >= 2 else []
        else:
            start = 0 if call.value == "COUNTIFS" else 1
            pairs = [(args[i], args[i + 1]) for i in range(start, len(args) - 1, 2)]
        for area, crit in pairs:
            if (isinstance(area, formel.Node) and area.kind == "range" and isinstance(crit, formel.Node)
                    and crit.kind == "ref" and crit.value[0] == row and not area.value[0] <= row <= area.value[2]):
                return crit.value[1]
    return None


def _source(sheet: Sheet, pos: Pos) -> tuple[int, Pos] | None:
    """(Schlüsselspalte, Zelle mit der Summe) für pos selbst oder eine Zelle derselben Zeile, die pos nur übernimmt."""
    node = sheet.nodes.get(pos)
    key = _keyed(node, pos[0])
    if key is not None:
        return key, pos
    if node is not None and node.kind == "ref" and node.value[0] == pos[0] and node.value != pos:
        key = _keyed(sheet.nodes.get(node.value), pos[0])
        if key is not None:
            return key, node.value
    return None


def _yes_column(sheet: Sheet, head_row: int, rows: list[int], dupes: dict) -> tuple[int, str] | None:
    """Eine Spalte, die je doppeltem Schlüssel genau eine Zeile als gültig markiert (Aktuell = Ja)."""
    from .analyse import _CURRENT, _YES

    for c in range(1, 60):
        cell = sheet.cells.get((head_row, c))
        if cell is None or cell.formula or not isinstance(cell.value, str) or not _CURRENT.search(cell.value.strip()):
            continue
        marks = {r: str(sheet.inputs.get((r, c))).strip() for r in rows}
        yes = {v for v in marks.values() if v.casefold() in _YES}
        if len(yes) != 1:
            continue
        word = yes.pop()
        if all(sum(marks[r] == word for r in group) == 1 for group in dupes.values()):
            return c, word
    return None


def _summed_by(sheet: Sheet, col: int, rows: list[int]) -> Pos | None:
    lo, hi = rows[0], rows[-1]
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or lo <= pos[0] <= hi and pos[1] == col:
            continue
        for call in formel.calls(node):
            if call.value not in ("SUM", "SUBTOTAL", "AGGREGATE", "SUMPRODUCT"):
                continue
            for arg in call.args:
                if isinstance(arg, formel.Node) and arg.kind == "range":
                    r1, c1, r2, c2 = arg.value
                    if c1 <= col <= c2 and r1 <= lo and r2 >= hi:
                        return pos
    return None


def checks(sheet: Sheet) -> list[Finding]:
    columns: dict[tuple[int, int], list[tuple[int, Pos]]] = {}
    for pos in sorted(sheet.nodes):
        source = _source(sheet, pos)
        if source is not None:
            columns.setdefault((pos[1], source[0]), []).append((pos[0], source[1]))
    found: list[Finding] = []
    for (col, key), cells in sorted(columns.items()):
        rows = [r for r, _ in cells]
        if len(rows) < 5 or rows[-1] - rows[0] + 1 != len(rows):
            continue
        groups: dict = {}
        for r in rows:
            value = sheet.inputs.get((r, key))
            if value is None or str(value).strip() == "":
                continue
            groups.setdefault(formel._norm(value), []).append(r)  # noqa: SLF001
        dupes = {k: g for k, g in groups.items() if len(g) > 1}
        total = _summed_by(sheet, col, rows)
        if not dupes or total is None:
            continue
        numbers = {r: sheet.inputs.get((r, col)) for r in rows}
        if not all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in numbers.values()):
            continue
        head_row = rows[0] - 1
        marker = _yes_column(sheet, head_row, rows, dupes)
        extra = 0.0
        for group in dupes.values():
            keep = next((r for r in group if marker and str(sheet.inputs.get((r, marker[0]))).strip() == marker[1]),
                        group[0])
            extra += sum(numbers[r] for r in group if r != keep)
        if abs(extra) < 1e-9:
            continue
        changes = {}
        k, c = column_letters(key), column_letters(marker[0]) if marker else ""
        for r in rows:
            text = (sheet.cells[(r, col)].formula or "").lstrip("=")
            test = f'${c}{r}="{marker[1]}"' if marker else f"COUNTIF(${k}${rows[0]}:${k}{r},${k}{r})=1"
            changes[(r, col)] = f"IF({test},{text},0)"
        fix = einsetzen.formulas(sheet, changes) if len(changes) <= einsetzen.MAX_CELLS else None
        first = sorted(dupes.values())[0]
        name = sheet.cells.get((head_row, col))
        head = f" ({name.value})" if name is not None and isinstance(name.value, str) else ""
        key_head = sheet.cells.get((head_row, key))
        key_name = key_head.value if key_head is not None and isinstance(key_head.value, str) else T("Schlüssel", "key")
        n_more = len(dupes) - 1
        more = T(f", und {n_more} weitere stehen auch doppelt",
                 f", and {n_more} more {'appears' if n_more == 1 else 'appear'} twice as well") if len(dupes) > 1 else ""
        via = T(f" (über {address(cells[0][1])})", f" (via {address(cells[0][1])})") if cells[0][1] != (rows[0], col) \
            else ""
        rule = T(f"nur die Zeilen mit {name_of(sheet, head_row, marker[0])} „{marker[1]}“",
                 f"only the rows with {name_of(sheet, head_row, marker[0])} “{marker[1]}”") if marker else \
            T("nur das erste Vorkommen jedes Schlüssels", "only the first occurrence of each key")
        span = f"{address((rows[0], col))}:{address((rows[-1], col))}{head}"
        shown = _fmt(sheet.inputs.get((first[0], key)))
        listed = ", ".join(map(str, first[:4]))
        text = T(f"{span} holt für jede Zeile die Summe ihrer "
                 f"{key_name}{via}. {key_name} „{shown}“ steht aber "
                 f"{len(first)}-mal in der Tabelle (Zeilen {listed}){more}. Jede dieser Zeilen "
                 f"bekommt den ganzen Betrag, und {address(total)} summiert die Spalte: Dort zählen {_fmt(extra)} "
                 f"doppelt.",
                 f"{span} fetches, for each row, the sum for its {key_name}{via}. But {key_name} “{shown}” appears "
                 f"{len(first)} times in the table (rows {listed}){more}. Each of these rows gets the full amount, "
                 f"and {address(total)} sums the column: {_fmt(extra)} is counted twice there.")
        found.append(Finding(f"dupkey:{address((rows[0], col))}", text, "fehler",
                             T(f"Zähl {rule}.", f"Count {rule}."), fix=fix))
        if len(found) >= MAX_FOUND:
            break
    return found


def name_of(sheet: Sheet, row: int, col: int) -> str:
    cell = sheet.cells.get((row, col))
    return cell.value.strip() if cell is not None and isinstance(cell.value, str) else column_letters(col)
