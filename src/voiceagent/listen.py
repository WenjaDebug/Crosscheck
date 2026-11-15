"""Ein fester Bereich, der vor dem Ende seiner Liste aufhört.

Ohne KI: Unter einem Titel („Negative Stichworte“) steht eine Liste aus Wörtern, ohne Lücke bis zur nächsten leeren
Zelle. Liest eine Formel diese Liste mit festem Bereich ($A$30:$A$31), muss der Bereich bis zum letzten Eintrag
reichen. Endet er früher, fehlen die Einträge darunter still: SUCHEN findet „falsch“ nie, ZÄHLENWENN zählt es nicht.
"""

from __future__ import annotations

import re

from . import einsetzen
from .analyse import Finding, address
from .sprache import T
from .workbook import Sheet
from .xlsx import column_letters

Pos = tuple[int, int]

_FIXED = re.compile(r"(?<![A-Za-z0-9_!\]])\$([A-Z]{1,3})\$(\d+):\$([A-Z]{1,3})\$(\d+)(?![0-9(])")
# Wo eine Liste von Wörtern gelesen wird; SUMME und MITTELWERT über Zahlen prüft bereiche.py.
_LIST_USE = re.compile(r"\b(SEARCH|FIND|COUNTIFS?|MATCH|XMATCH|XLOOKUP|ISNUMBER|OR|SUMPRODUCT)\(", re.I)
MAX_FOUND = 4


def _col(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _text(sheet: Sheet, pos: Pos) -> str | None:
    cell = sheet.cells.get(pos)
    if cell is None or cell.formula or not isinstance(cell.value, str) or not cell.value.strip():
        return None
    value = cell.value.strip()
    return value if re.search(r"[^\W\d_]", value) else None


def _list_end(sheet: Sheet, col: int, start: int) -> int:
    row = start
    while _text(sheet, (row + 1, col)) is not None:
        row += 1
    return row


def checks(sheet: Sheet) -> list[Finding]:
    uses: dict[tuple, list[Pos]] = {}
    for pos, cell in sorted(sheet.cells.items()):
        formula = cell.formula or ""
        if not formula or not _LIST_USE.search(formula):
            continue
        for m in _FIXED.finditer(formula):
            c1, r1, c2, r2 = _col(m.group(1)), int(m.group(2)), _col(m.group(3)), int(m.group(4))
            if c1 == c2 and r2 >= r1:
                uses.setdefault((c1, r1, r2, m.group(0)), []).append(pos)
    found = []
    for (col, top, bottom, written), cells in sorted(uses.items(), key=lambda kv: kv[1][0]):
        title = _text(sheet, (top - 1, col))
        items = [_text(sheet, (r, col)) for r in range(top, bottom + 1)]
        if title is None or any(v is None for v in items):
            continue
        end = _list_end(sheet, col, bottom)
        if end == bottom:
            continue
        # Dieselbe Liste anderswo in voller Länge gelesen: Dann ist der kurze Bereich Absicht (nur die ersten).
        if any(c == col and t == top and b >= end for c, t, b, _ in uses):
            continue
        missing = [(r, _text(sheet, (r, col))) for r in range(bottom + 1, end + 1)]
        letters = column_letters(col)
        right = f"${letters}${top}:${letters}${end}"
        changes = {}
        for pos in cells:
            text = (sheet.cells[pos].formula or "").lstrip("=")
            changes[pos] = text.replace(written, right)
        fix = einsetzen.formulas(sheet, changes) if len(changes) <= einsetzen.MAX_CELLS else None
        shown = ", ".join(T(f"„{v}“ ({address(p)})", f"“{v}” ({address(p)})") for p, v in [((r, col), v) for r, v in missing][:3])
        first = cells[0]
        more = T(f" Genauso in {len(cells) - 1} weiteren Formeln.", f" The same in {len(cells) - 1} more formulas.") if len(cells) > 1 else ""
        one = len(missing) == 1
        text = T(f"{address(first)} liest die Liste „{title}“ nur bis {address((bottom, col))} ({written}). Die Liste "
                 f"geht aber bis {address((end, col))}: {shown} {'fehlt' if one else 'fehlen'} in der "
                 f"Formel und {'wird' if one else 'werden'} nie berücksichtigt.{more}",
                 f"{address(first)} reads the list “{title}” only up to {address((bottom, col))} ({written}). But the list "
                 f"goes down to {address((end, col))}: {shown} {'is' if one else 'are'} missing from the "
                 f"formula and {'is' if one else 'are'} never taken into account.{more}")
        found.append(Finding(f"listshort:{address(first)}", text, "fehler",
                             T(f"Nimm {right} statt {written}.", f"Use {right} instead of {written}."), fix=fix))
        if len(found) >= MAX_FOUND:
            break
    return found
