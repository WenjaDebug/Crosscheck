"""Werte aus einer anderen Datei: =[Budget.xlsx]Plan!B3 holt, was dort zuletzt stand. Geprüft werden kann davon nur
der zuletzt gespeicherte Wert; wurde die andere Datei seither geändert, verschoben oder umbenannt, stimmt er vielleicht
nicht mehr, ohne dass es auffällt. Ein Hinweis je Blatt, mit den Zellen und den Dateien.

Ein häufiger Unfall: Beim Kopieren eines Blatts aus einer anderen Mappe nimmt Excel die Bezüge auf die alte Datei mit.
Gibt es das Blatt, auf das gezeigt wird, auch in dieser Mappe, ist meist dieses gemeint. Dann zeigt der Hinweis, was
die Formel hier rechnen würde, und Einsetzen schreibt den Bezug auf das eigene Blatt."""

from __future__ import annotations

import re
from urllib.parse import unquote

from . import einsetzen, formel
from .analyse import Finding, Sheet, _checked, address, said
from .sprache import T
from .workbook import Layout
from .xlsx import Cell

Pos = tuple[int, int]
MAX_LISTED = 3
MAX_CELLS = 20
# '[Budget.xlsx]Plan 2025'!B3, [1]Plan!B3, 'C:\Daten\[Budget.xlsx]Plan'!B3, [1]!Name
_LINK = re.compile(r"'((?:[^']|'')*?)\[([^\]]+)\]((?:[^']|'')*)'!|(?<![\w.'\]])\[([^\]]+)\]([^\W\d][\w.]*)?!")


def links_in(formula: str) -> list[tuple[str, str, str]]:
    """Die Bezüge auf andere Dateien: (Text des Bezugs bis „!“, Datei oder Nummer, Blatt)."""
    found = []
    for part in re.split(r'("(?:[^"]|"")*")', formula)[::2]:
        for match in _LINK.finditer(part):
            if match.group(2) is not None:
                found.append((match.group(0), match.group(2), match.group(3).replace("''", "'")))
            else:
                found.append((match.group(0), match.group(4), match.group(5) or ""))
    return found


def checks(cells: dict[Pos, Cell], plan: Layout, name: str | None, book: dict[str, dict[Pos, Cell]],
           files: dict[str, str] | None = None) -> list[Finding]:
    """files: die Nummern der Datei (1 in [1]Plan!B3) und ihre Namen, aus der .xlsx-Datei."""
    files = files or {}
    uses: dict[Pos, list[tuple[str, str, str]]] = {}
    for pos, cell in cells.items():
        if cell.formula and "[" in cell.formula and "!" in cell.formula:
            found = links_in(cell.formula)
            if found:
                uses[pos] = found
    if not uses:
        return []
    where = sorted(uses, key=lambda p: (p[1], p[0]))
    sheet = Sheet(cells, plan, name, book)
    local = _local(sheet, uses, book)
    out = []
    if local:
        out.append(local)
        where = [p for p in where if address(p) not in (local.fix or {}).get("cells", {})]
        if not where:
            return out
    named = sorted({_file(f, files) for p in where for _, f, _ in uses[p]})
    cells_listed = _listed([address(p) for p in where], T("weitere", "more"))
    files_listed = _listed([T("„" + n + "“", "“" + n + "”") for n in named], T("weitere", "more"))
    verb = "holt" if len(where) == 1 else "holen"
    source = "einer anderen Datei" if len(named) == 1 else "anderen Dateien"
    text = T(f"{cells_listed} {verb} Werte aus {source}: "
             f"{files_listed}. Geprüft sind nur die Werte, die Excel zuletzt von dort "
             "geholt hat. Wurde die Datei seither geändert, verschoben oder umbenannt, stimmen sie vielleicht nicht mehr, "
             "ohne dass es auffällt.",
             f"{cells_listed} {'gets' if len(where) == 1 else 'get'} values from "
             f"{'another file' if len(named) == 1 else 'other files'}: {files_listed}. Only the values Excel last fetched "
             "from there are checked. If the file has been changed, moved or renamed since, they may no longer be right "
             "without anyone noticing.")
    tip = T("In Excel zeigt Daten → Arbeitsmappenlinks (in älteren Versionen „Verknüpfungen bearbeiten“), woher die Werte "
            "kommen; dort lassen sie sich aktualisieren.",
            "In Excel, Data → Workbook Links (in older versions “Edit Links”) shows where the values come from; "
            "you can update them there.")
    cell = ",".join(address(p) for p in where[:MAX_CELLS])
    out.append(Finding(f"extern:{address(where[0])}", text, "info", tip, cell=cell))
    return out


def _listed(items: list[str], more: str) -> str:
    """B3, B4 und B5; bei mehr als MAX_LISTED: B3, B4, B5 und 4 weitere."""
    word = T(" und ", " and ")
    if len(items) > MAX_LISTED:
        return ", ".join(items[:MAX_LISTED]) + f"{word}{len(items) - MAX_LISTED} {more}"
    return items[-1] if len(items) == 1 else ", ".join(items[:-1]) + f"{word}{items[-1]}"


def _file(raw: str, files: dict[str, str]) -> str:
    """„Budget.xlsx“ aus [1], aus einem Pfad oder so, wie es dasteht."""
    # In der .xlsx-Datei steht der Pfad als Adresse (Budget%202025.xlsx), in der Formel so, wie er heißt.
    raw = unquote(files[raw]) if raw in files else raw
    return re.split(r"[\\/:]", raw)[-1] or raw


def _local(sheet: Sheet, uses: dict[Pos, list[tuple[str, str, str]]], book: dict[str, dict[Pos, Cell]]) -> Finding | None:
    """Bezüge auf ein Blatt einer anderen Datei, das es genauso in dieser Mappe gibt: wohl beim Kopieren mitgekommen."""
    names = {n.casefold(): n for n in book}
    if sheet.name:
        names.setdefault(sheet.name.casefold(), sheet.name)
    changes: dict[Pos, str] = {}
    for pos, found in sorted(uses.items(), key=lambda item: (item[0][1], item[0][0])):
        formula = sheet.cells[pos].formula or ""
        rewritten = formula
        for text, _, target in found:
            mine = names.get(target.casefold()) if target else None
            if mine is None:
                rewritten = None
                break
            quoted = f"'{mine}'" if re.search(r"[^\w.]", mine) or mine[:1].isdigit() else mine
            rewritten = rewritten.replace(text, f"{quoted}!")
        if rewritten and rewritten != formula:
            changes[pos] = rewritten
    if not changes or len(changes) > einsetzen.MAX_CELLS:
        return None
    fix = einsetzen.formulas(sheet, changes)
    if fix is None:
        return None
    first = min(changes, key=lambda p: (p[1], p[0]))
    calc, _ = sheet.evaluator()
    try:
        here = calc.run(formel.parse(changes[first]))
    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError, IndexError, KeyError):
        return None
    now = sheet.inputs.get(first)
    target = next(t for _, _, t in uses[first] if t)
    cells = sorted(changes, key=lambda p: (p[1], p[0]))
    rest = [address(p) for p in cells[1:]]
    others = T(f" Ebenso in {_listed(rest, 'weiteren Zellen')}.", f" The same in {_listed(rest, 'more cells')}.") if rest else ""
    differs = "" if isinstance(now, float) and isinstance(here, float) and abs(now - here) <= 1e-9 * max(1.0, abs(here)) \
        else T(f" Mit dem Blatt dieser Mappe ergäbe {address(first)} {said(here)} statt {said(now)}.",
               f" With the sheet in this workbook, {address(first)} would give {said(here)} instead of {said(now)}.")
    text = T(f"{address(first)} holt seinen Wert vom Blatt „{target}“ einer anderen Datei, und ein Blatt „{target}“ gibt "
             f"es auch in dieser Mappe. Das passiert, wenn ein Blatt aus einer anderen Mappe kopiert wird: Die Bezüge "
             f"zeigen dann weiter auf die alte Datei.{differs}{others}",
             f"{address(first)} gets its value from the sheet “{target}” in another file, and there is a sheet “{target}” "
             f"in this workbook too. This happens when a sheet is copied from another workbook: the references "
             f"then still point to the old file.{differs}{others}")
    checked = _checked(sheet, changes[first])
    tip = T(f"Korrigiert in {address(first)}: {checked or fix['german']}.",
            f"Corrected in {address(first)}: {checked or fix['german']}.")
    return Finding(f"copylink:{address(first)}", text, "pruefen", tip, fix=fix,
                   cell=",".join(address(p) for p in cells[:MAX_CELLS]))
