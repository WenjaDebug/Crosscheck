"""Korrekturen zum Einsetzen: die nachgerechnete Formel aus dem Tipp eines Hinweises, so dass das Seitenfenster sie
mit einem Klick in Excel schreiben (und wieder rückgängig machen) kann.

Die Prüfungen schreiben ihre Korrektur auf Deutsch in den Tipp („Korrigiert: =SUMME(E2:E5), nachgerechnet 1137,5.“,
„Korrigiert in C2: =B2/$B$5, dann bis C4 herunterziehen.“). Hier wird daraus die englische Formel, die Excel über
Office.js annimmt, für jede Zelle, in die sie gehört. Eingesetzt wird nur, was hin und zurück genau dieselbe deutsche
Formel ergibt, sich auf dem Blatt ausrechnen lässt und, wo der Tipp einen Wert nennt, genau diesen Wert ergibt.

Wo der Tipp die Korrektur nur in Worten sagt („Schreib FALSCH als letztes Argument“, „Zieh den Bereich bis D30“),
baut die Prüfung sie selbst mit formulas(), values(), stretched() oder exact(), ebenso geprüft. Für „Alle einsetzen“
setzt pruefer.together() mehrere zusammen ein (applied(), refilled()) und prüft das Blatt danach neu: Was sich
gegenseitig stört, bleibt draußen; merged() macht daraus eine Korrektur, die sich zusammen zurücknehmen lässt.
"""

from __future__ import annotations

import re
from itertools import pairwise

from . import formel
from .analyse import Finding, Sheet, address
from .bereiche import _german_formula
from .sprache import T
from .sprache import english as _in_english
from .workbook import FUNCTIONS, pattern
from .xlsx import Cell, Cells, _date_text, cut_of, kept_text, number_text, shift

Pos = tuple[int, int]

# Mehr Zellen schreibt ein Klick nicht; so lange Spalten zieht man besser selbst herunter. Eine Spalte mit 20 000
# Fällen geht noch: Excel schreibt sie in einem Schritt.
MAX_CELLS = 20_000
_GERMAN = {german: english for english, german in FUNCTIONS.items()}
_CELL = r"\$?[A-Z]{1,3}\$?[1-9][0-9]*"
# Eine deutsche Formel: alles bis zum nächsten Leerzeichen, Texte in Anführungszeichen dürfen Leerzeichen enthalten.
_FORMULA = r'=(?P<f>(?:"(?:[^"]|"")*"|[^\s"])+)'
_WAYS = [
    (re.compile(rf"Korrigiert in (?P<at>{_CELL}): {_FORMULA},? dann bis (?P<to>{_CELL}) herunterziehen"), "to"),
    (re.compile(rf"Schreib in (?P<at>{_CELL}) {_FORMULA}(?: \([^)]*\))? und zieh es bis (?P<to>{_CELL})"), "to"),
    (re.compile(rf"Schreib in (?P<at>{_CELL}) {_FORMULA}(?: \([^)]*\))? und zieh es (?:wieder )?herunter"), "down"),
    (re.compile(rf"Korrigiert: {_FORMULA}"), "here"),
    (re.compile(rf"(?:^|\. )Nimm {_FORMULA}"), "here"),
    (re.compile(rf"^Mit {_FORMULA} rechnet die Zelle selbst"), "here"),
    (re.compile(rf"(?:^|\. )Sicherer ist {_FORMULA}"), "here"),
    (re.compile(rf"^Gemeint ist meist die Quote über alle Zeilen: {_FORMULA}"), "here"),
    (re.compile(rf"Oder rechne (?P<at>{_CELL}) direkt aus der Tabelle: {_FORMULA}"), "here"),
]
_SAID = re.compile(r"^,? \(?nachgerechnet (?:ergibt das )?(?:(?P<d>\d{2}\.\d{2}\.\d{4})|(?P<n>-?\d+(?:,\d+)?)(?P<pct> ?%)?)"
                   r"(?=[.,;) ]|$)")
# Dieselben Korrekturen in englischen Tipps. Zahlen und Formeln stehen dort noch deutsch (sprache.output schreibt sie
# erst am Ende um); nur die Wörter drumherum sind englisch. Steht ein Tipp (noch) deutsch da, gilt _WAYS.
_WAYS_EN = [
    (re.compile(rf"(?:Corrected|Fixed) in (?P<at>{_CELL}): {_FORMULA},? then (?:fill|drag|copy) (?:it )?down to "
                rf"(?P<to>{_CELL})"), "to"),
    (re.compile(rf"In (?P<at>{_CELL}),? (?:write|enter|type) {_FORMULA}(?: \([^)]*\))? and copy it to (?P<to>{_CELL})"), "to"),
    (re.compile(rf"(?:Write|Enter|Type) {_FORMULA} in (?P<at>{_CELL})(?: \([^)]*\))? and copy it to (?P<to>{_CELL})"), "to"),
    (re.compile(rf"In (?P<at>{_CELL}),? (?:write|enter|type) {_FORMULA}(?: \([^)]*\))? and (?:fill|drag|copy) it "
                rf"(?:back )?down"), "down"),
    (re.compile(rf"(?:Write|Enter|Type) {_FORMULA} in (?P<at>{_CELL})(?: \([^)]*\))? and (?:fill|drag|copy) it "
                rf"(?:back )?down"), "down"),
    (re.compile(rf"(?:Corrected|Fixed): {_FORMULA}"), "here"),
    (re.compile(rf"(?:^|\. )Use {_FORMULA}"), "here"),
    (re.compile(rf"^With {_FORMULA},? the cell (?:calculates|computes)"), "here"),
    (re.compile(rf"(?:^|\. )(?:Safer is|Safer:|Safer would be|It is safer to use|A safer option is) {_FORMULA}"), "here"),
    (re.compile(rf"^[^=:.]*\b(?:[Uu]sually|[Mm]ostly|[Tt]ypically)\b[^=:.]*\ball rows\b[^=:.]*: {_FORMULA}"), "here"),
    (re.compile(rf"Or (?:calculate|compute) (?P<at>{_CELL}) (?:directly )?from the table(?: directly)?: {_FORMULA}"), "here"),
]
_SAID_EN = re.compile(r"^,? \(?recalculated(?:,? (?:this|that|it) (?:gives|comes to|equals|is)|:)? "
                      r"(?:(?P<d>\d{2}\.\d{2}\.\d{4})|(?P<n>-?\d+(?:,\d+)?)(?P<pct> ?%)?)(?=[.,;) ]|$)")
# „Danach wieder bis unten herunterziehen.“: die ganze heruntergezogene Spalte.
_TO_BOTTOM = re.compile(r"bis unten herunterziehen|down to the (?:bottom|end|last row)")


def _said(rest: str):
    """Der Wert, den der Tipp hinter der Formel nennt („, nachgerechnet 577“, „, recalculated 577“), oder None."""
    if _in_english():
        found = _SAID_EN.match(rest)
        if found:
            return found
    return _SAID.match(rest)


def english(german: str) -> str:
    """SUMMEWENNS(D:D;B:B;"Nord") → SUMIFS(D:D,B:B,"Nord"): umgekehrt wie bereiche._german_formula."""
    out = []
    # Texte in Anführungszeichen und Blattnamen in Hochkommas ('Q1,2'!A1) bleiben, wie sie sind.
    for i, part in enumerate(re.split(r'("(?:[^"]|"")*"|\'(?:[^\']|\'\')*\')', german)):
        if i % 2 == 0:
            part = re.sub(r"(?<![A-Za-z0-9_.])([A-ZÄÖÜ][A-ZÄÖÜ0-9.]*)\(", lambda m: _GERMAN.get(m.group(1), m.group(1)) + "(", part)
            part = re.sub(r"(?<![A-Za-z0-9_.!'])(WAHR|FALSCH)(?![A-Za-z0-9_(!])",
                          lambda m: "TRUE" if m.group(1) == "WAHR" else "FALSE", part)
            part = re.sub(r"(?<![A-Z\d])(\d+),(\d+)", r"\1.\2", part).replace(";", ",")
        out.append(part)
    return "".join(out)


def _here(finding: Finding) -> Pos | None:
    tail = finding.key.partition(":")[2]
    return formel.position(tail) if re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]*", tail) else None


def _block(sheet: Sheet, start: Pos) -> list[Pos]:
    """Die Zellen ab start nach unten mit derselben Formel (heruntergezogen), wie Excel sie füllt."""
    cell = sheet.cells.get(start)
    if cell is None or not cell.formula:
        return [start]
    shape = pattern(cell.formula, start)
    cells, row = [start], start[0] + 1
    while len(cells) < MAX_CELLS:
        below = sheet.cells.get((row, start[1]))
        if below is None or not below.formula or pattern(below.formula, (row, start[1])) != shape:
            break
        cells.append((row, start[1]))
        row += 1
    return cells


def _close(value, said: str, percent: bool = False) -> bool:
    """Ergibt die Formel den Wert, den der Tipp nennt (so gerundet, wie er dasteht, „52,4 %“ als 0,524)?"""
    if not isinstance(value, float) or isinstance(value, bool):
        return False
    shown = float(said.replace(",", "."))
    digits = len(said.partition(",")[2])
    if percent:
        value *= 100
    return abs(value - shown) <= 0.5 * 10 ** -digits + 1e-9 * max(1.0, abs(shown))


def fix(sheet: Sheet, finding: Finding) -> dict | None:
    """{"range": "C2:C4", "cells": {"C2": "=B2/$B$5", …}, "was": {"C2": "B2/$F$2", …}, "german": "=B2/$B$5"} oder
    None, wenn nichts sicher einzusetzen ist. was ist die Formel, die beim Prüfen in der Zelle stand ("" ohne Formel),
    german die erste Formel so, wie sie im deutschen Excel aussieht."""
    tip = finding.tip or ""
    for way, reach in (_WAYS_EN + _WAYS) if _in_english() else _WAYS:
        match = way.search(tip)
        if match:
            break
    else:
        return None
    german = match.group("f").rstrip(".,")
    text = english(german)
    if _german_formula(text) != german:
        return None
    try:
        node = formel.parse(text)
    except (ValueError, IndexError, KeyError, formel.Unknown):
        return None
    start = formel.position(match.group("at").replace("$", "")) if "at" in way.groupindex else _here(finding)
    if start is None:
        return None
    if reach == "to":
        end = formel.position(match.group("to").replace("$", ""))
        if end[1] != start[1] or end[0] < start[0]:
            return None
        targets = [(row, start[1]) for row in range(start[0], end[0] + 1)]
    elif reach == "down" or _TO_BOTTOM.search(tip[match.end():]):
        targets = _block(sheet, start)
    else:
        targets = [start]
    if len(targets) > MAX_CELLS:
        return None
    calc, _ = sheet.evaluator()
    try:
        value = calc.run(node)
    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError):
        return None
    if isinstance(value, formel.Err) and "#" not in tip[match.end():match.end() + 80]:
        return None
    said = _said(tip[match.end():])
    if said and said.group("d"):
        # „nachgerechnet 09.04.2026“: das Ergebnis als Datum.
        if not isinstance(value, float) or _date_text(value, False) != said.group("d"):
            return None
    elif said and not _close(value, said.group("n"), bool(said.group("pct"))):
        return None
    cells = {}
    for pos in targets:
        moved = text if pos == start else shift(text, pos[0] - start[0], pos[1] - start[1])
        if "#BEZUG!" in moved:
            return None
        cells[address(pos)] = "=" + moved
    was = {address(pos): (sheet.cells[pos].formula or "") if pos in sheet.cells else "" for pos in targets}
    # Die Zellen liegen untereinander: Das Seitenfenster schreibt sie als einen Bereich.
    area = address(targets[0]) if len(targets) == 1 else f"{address(targets[0])}:{address(targets[-1])}"
    return {"range": area, "cells": cells, "was": was, "german": "=" + german}


def attach(sheet: Sheet, findings: list[Finding]) -> None:
    for item in findings:
        # Hat die Prüfung die Korrektur schon selbst gebaut, bleibt es dabei.
        if item.tip and "=" in item.tip and not item.fix:
            item.fix = fix(sheet, item)


def _area(targets: list[Pos]) -> str:
    """C2 oder C2:C9 für Zellen untereinander, sonst „C2 und C7“ oder „C2, C7 und 3 weitere“ (nur zum Anzeigen)."""
    if len(targets) == 1:
        return address(targets[0])
    if all(b == (a[0] + 1, a[1]) for a, b in pairwise(targets)):
        return f"{address(targets[0])}:{address(targets[-1])}"
    # Ein volles Rechteck (die Formeln einer neuen Zeile, G4002 bis K4002).
    top, left = min(r for r, _ in targets), min(c for _, c in targets)
    bottom, right = max(r for r, _ in targets), max(c for _, c in targets)
    if len(set(targets)) == (bottom - top + 1) * (right - left + 1):
        return f"{address((top, left))}:{address((bottom, right))}"
    if len(targets) <= 3:
        return ", ".join(address(p) for p in targets[:-1]) + T(f" und {address(targets[-1])}", f" and {address(targets[-1])}")
    return T(f"{address(targets[0])}, {address(targets[1])} und {len(targets) - 2} weitere",
             f"{address(targets[0])}, {address(targets[1])} and {len(targets) - 2} more")


def _was(sheet: Sheet, pos: Pos) -> str:
    """Was beim Prüfen in der Zelle stand, so wie das Seitenfenster es aus Excel liest: die Formel ohne „=“ ("" ohne
    Formel). Für eine Korrektur, die Werte schreibt, steht dort der Text selbst: {"text": …} (values())."""
    cell = sheet.cells.get(pos)
    return (cell.formula or "") if cell is not None else ""


def formulas(sheet: Sheet, changes: dict[Pos, str], errors: bool = False, calc=None,
             parts: dict[Pos, list[str]] | None = None) -> dict | None:
    """Eine Korrektur aus fertigen englischen Formeln je Zelle (ohne „=“), wie fix() sie aus dem Tipp liest: jede
    lässt sich auf dem Blatt ausrechnen, ein Fehlerwert nur mit errors (etwa #NV, wenn etwas fehlt, sichtbar sein soll).
    None, wenn eine davon nicht sicher ist. calc: ein eigener Rechner, etwa einer, der HEUTE() als heutiges Datum nimmt.
    parts: die geänderten Teile je Zelle. Kennt der Prüfer eine Funktion außerhalb davon nicht (PIVOTDATENZUORDNEN),
    reicht es, dass sich die geänderten Teile ausrechnen lassen: Der Rest der Formel bleibt, wie er war."""
    if not changes or len(changes) > MAX_CELLS:
        return None
    if calc is None:
        calc, _ = sheet.evaluator()
    targets = sorted(changes, key=lambda p: (p[1], p[0]))
    for pos in targets:
        text = changes[pos]
        try:
            value = calc.run(formel.parse(text))
        except formel.Unknown:
            if not (parts or {}).get(pos):
                return None
            try:
                values = [calc.run(formel.parse(part)) for part in parts[pos]]
            except (ValueError, IndexError, KeyError, TypeError, ZeroDivisionError, RecursionError, formel.Unknown):
                return None
            value = next((v for v in values if isinstance(v, formel.Err)), values[0])
        except (ValueError, IndexError, KeyError, TypeError, ZeroDivisionError, RecursionError):
            return None
        if (isinstance(value, formel.Err) and not errors) or "#BEZUG!" in text or "#REF!" in text:
            return None
        if english(_german_formula(text)) != text:
            return None
    return {"range": _area(targets), "cells": {address(p): "=" + changes[p] for p in targets},
            "was": {address(p): _was(sheet, p) for p in targets}, "german": "=" + _german_formula(changes[targets[0]])}


def values(sheet: Sheet, changes: dict[Pos, float], like: Pos | None, said: str) -> dict | None:
    """Eine Korrektur, die Zahlen statt Text in die Zellen schreibt (ein Datum als seine Excel-Zahl). Das Zahlenformat
    kommt von like, einer Zelle derselben Spalte, in der schon eine echte Zahl oder ein echtes Datum steht. said sagt,
    was geschrieben wird („die Zahlen“)."""
    if not changes or len(changes) > MAX_CELLS:
        return None
    targets = sorted(changes, key=lambda p: (p[1], p[0]))
    was = {}
    for pos in targets:
        cell = sheet.cells.get(pos)
        if cell is None or cell.formula:
            return None
        was[address(pos)] = {"text": cell.text if cell.text is not None else cell.value}
    found = {"range": _area(targets), "cells": {address(p): changes[p] for p in targets}, "was": was, "german": said}
    if like is not None:
        found["format"] = {address(p): address(like) for p in targets}
    return found


def cleared(sheet: Sheet, targets: list[Pos]) -> dict | None:
    """Eine Korrektur, die Zellen leer macht: ein Code wie 999 für „unbekannt“. MITTELWERT, SUMME, ZÄHLENWENN und
    Verwandte lassen leere Zellen aus. Nicht, wenn eine Formel eine der Zellen einzeln liest (=C812*2 rechnete dann
    mit 0 statt 999, wieder falsch, nur leiser)."""
    if not targets or len(targets) > MAX_CELLS:
        return None
    was = {}
    for pos in targets:
        cell = sheet.cells.get(pos)
        if cell is None or cell.formula:
            return None
        for reader in sheet.dependents.get(pos, ()):
            node = sheet.nodes.get(reader)
            if node is None or pos in formel.references(node, sheet.name)[0]:
                return None
        was[address(pos)] = {"text": cell.text if cell.text is not None else cell.value}
    targets = sorted(targets, key=lambda p: (p[1], p[0]))
    return {"range": _area(targets), "cells": {address(p): "" for p in targets}, "was": was, "german": T("leer", "empty"),
            "clear": True}


def rewritten(formula: str, change, moved: dict | None = None) -> str | None:
    """Die Formel mit change(call) für jeden Aufruf darin (der geänderte Aufruf oder None für unverändert), zurück als
    Text mit denselben $ wie vorher. moved: Bereiche, die change verlängert (alt → neu), behalten ihre $ auch.
    None, wenn sich nichts ändert oder es nicht sicher geht."""
    from .bereiche import _kept, _text

    try:
        node = formel.parse(formula)
    except (ValueError, IndexError, KeyError, formel.Unknown):
        return None
    kept = _kept(formula)
    for old, new in (moved or {}).items():
        if isinstance(old[0], str):
            # Ein Bereich auf einem anderen Blatt: (Blatt, Bereich), mit Blattnamen und $ wie vorher.
            name, _, shown = (kept.get(("xrange", old)) or "").rpartition("!")
            match = _RANGE.fullmatch(shown)
            if name and match:
                kept.setdefault(("xrange", new), f"{name}!{match[1]}{new[1][0]}{match[2]}{new[1][2]}")
            continue
        shown = kept.get(("range", old))
        match = _RANGE.fullmatch(shown or "")
        if match:
            kept.setdefault(("range", new), f"{match[1]}{new[0]}{match[2]}{new[2]}")

    def walk(item):
        if not isinstance(item, formel.Node):
            return item
        item = formel.Node(item.kind, item.value, tuple(walk(a) for a in item.args))
        return (change(item) or item) if item.kind == "call" else item

    new = walk(node)
    text = _text(new, kept)
    return text if text is not None and text != _text(node, kept) else None


_RANGE = re.compile(r"(\$?[A-Z]{1,3}\$?)[1-9][0-9]*(:\$?[A-Z]{1,3}\$?)[1-9][0-9]*")


def stretched(sheet: Sheet, places: list[Pos], function: str, area: tuple, rows: tuple[int, int],
              other: str | None = None) -> dict | None:
    """Der Bereich area in den Aufrufen von function auf die Zeilen rows gezogen (D2:D10 → D2:D12), in jeder Formel
    in places (eine heruntergezogene Spalte zusammen), ebenso gleich hohe Bereiche daneben im selben Aufruf
    (SUMMEWENN(B2:B10;…;D2:D10)), damit sie zusammenpassen. other: Der Bereich liegt auf diesem anderen Blatt
    (ZÄHLENWENN(Fälle!$D$2:$D$3989;A5)), dann auch die gleich hohen daneben auf demselben Blatt."""
    r1, _, r2, _ = area
    start, end = rows
    if (start, end) == (r1, r2) or start > end:
        return None
    moved = {}
    kind = "range" if other is None else "xrange"

    def where(arg):
        """Der Bereich des Arguments, wenn es einer auf dem gemeinten Blatt ist, sonst None."""
        if arg.kind != kind:
            return None
        if other is None:
            return arg.value
        return arg.value[1] if arg.value[0].casefold() == other.casefold() else None

    def change(item):
        if item.value != function or not any(where(a) == area for a in item.args):
            return None
        args = []
        for arg in item.args:
            got = where(arg)
            if got is not None and (got[0], got[2]) == (r1, r2):
                grown = (start, got[1], end, got[3])
                new = grown if other is None else (arg.value[0], grown)
                moved[arg.value] = new
                arg = formel.Node(kind, new)
            args.append(arg)
        grown = formel.Node("call", item.value, tuple(args))
        changed.append(grown)
        return grown

    from .bereiche import _text

    changes, parts = {}, {}
    changed: list = []
    for pos in places:
        cell = sheet.cells.get(pos)
        if cell is None or not cell.formula:
            return None
        # Erst einmal durchgehen, um die verlängerten Bereiche zu kennen, dann mit ihren $ schreiben.
        if rewritten(cell.formula, change) is None:
            return None
        changed.clear()
        text = rewritten(cell.formula, change, moved)
        if text is None:
            return None
        changes[pos] = text
        parts[pos] = [t for t in (_text(c, {}) for c in changed) if t]
    return formulas(sheet, changes, parts=parts)


_EXACT = {"VLOOKUP": 3, "HLOOKUP": 3, "MATCH": 2}


def exact(sheet: Sheet, places: list[Pos], function: str, table, errors: bool) -> dict | None:
    """SVERWEIS, WVERWEIS und VERGLEICH, die ungefähr suchen, in places auf genaue Suche gestellt (FALSCH, bei
    VERGLEICH 0), nur die Aufrufe auf table. errors: Darf danach #NV erscheinen, weil ein Schlüssel wirklich fehlt?"""
    index = _EXACT.get(function)
    if index is None:
        return None
    calc, _ = sheet.evaluator()
    flag = formel.Node("num", 0.0) if function == "MATCH" else formel.Node("bool", False)

    def change(call):
        if call.value != function or len(call.args) < index:
            return None
        try:
            if calc.area_of(call.args[1]) != table:
                return None
        except formel.Unknown:
            return None
        args = list(call.args)
        if len(args) == index:
            args.append(flag)
        else:
            args[index] = flag
        return formel.Node("call", function, tuple(args))

    changes = {}
    for pos in places:
        cell = sheet.cells.get(pos)
        text = rewritten(cell.formula, change) if cell is not None and cell.formula else None
        if text is None:
            return None
        changes[pos] = text
    return formulas(sheet, changes, errors)


# --- Mehrere zusammen ---


def merged(fixes: list[dict]) -> dict:
    """Mehrere Korrekturen als eine, so dass das Seitenfenster sie zusammen einsetzt und zusammen zurücknimmt."""
    if len(fixes) == 1:
        return dict(fixes[0])
    cells, was, form, unmerge, general, decimals = {}, {}, {}, [], [], []
    for item in fixes:
        decimals += [a for a in item.get("decimals", ()) if a not in decimals]
        cells |= item["cells"]
        was |= item["was"]
        form |= item.get("format", {})
        unmerge += [a for a in item.get("unmerge", ()) if a not in unmerge]
        general += [a for a in item.get("general", ()) if a not in general]
    targets = sorted((formel.position(a) for a in cells), key=lambda p: (p[1], p[0]))
    found = {"range": _area(targets), "cells": cells, "was": was, "german": T(f"{len(fixes)} Korrekturen", f"{len(fixes)} fixes"),
             "count": len(fixes),
             "parts": [item["german"] for item in fixes]}
    if form:
        found["format"] = form
    if unmerge:
        found["unmerge"] = unmerge
    if decimals:
        found["decimals"] = decimals
    if general:
        found["general"] = general
    return found


def applied(cells: dict[Pos, Cell], fixes: list[dict]) -> dict[Pos, Cell]:
    """Die Zellen, wie sie nach dem Einsetzen aussehen; die Werte der Formeln rechnet refilled() neu."""
    new = Cells(cells, cut_of(cells))
    for item in fixes:
        for where, put in item["cells"].items():
            pos = formel.position(where)
            if put == "":
                new.pop(pos, None)  # cleared(): die Zelle ist danach leer
            elif isinstance(put, str) and put.startswith("="):
                new[pos] = Cell("", put[1:])
            elif isinstance(put, str):
                new[pos] = Cell(put, None, kept_text(put))  # kategorien.py: der Text, wie er sonst geschrieben ist
            else:
                like = item.get("format", {}).get(where)
                dated = like is not None and _DATE_TEXT.fullmatch((cells.get(formel.position(like)) or Cell("")).value or "")
                new[pos] = Cell(_date_text(put, False) if dated else number_text(put))
    return new


_DATE_TEXT = re.compile(r"\d{2}\.\d{2}\.\d{4}")


def _shown(value, like: str) -> str:
    """Ein Ergebnis so, wie das Seitenfenster es aus Excel lesen würde: im Stil des Werts, der vorher dastand."""
    if isinstance(value, formel.Err):
        return value.code
    if isinstance(value, bool):
        return "WAHR" if value else "FALSCH"
    if isinstance(value, float):
        if _DATE_TEXT.fullmatch(like or ""):
            return _date_text(value, False)
        if (like or "").endswith(" %"):
            return f"{number_text(round(value * 100, 10))} %"
        return number_text(value)
    return "" if value is None else str(value)


def refilled(cells: dict[Pos, Cell], plan, name: str, book: dict, before: dict[Pos, Cell]) -> dict[Pos, Cell]:
    """Die Werte aller Formeln neu gerechnet, wie Excel es nach dem Einsetzen tut. Was gleich bleibt, behält den Text
    von vorher (before), damit nur die Folgen des Einsetzens anders aussehen."""
    sheet = Sheet(cells, plan, name, book)
    chain = {pos for pos, cell in cells.items() if cell.formula}
    _, now = sheet.evaluator(chain=chain)
    old_sheet = Sheet(before, plan, name, book)
    _, old = old_sheet.evaluator(chain={pos for pos, cell in before.items() if cell.formula})
    out = Cells(cells, cut_of(cells))
    for pos in chain:
        try:
            value = now(pos)
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, IndexError, RecursionError):
            continue
        prior = before.get(pos)
        try:
            same = prior is not None and prior.formula and old(pos) == value
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, IndexError, RecursionError):
            same = False
        out[pos] = Cell(prior.value if same else _shown(value, prior.value if prior else ""), cells[pos].formula)
    return out
