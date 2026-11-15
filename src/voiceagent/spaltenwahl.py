"""Eine Kennzahl summiert die falsche Spalte ihrer Tabelle.

Ohne KI: „Gesamtumsatz gebucht (EUR)“ steht neben =SUMME(tblRechnungen[UmsatzEUR]). Die Tabelle hat aber auch die
Spalte „UmsatzGebucht“, und deren Kopf trifft mehr Wörter des Namens (Umsatz und gebucht) als der gelesene (nur
Umsatz). Dann hat die Formel die Nachbarspalte erwischt. Einheiten (EUR, Tsd.) und Klammern zählen nicht; es muss
genau eine Spalte geben, die mehr trifft, und die gelesene darf kein Wort treffen, das die andere nicht auch trifft.
"""

from __future__ import annotations

import re

from . import einsetzen, formel, woerter
from .analyse import Finding, Sheet, _fmt, address, calc_value
from .sprache import T
from .xlsx import column_letters

Pos = tuple[int, int]

_AGGREGATES = {"SUM", "AVERAGE", "MIN", "MAX", "SUMIFS", "SUMIF", "AVERAGEIFS", "AVERAGEIF", "SUBTOTAL", "AGGREGATE"}
_UNITS = {"eur", "euro", "usd", "chf", "gbp", "tsd", "mio", "mrd", "prozent", "betrag", "wert", "summe", "gesamt",
          "total", "anzahl", "in", "und", "der", "die", "das", "des", "nach", "pro", "je", "von", "mit", "ohne"}
MAX_FOUND = 3


def _words(text: str) -> set[str]:
    text = woerter.spaced(re.sub(r"\([^)]*\)", " ", text))  # UmsatzGebucht → Umsatz Gebucht
    return {woerter.fold(w) for w in re.findall(r"[^\W\d_]+", text.casefold()) if len(w) >= 4 and w not in _UNITS}


def _hits(head: set[str], label: set[str]) -> set[str]:
    """Die Wörter des Kopfs, die im Namen vorkommen, auch in Zusammensetzungen (Gesamtumsatz) und gebeugt."""
    return {w for w in head if any(w in v or (len(w) >= 5 and v.startswith(w[:-1]) and len(v) - len(w) <= 2)
                                   for v in label)}


def _label(sheet: Sheet, pos: Pos) -> str:
    """Der Name einer Kennzahl links in ihrer Zeile. Stehen dort mehrere Werte (eine Datenzeile), gibt es keinen."""
    texts = [sheet.cells[(pos[0], c)].value for c in range(1, pos[1]) if (pos[0], c) in sheet.cells
             and sheet.cells[(pos[0], c)].formula is None and isinstance(sheet.cells[(pos[0], c)].value, str)
             and str(sheet.cells[(pos[0], c)].value).strip()]
    return " ".join(texts) if len(texts) <= 2 else ""


def _head(sheet: Sheet, row: int, col: int) -> str | None:
    cell = sheet.cells.get((row, col))
    if cell is None or cell.formula or not isinstance(cell.value, str) or not cell.value.strip():
        return None
    return cell.value.strip()


def _numeric(sheet: Sheet, col: int, r1: int, r2: int) -> bool:
    seen = 0
    for r in range(r1, min(r2, r1 + 50) + 1):
        value = sheet.inputs.get((r, col))
        if isinstance(value, bool):
            return False
        if isinstance(value, (int, float)):
            seen += 1
    return seen >= 3


def _summed(node):
    """Die Bereiche über eine Spalte, die eine Kennzahl zusammenfasst (bei SUMMEWENNS nur der Summenbereich). Nur,
    wenn die Formel genau das ist (höchstens in WENNFEHLER oder RUNDEN): Teilt sie durch eine Summe (gewichteter
    Preis = SUMMENPRODUKT/SUMME(Packungen)), ist das ein Nenner, kein Name der Kennzahl."""
    while node is not None and node.kind == "call" and node.value in ("IFERROR", "IFNA", "ROUND", "ABS") and node.args:
        node = node.args[0]
    for call in [node] if node is not None and node.kind == "call" else []:
        if call.value not in _AGGREGATES or not call.args:
            continue
        if call.value in ("SUMIFS", "AVERAGEIFS"):
            args = call.args[:1]
        elif call.value in ("SUMIF", "AVERAGEIF"):
            args = call.args[2:3] or call.args[:1]
        elif call.value in ("SUBTOTAL", "AGGREGATE"):
            args = call.args[1:]
        else:
            args = call.args
        for arg in args:
            if isinstance(arg, formel.Node) and arg.kind == "range":
                r1, c1, r2, c2 = arg.value
                if c1 == c2 and r2 - r1 >= 4 and r2 < formel.MAX_ROW:
                    yield arg.value


def _replaced(sheet: Sheet, formula: str, area, old_head: str, new_head: str, target: int) -> str | None:
    text = formula.lstrip("=")
    if f"[{old_head}]" in text:
        return text.replace(f"[{old_head}]", f"[{new_head}]")
    r1, c1, r2, _ = area
    old, new = column_letters(c1), column_letters(target)
    pattern = rf"(?<![A-Za-z])(\$?){old}(\$?){r1}:(\$?){old}(\$?){r2}(?!\d)"
    if not re.search(pattern, text):
        return None
    return re.sub(pattern, rf"\g<1>{new}\g<2>{r1}:\g<3>{new}\g<4>{r2}", text)


def checks(sheet: Sheet, before: list[Finding] = ()) -> list[Finding]:
    found: list[Finding] = []
    # Was ein anderer Hinweis zur selben Zelle schon sagt (die Spalte zum Namen in einer Tabelle), bleibt dort.
    told = {f.key.partition(":")[2] for f in before}
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or len(found) >= MAX_FOUND or address(pos) in told:
            continue
        label = _words(_label(sheet, pos))
        if not label:
            continue
        for area in _summed(node):
            r1, col, r2, _ = area
            head = _head(sheet, r1 - 1, col)
            if head is None:
                continue
            used = _hits(_words(head), label)
            # Die übrigen Spalten derselben Kopfzeile, so weit sie ohne Lücke reicht.
            others = []
            for step in (-1, 1):
                c = col + step
                while c >= 1 and (name := _head(sheet, r1 - 1, c)) is not None:
                    others.append((c, name))
                    c += step
            better = [(c, name, hits) for c, name in others
                      if (hits := _hits(_words(name), label)) > used and _numeric(sheet, c, r1, r2)]
            if len(better) != 1 or not _numeric(sheet, col, r1, r2):
                continue
            target, name, hits = better[0]
            fix = None
            text = _replaced(sheet, sheet.cells[pos].formula or "", area, head, name, target)
            if text is not None:
                fix = einsetzen.formulas(sheet, {pos: text})
            if fix is None:
                continue
            calc, _ = sheet.evaluator()
            try:
                new = calc_value(calc.run(formel.parse(text)))
            except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, KeyError, IndexError):
                continue
            old = sheet.inputs.get(pos)
            if not isinstance(new, (int, float)) or not isinstance(old, (int, float)) or abs(new - old) < 1e-9:
                continue
            extra = ", ".join(sorted(hits - used))
            letters = column_letters(col)
            label_text = _label(sheet, pos).strip()
            span = f"{letters}{r1}:{letters}{r2}"
            other = f"{column_letters(target)}{r1}:{column_letters(target)}{r2}"
            found.append(Finding(
                f"colname:{address(pos)}",
                T(f"{address(pos)} („{label_text}“) fasst die Spalte „{head}“ zusammen "
                  f"({span}). Der Name sagt aber auch „{extra}“, und dazu passt die Spalte "
                  f"„{name}“ ({other}). Mit ihr wäre {address(pos)} "
                  f"{_fmt(new)} statt {_fmt(old)}.",
                  f"{address(pos)} (“{label_text}”) summarizes the column “{head}” ({span}). But the name also "
                  f"says “{extra}”, and the column “{name}” ({other}) matches that. With it, {address(pos)} would "
                  f"be {_fmt(new)} instead of {_fmt(old)}."),
                "fehler", T(f"Nimm die Spalte „{name}“.", f"Use the column “{name}”."), fix=fix))
            break
    return found
