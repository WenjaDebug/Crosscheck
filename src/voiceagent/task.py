"""Was Gegenprobe über ein ganzes Blatt selbst nachrechnet: Fehlerwerte, Summen mit Lücke, feste Zahlen in einer
rechnenden Spalte, Rundung, die nicht aufgeht, und feste Zahlen in vielen Formeln (Fact). material() gibt dazu das
Blatt als Text mit Formeln und Werten."""

from __future__ import annotations

import math
import re
from collections import Counter
from dataclasses import dataclass, field

from . import formel, woerter
from .workbook import Layout, _as_number, readable
from .xlsx import Cell, column_letters, number_text

# So viel vom Blatt kommt in den Text: Zeilen, Zellen je Zeile, Zeichen.
MAX_ROWS = 60
MAX_CELLS = 20
MAX_GRID = 6000


# Fehlerwerte in Excel, deutsch und englisch.
ERRORS = {
    "#DIV/0!", "#NV", "#N/A", "#WERT!", "#VALUE!", "#BEZUG!", "#REF!", "#NAME?", "#ZAHL!", "#NUM!",
    "#NULL!", "#ÜBERLAUF!", "#SPILL!", "#KALK!", "#CALC!",
}
# Spalten, deren Werte ganze Dinge zählen: Da ist Runden eine fachliche Frage.
_COUNTING = re.compile(r"anzahl|sitze|stück|stk|personen|plätze|mandat|betten|fälle|patient|count|seats|menge|nombre|lits|letti|beds|\bcas\b|\bcasi\b", re.I)
_SUM = re.compile(r"^=?\s*(?:SUM\(|SUBTOTAL\(\s*(?:9|109)\s*,)\s*\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)\s*\)\s*$", re.I)
_TOTALS = ("summe", "total", "gesamt", "zusammen", "sum")
_LITERAL = re.compile(r"(?<![A-Za-z0-9_$.:])\d+(?:\.\d+)?(?![A-Za-z0-9_(:])")
_STRING = re.compile(r'("(?:[^"]|"")*")')
# Eine Zahl für sich in einem Text („Sitze (577)“), nicht die Ziffern eines Codes („D002“, „PAT-0042“).
_IN_TEXT = re.compile(r"(?<![\w.,-])\d+(?:[.,]\d+)?(?![\w])")


@dataclass(frozen=True)
class Fact:
    """Etwas, das Gegenprobe selbst nachgerechnet hat. ask: die Frage dazu, als Tipp im Prüfer."""

    key: str
    text: str
    ask: str = ""
    weight: int = 0


@dataclass
class Material:
    """Das Blatt, wie der Prüfer es zusammenfasst."""

    book: str
    sheet: str
    overview: str
    grid: str
    facts: list[Fact] = field(default_factory=list)
    headers: list[str] = field(default_factory=list)


def _number(value: str) -> float | None:
    return _as_number(value) if value and not value.startswith("#") else None


def _round(value: float) -> int:
    """Wie RUNDEN in Excel: 0,5 wird aufgerundet (Python rundet 0,5 auf die gerade Zahl)."""
    return int(math.copysign(math.floor(abs(value) + 0.5), value))


def _text(value: float) -> str:
    return number_text(round(value, 6))


def _formula(cell: Cell, pos: tuple[int, int]) -> str:
    """Die Formel, wie sie in der deutschen Excel steht, mit Zelladressen."""
    return readable(cell.formula or "", pos, Layout())


def grid(cells: dict[tuple[int, int], Cell]) -> str:
    """Das Blatt zeilenweise: „Zeile 6: B6 RN | C6 33,2 | D6 =0,332 * 577 → 191,564“."""
    rows: dict[int, list[int]] = {}
    for (row, col), cell in cells.items():
        if cell.value or cell.formula is not None:
            rows.setdefault(row, []).append(col)
    lines = []
    for row in sorted(rows)[:MAX_ROWS]:
        parts = []
        for col in sorted(rows[row])[:MAX_CELLS]:
            cell, address = cells[(row, col)], f"{column_letters(col)}{row}"
            if cell.formula is not None:
                parts.append(f"{address} {_formula(cell, (row, col))} → {cell.value or '(leer)'}")
            else:
                parts.append(f"{address} {cell.value}")
        if len(rows[row]) > MAX_CELLS:
            parts.append(f"… {len(rows[row]) - MAX_CELLS} weitere Zellen")
        lines.append(f"Zeile {row}: " + " | ".join(parts))
    if len(rows) > MAX_ROWS:
        lines.append(f"… und {len(rows) - MAX_ROWS} weitere Zeilen")
    text = "\n".join(lines)
    return text if len(text) <= MAX_GRID else text[: MAX_GRID - 1].rstrip() + "…"


def _names(rows: list[int], plan: Layout) -> str:
    names = [plan.names.get(r, f"Zeile {r}") for r in rows]
    return names[0] if len(names) == 1 else ", ".join(names[:-1]) + f" und {names[-1]}"


def facts(cells: dict[tuple[int, int], Cell], plan: Layout) -> list[Fact]:
    """Was sich nachrechnen lässt: Fehlerwerte, Summen, Rundung, eingetippte Zahlen, fehlende Formeln."""
    found: list[Fact] = []
    found += _error_facts(cells)
    if plan.header_row is not None:
        rows = sorted(r for r in plan.names if plan.header_row < r <= plan.last_row)
        # Die Summenzeile gehört nicht zu den Zahlen, die zusammen gerundet aufgehen sollen.
        data = [r for r in rows if not (plan.names[r].strip().lower().startswith(_TOTALS) or woerter.starts_with(plan.names[r], woerter.TOTAL))]
        # In einem Zeilenmodell (Finanzplan) ist eine Spalte ein Jahr, kein Merkmal: Ihre Zahlen zählt niemand
        # zusammen, und Annahmen in Prozent stehen dort mit Absicht als Zahl (analyse._typed_in_rows sagt es je Zeile).
        for col in sorted(plan.headers):
            if col != plan.name_column and not plan.across:
                found += _column_facts(cells, plan, data, col)
        found += _sum_facts(cells, plan, rows)
        found += _row_value_facts(cells, plan, rows)
    found += _constant_facts(cells, plan)
    return found


def _error_facts(cells) -> list[Fact]:
    result = []
    for (row, col), cell in sorted(cells.items()):
        if cell.value.strip().upper() in ERRORS:
            address = f"{column_letters(col)}{row}"
            formula = f" (Formel {_formula(cell, (row, col))})" if cell.formula else ""
            result.append(Fact(
                f"error:{address}", f"{address} zeigt den Fehler {cell.value}{formula}.",
                f"In {address} steht {cell.value}. Weißt du schon, woran das liegt?", 100,
            ))
    return result[:3]


def _column_facts(cells, plan: Layout, rows: list[int], col: int) -> list[Fact]:
    header, letters = plan.headers[col], column_letters(col)
    values = {r: _number(cells[(r, col)].value) for r in rows if (r, col) in cells}
    numbers = {r: v for r, v in values.items() if v is not None}
    result = []
    if len(numbers) >= 2:
        total = sum(numbers.values())
        span = f"{letters}{min(numbers)}:{letters}{max(numbers)}"
        text = f"{header} ({span}): {len(numbers)} Zahlen von {_text(min(numbers.values()))} bis {_text(max(numbers.values()))}, zusammen {_text(total)}"
        fractions = any(abs(v - round(v)) > 1e-9 for v in numbers.values())
        rounded = sum(_round(v) for v in numbers.values())
        ask = ""
        from .bereiche import _NOT_SUMMED, _words

        # Mittelwerte (Mittlere Liegedauer, Quote) zählt niemand zusammen: Ihr Runden ändert keine Summe.
        averaged = bool(set(_words(header)) & _NOT_SUMMED)
        # Nur bei Dingen, die man ganz zählt (Sitze, Betten, Stück), in einer kleinen Aufstellung: Ein BMI oder ein
        # Anteil mit Kommastellen wird nicht auf ganze gerundet, und 2000 Zeilen mit je 0,5 rundet niemand einzeln.
        listing = len(numbers) <= 50 and len(set(numbers.values())) >= 2
        # „Stückpreis“, „Kosten je Fall“: ein Betrag je Stück, nichts, was man ganz zählt.
        money = bool(re.search(r"preis|kosten|betrag|wert|price|cost|prix|prezzo|€|eur\b|chf", header, re.I))
        if fractions and listing and _COUNTING.search(header) and not money and rounded != _round(total) and not averaged:
            text += f"; jede Zahl einzeln auf ganze gerundet ergibt zusammen {rounded}, nicht {_text(total)}"
            ask = f"Einzeln gerundet ergeben deine Werte in {header} zusammen {rounded} statt {_text(total)}. Wie gehst du mit dem Unterschied um?"
        result.append(Fact(f"round:{header}" if ask else f"sum:{header}", text + ".", ask, 60 if ask else 0))
    # Formeln in manchen Zeilen, in anderen nur eine Zahl: Ist das so gewollt?
    with_formula = [r for r in rows if (r, col) in cells and cells[(r, col)].formula is not None]
    computed = set(with_formula)
    typed = [r for r in rows if r not in computed and values.get(r) is not None]
    empty = [r for r in rows if r not in computed and not (cells.get((r, col)) and cells[(r, col)].value)]
    if typed and len(with_formula) >= 2:
        typed = _not_inputs(cells, col, with_formula, typed, rows[0] if rows else 0)
    if len(with_formula) >= 2 and typed:
        at = [f"{letters}{r} ({plan.word(r)})" if r in plan.names else f"{letters}{r}" for r in typed[:3]]
        at = at[0] if len(at) == 1 else ", ".join(at[:-1]) + f" und {at[-1]}"
        more = f" und {len(typed) - 3} weiteren" if len(typed) > 3 else ""
        result.append(Fact(
            f"typed:{header}", f"In {header} rechnen {len(with_formula)} Zeilen mit Formel, bei {at}{more} steht nur eine Zahl.",
            f"Bei {_names(typed, plan)} steht in {header} eine feste Zahl statt einer Formel. Gehört das so?", 50,
        ))
    if with_formula and empty:
        result.append(Fact(f"empty:{header}", f"In {header} ist bei {_names(empty, plan)} noch nichts eingetragen (vielleicht noch unfertig)."))
    return result


# Zahlen in einer Formel in Z1S1-Schreibweise, ohne die Bezüge (Z[-1]S[0], Z12S2) und Namen wie LOG10.
_CONSTANT = re.compile(r"(Z\[-?\d+\]|Z\d+|S\[-?\d+\]|S\d+)|(?<![\w.])\d+(?:\.\d+)?(?:E[+-]?\d+)?")


def _not_inputs(cells, col: int, with_formula: list[int], typed: list[int], first: int = 0) -> list[int]:
    """Die festen Zahlen, bei denen eine Formel fehlen könnte: Rechnen die Formeln der Spalte alle gleich (Anteil mal
    Sitze), fällt eine Zahl dazwischen auf. Eine Aufstellung wie Umsatz, Kosten, Gewinn = Umsatz - Kosten rechnet
    dagegen jede Zeile anders, und was die Formeln der Spalte lesen, ist eine Eingabe, keine vergessene Formel.
    first: die erste Datenzeile. Liest die Formel die Zeile darüber (km Anfang = km Ende der Fahrt davor), ist die
    Zahl in der ersten Zeile der Anfang der Kette: Dort gibt es keine Zeile davor."""
    from .workbook import pattern

    # 0,332*577 und 0,28*577 rechnen gleich, nur mit anderem Anteil.
    shapes = Counter(_CONSTANT.sub(lambda m: m.group(1) or "#", pattern(cells[(r, col)].formula or "", (r, col)))
                     for r in with_formula)
    shape, same = shapes.most_common(1)[0]
    if same < 2 or same * 2 < len(with_formula):
        return []
    sample = next(r for r in with_formula
                  if _CONSTANT.sub(lambda m: m.group(1) or "#", pattern(cells[(r, col)].formula or "", (r, col))) == shape)
    try:
        refs, _, _ = formel.references(formel.parse(cells[(sample, col)].formula or ""))
        above = min((row - sample for row, _ in refs), default=0)
    except formel.Unknown:
        above = 0
    typed = [r for r in typed if r + above >= first]
    # Einzeln gelesen (=B2-B3) ist eine Zahl eine Eingabe; in einer Summe über die Spalte nicht.
    read: set[int] = set()
    for r in with_formula:
        try:
            refs, _, _ = formel.references(formel.parse(cells[(r, col)].formula or ""))
        except formel.Unknown:
            continue
        read |= {row for row, c in refs if c == col}
    return [r for r in typed if r not in read]


def _section(cells, plan: Layout, top: int, bottom: int) -> tuple[int, int]:
    """Die Zeilen des Abschnitts um top bis bottom: bis zur nächsten Zeile ohne eine Zahl oder Formel in den Spalten
    der Tabelle (eine Leerzeile oder eine Zwischenüberschrift)."""
    values = [c for c in plan.headers if c != plan.name_column]

    def filled(r: int) -> bool:
        return any((cell := cells.get((r, c))) is not None and (cell.formula or _number(cell.value) is not None) for c in values)

    start, end = top, bottom
    while start - 1 > (plan.header_row or 0) and filled(start - 1):
        start -= 1
    while end + 1 <= plan.last_row and filled(end + 1):
        end += 1
    return start, end


def _read_rows(cells, col: int, top: int, bottom: int, start: int, end: int) -> set[int]:
    """Die Zeilen start bis end der Spalte col, die eine der Formeln in top bis bottom liest, auch über eine andere:
    EBIT = SUMME(EBITDA; Abschreibungen), EBITDA = SUMME(Bruttogewinn bis F&E), Bruttogewinn = Umsatz + Kosten."""
    read: set[int] = set()
    todo = list(range(top, bottom + 1))
    seen: set[int] = set()
    while todo:
        r = todo.pop()
        cell = cells.get((r, col))
        if r in seen or cell is None or not cell.formula:
            continue
        seen.add(r)
        try:
            singles, areas, _ = formel.references(formel.parse(cell.formula))
        except (formel.Unknown, ValueError):
            continue
        here = {p[0] for p in singles if p[1] == col}
        here |= {x for r1, c1, r2, c2 in areas if c1 <= col <= c2 for x in range(max(r1, start), min(r2, end) + 1)}
        here = {x for x in here if start <= x <= end}
        read |= here
        todo += [x for x in here if x not in seen]
    return read


def _column_sums(cells) -> dict[int, list[tuple[int, int, int]]]:
    """Je Spalte die SUMMEN (und TEILERGEBNIS(9; …)) über einen Bereich derselben Spalte: (Zeile, erste, letzte)."""
    sums: dict[int, list[tuple[int, int, int]]] = {}
    for (row, col), cell in sorted(cells.items()):
        match = _SUM.match(cell.formula or "")
        if not match or match.group(1).upper() != match.group(3).upper() or match.group(1).upper() != column_letters(col):
            continue
        sums.setdefault(col, []).append((row, int(match.group(2)), int(match.group(4))))
    return sums


def _groups_of(siblings: list[tuple[int, int, int]], row: int, top: int, bottom: int) -> set[int]:
    """Die Zeilen der anderen Gruppen in einem Bericht mit Zwischensummen („Summe Nord“ über Nord, „Summe Süd“ über
    Süd): Was eine andere Zwischensumme derselben Spalte zusammenfasst, fehlt dieser nicht. Gruppe heißt der Bereich
    samt der Zeile der Zwischensumme (und was dazwischen liegt). Eine Summe, deren Bereich diesen überdeckt (die
    Gesamtsumme), ist keine andere Gruppe."""
    out: set[int] = set()
    for other, first, last in siblings:
        if other == row or first <= bottom and last >= top or top <= other <= bottom:
            continue
        out |= set(range(min(first, other), max(last, other) + 1))
    return out


def sum_gaps(cells, plan: Layout, rows: list[int]):
    """Je SUMME über einen Bereich ihrer eigenen Spalte: (Zeile, Spalte, erste, letzte Zeile des Bereichs, die Zeilen
    der Tabelle mit einer Zahl, die sie auslässt). Leer, wenn sie keine auslässt."""
    sums = _column_sums(cells)
    for row, col, top, bottom in sorted((r, c, t, b) for c, found in sums.items() for r, t, b in found):
        yield _sum_gap(cells, plan, rows, col, sums[col], row, top, bottom)


def _sum_gap(cells, plan: Layout, rows: list[int], col: int, siblings, row: int, top: int, bottom: int):
    counted = [r for r in rows if (r, col) in cells and _number(cells[(r, col)].value) is not None]
    groups = _groups_of(siblings, row, top, bottom) if len(siblings) > 1 else set()
    # Eine Summe unter ihren Zahlen zählt nichts darunter (dort steht etwas anderes), eine darüber nichts darüber.
    # Eine andere SUMME über Zeilen, die diese schon zählt (die Zeile „Summe“ über B5:B16, darunter
    # Gesamtumsatz =SUMME(B5:B16)), fasst nur dieselben Zahlen zusammen: Sie mitzuzählen, zählte doppelt.
    inner = {other for other, first, last in siblings if other != row and top <= first and last <= bottom}
    left_out = [r for r in counted if not top <= r <= bottom and r != row and r not in groups and r not in inner
                and not (bottom < row < r or r < row < top)]
    if left_out:
        # Über einer Leerzeile oder einer Zwischenüberschrift („Cashflow“) beginnt ein anderer Abschnitt, und was
        # eine der summierten Zellen schon liest (EBITDA = SUMME(Bruttogewinn bis F&E), Bruttogewinn = Umsatz +
        # Kosten), zählt mit.
        start, end = _section(cells, plan, top, bottom)
        read = _read_rows(cells, col, top, bottom, start, end)
        left_out = [r for r in left_out if start <= r <= end and r not in read]
    return row, col, top, bottom, (left_out if plan.in_table(top, col) else [])


def _sum_facts(cells, plan: Layout, rows: list[int]) -> list[Fact]:
    """Eine SUMME, die eine Zeile der Tabelle auslässt, oder anders rechnet als nachgerechnet."""
    result = []
    for row, col, top, bottom, left_out in sum_gaps(cells, plan, rows):
        cell = cells[(row, col)]
        address, formula = f"{column_letters(col)}{row}", _formula(cell, (row, col))
        # In einem Zeilenmodell sagt analyse (zeilenmodell.py) es einmal für die ganze Zeile, mit der Korrektur.
        if left_out and not plan.across:
            result.append(Fact(
                f"gap:{address}", f"Die Summe in {address} ({formula}) zählt {_names(left_out, plan)} nicht mit.",
                f"Die Summe in {address} zählt {_names(left_out, plan)} nicht mit. Ist das Absicht?", 90,
            ))
        shown = _number(cell.value)
        # TEILERGEBNIS lässt andere Teilergebnisse im Bereich und weggefilterte Zeilen weg: Das zählt formel.py nach.
        if shown is None or "SUBTOTAL" in (cell.formula or "").upper():
            continue
        total = sum(_number(cells[(r, col)].value) or 0.0 for r in range(top, bottom + 1) if (r, col) in cells)
        if abs(total - shown) > 1e-6 * max(1.0, abs(total)):
            result.append(Fact(
                f"sum:{address}", f"Die Summe in {address} zeigt {cell.value}, nachgerechnet ergeben {letters_range(col, top, bottom)} {_text(total)} "
                "(vielleicht stehen Zahlen als Text da).",
                f"Die Summe in {address} zeigt {cell.value}, nachgerechnet komme ich auf {_text(total)}. Woran liegt das?", 80,
            ))
    return result[:3]


def letters_range(col: int, top: int, bottom: int) -> str:
    return f"{column_letters(col)}{top}:{column_letters(col)}{bottom}"


_ROUNDS = {"ROUND", "ROUNDUP", "ROUNDDOWN", "INT", "TRUNC"}


def _rounding_steps(node) -> list[float]:
    """RUNDEN(x*2;0)/2 rundet auf halbe Noten, RUNDEN(x*20;0)/20 auf 5 Rappen: Die 2 und die 20 sind die Rundung,
    keine Zahl, die in eine eigene Zelle gehört."""
    found, stack = [], [node]
    while stack:
        n = stack.pop()
        stack.extend(n.args or [])
        if n.kind != "op" or n.value != "/" or len(n.args) != 2 or n.args[1].kind != "num":
            continue
        inner = n.args[0]
        if inner.kind == "call" and inner.value in _ROUNDS and inner.args:
            times = inner.args[0]
            if times.kind == "op" and times.value == "*" and len(times.args) == 2 and times.args[1].kind == "num" \
                    and times.args[1].value == n.args[1].value:
                found.append(n.args[1].value)
    return found


def _literals(formula: str) -> tuple[float, ...]:
    """Zahlen, die in einer Formel eingetippt sind (nicht in Texten, nicht in Zelladressen, nicht als
    Einstellung wie die Spalte bei SVERWEIS). Je Formel einmal, die Prüfungen fragen mehrmals."""
    return formel.remember("\x00literals", formula, lambda: _literals_of(formula))


def _literals_of(formula: str) -> tuple[float, ...]:
    try:
        # Hochzahlen (^2 im BMI) sind Rechnen, keine Zahl, die in eine eigene Zelle gehört.
        node = formel.parse(formula)
        numbers = formel.literals(node, exponents=False)
        numbers = list(numbers)
        for step in _rounding_steps(node):
            if numbers.count(step) >= 2:
                numbers.remove(step)
                numbers.remove(step)
    except formel.Unknown:
        numbers = []
        for part in _STRING.split(formula):
            if part.startswith('"'):
                continue
            numbers += [float(n) for n in _LITERAL.findall(part)]
    return tuple(n for n in numbers if n not in (0.0, 1.0, 100.0))


def _row_value_facts(cells, plan: Layout, rows: list[int]) -> list[Fact]:
    """„0,332 in D6 ist Wähleranteil(%) aus C6 geteilt durch 100“: abgetippt statt auf die Zelle zu verweisen."""
    result = []
    for col in sorted(plan.headers):
        hits = []
        # Dieselbe Zahl in den Formeln mehrerer Zeilen ist eine feste Zahl der Spalte (das sagt _constant_facts), kein
        # abgetippter Wert: Der wäre in jeder Zeile ein anderer. Sonst trifft eine 2 irgendwo zufällig eine 2 daneben.
        typed = {row: _literals(cells[(row, col)].formula) for row in rows
                 if (row, col) in cells and cells[(row, col)].formula is not None}
        shared = {n for n, count in Counter(n for numbers in typed.values() for n in set(numbers)).items() if count >= 2}
        for row, numbers in typed.items():
            for literal in numbers:
                if literal in shared:
                    continue
                for other in sorted(plan.headers):
                    value = _number(cells[(row, other)].value) if other != col and (row, other) in cells else None
                    if not value:
                        continue
                    for factor, words in ((1, ""), (100, " / 100"), (0.01, " × 100")):
                        if abs(literal * factor - value) < 1e-9 * max(1.0, abs(value)):
                            hits.append((row, literal, other, value, words))
        if not hits:
            continue
        row, literal, other, value, words = hits[0]
        columns = {h[2] for h in hits}
        example = f"{column_letters(col)}{row}: {_text(literal)} = {plan.headers[other]} {_text(value)}{words}"
        every = len({h[0] for h in hits}) >= 2
        result.append(Fact(
            f"copied:{plan.headers[col]}",
            f"In {plan.headers[col]} steht in {'jeder Formel' if every else 'der Formel'} eine abgetippte Zahl aus derselben Zeile "
            f"statt eines Bezugs auf die Zelle ({example}).",
            f"In {plan.headers[col]} tippst du {plan.headers[min(columns)]} als Zahl ab, zum Beispiel {_text(literal)}. "
            "Warum nicht als Bezug auf die Zelle?",
            30,
        ))
    return result


_CALENDAR = {7.0, 12.0, 24.0, 60.0, 365.0, 365.25, 366.0}


def _per_unit(text: str, literal: float) -> bool:
    """„HAI je 1.000 PT“, „pro 1000 Einwohner“, „per 100k“: die Zahl als Bezugsgröße im Namen."""
    whole = int(literal)
    if whole != literal:
        return False
    spelled = {str(whole), f"{whole:,}".replace(",", "."), f"{whole:,}".replace(",", "'"), f"{whole:,}"}
    if whole % 1000 == 0:
        spelled.add(f"{whole // 1000}k")
    low = str(text).casefold()
    return any(re.search(rf"(?:je|pro|per|/)\s*{re.escape(s)}(?![\d.,'])", low) for s in spelled)


_CHECK_TEXT = re.compile(r'"\s*(ok|okay|stimmt|passt|fehler|error|prüfen|check|falsch)\s*"', re.I)
_DATE_PART = re.compile(r"\b(?:MONTH|MONAT|DAY|TAG|WEEKDAY|WOCHENTAG|HOUR|STUNDE)\([^()]*\)\s*(?:[<>]=?|=|/|-|\+)\s*(\d+(?:\.\d+)?)", re.I)


def _tolerance(formula: str | None, literal: float) -> bool:
    """=WENN(ABS(C9-…)<0,001;"OK";"FEHLER"): die Toleranz einer Prüfung, keine Annahme des Modells."""
    text = formula or ""
    if not _CHECK_TEXT.search(text):
        return False
    return any(abs(float(n) - literal) < 1e-12 for n in re.findall(r"[<>]=?\s*(\d+(?:\.\d+)?(?:E-?\d+)?)", text, re.I))


def _date_part(formula: str | None, literal: float) -> bool:
    """MONAT(A16)/3 für das Quartal, MONAT(A16)>9 für das Geschäftsjahr: eine Regel des Kalenders an diesem Datum."""
    return any(abs(float(n) - literal) < 1e-12 for n in _DATE_PART.findall(formula or ""))


def _mathematical(formula: str | None) -> bool:
    text = (formula or "").upper()
    return "SQRT(" in text or "^" in text or "WURZEL(" in text


_ALPHA = (0.05, 0.01, 0.001, 0.1)
_Z = (1.645, 1.96, 2.576, 2.58)
_P_WORDS = re.compile(r"p[- ]?wert|p[- ]?value|signifikan|significan", re.I)
_CI_WORDS = re.compile(r"\bki\b|konfidenz|confidence|\bci\b|intervall|interval", re.I)


def _explained(cells, plan: Layout, literal: float, pos) -> bool:
    """Steht die Zahl im Kopf ihrer Spalte oder im Namen ihrer Zeile, oder ist sie eine Konvention der Statistik an
    ihrem Ort (0,05 neben p-Werten, 1,96 in einem Konfidenzintervall)?"""
    row, col = pos
    names = [plan.headers.get(col, ""), plan.names.get(row, "")]
    left = cells.get((row, col - 1))
    if left is not None and left.formula is None:
        names.append(left.value)
    for name in names:
        if any(abs(float(n.replace(",", ".")) - literal) < 1e-9 for n in _IN_TEXT.findall(name or "")):
            return True
    formula = cells[pos].formula or ""
    if literal in _ALPHA:
        read = {plan.headers.get(c, "") for c in range(1, 200) if re.search(rf"\b{column_letters(c)}\$?\d", formula)}
        return any(_P_WORDS.search(h or "") for h in read | set(names))
    if literal in _Z:
        return any(_CI_WORDS.search(n or "") for n in names)
    return False


_FORECASTS = {"FORECAST", "FORECAST.LINEAR", "TREND", "GROWTH"}


def _forecast_point(cells, formula: str | None, literal: float) -> bool:
    """Steht die Zahl in der Formel nur als neuer x-Wert einer Prognose (PROGNOSE, TREND, VARIATION), und liegt sie
    auf dem gleichmäßigen Raster der bekannten x-Werte (Monat 1 bis 12, dann 13)?"""
    try:
        node = formel.parse(formula or "")
    except formel.Unknown:
        return False
    hits = 0

    def grid(area) -> bool:
        if area.kind != "range":
            return False
        r1, c1, r2, c2 = area.value
        xs = [_number(str(cells[(r, c)].value)) if (r, c) in cells else None
              for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
        if len(xs) < 3 or any(x is None for x in xs):
            return False
        step = xs[1] - xs[0]
        if step == 0 or any(abs((b - a) - step) > 1e-9 for a, b in zip(xs, xs[1:])):
            return False
        k = (literal - xs[0]) / step
        return abs(k - round(k)) < 1e-9

    def walk(n) -> None:
        nonlocal hits
        if n.kind == "call" and n.value in _FORECASTS:
            if n.value.startswith("FORECAST") and len(n.args) == 3:
                x, known = n.args[0], n.args[2]
            elif len(n.args) >= 3:
                x, known = n.args[2], n.args[1]
            else:
                x, known = None, None
            if x is not None and x.kind == "num" and x.value == literal and grid(known):
                hits += 1
        for a in getattr(n, "args", None) or ():
            walk(a)

    def months_left(n) -> None:
        # =13-$B$11 mit dem Startmonat 1 bis 12 in B11: Monate bis Jahresende, den Startmonat mitgezählt.
        nonlocal hits
        if literal == 13 and n.kind == "op" and n.value == "-" and len(n.args) == 2 and n.args[0].kind == "num" \
                and n.args[0].value == 13 and n.args[1].kind == "ref" and n.args[1].value in cells:
            month = _number(str(cells[n.args[1].value].value))
            if month is not None and month == int(month) and 1 <= month <= 12:
                hits += 1
        for a in getattr(n, "args", None) or ():
            months_left(a)

    walk(node)
    months_left(node)
    return hits > 0 and formel.literals(node, exponents=False).count(literal) == hits


def _constant_facts(cells, plan: Layout) -> list[Fact]:
    """Eine feste Zahl in mehreren Formeln (577), am besten noch irgendwo auf dem Blatt erklärt („Sitze(577)“)."""
    from .sensitivitaet import axes

    used: Counter = Counter()
    where: dict[float, list[str]] = {}
    # Die Stufen in den Köpfen einer Sensitivität (=$B$15-0,01) gehören in keine eigene Zelle.
    steps = axes(cells)
    for (row, col), cell in sorted(cells.items()):
        if cell.formula is None or (row, col) in steps:
            continue
        for literal in set(_literals(cell.formula)):
            used[literal] += 1
            where.setdefault(literal, []).append(f"{column_letters(col)}{row}")
    result = []
    # Tage im Jahr, Monate, Stunden: Kalenderzahlen gehören in keine eigene Zelle. Ob /365 genau genug ist, sagt
    # zeitraeume.py.
    for literal in _CALENDAR:
        used.pop(literal, None)
    # „je 1.000 PT“ im Kopf: Die 1000 ist die Einheit der Spalte, keine Annahme.
    for literal in list(used):
        if literal >= 10 and any(_per_unit(cell.value, literal) for cell in cells.values() if cell.formula is None and cell.value):
            used.pop(literal)
    # 2 und 4 in einer Formel mit Wurzel oder Hochzahl (Wilson-Intervall, Varianz): Teil der Rechenregel.
    for literal in (2.0, 3.0, 4.0):
        if literal in used and all(_mathematical(cells[formel.position(at)].formula) for at in where[literal]):
            used.pop(literal)
    # Die Schwelle im eigenen Spaltenkopf („Ausreißer (|z|>3)“) oder eine Konvention der Statistik am richtigen
    # Ort (0,05 für p-Werte, 1,96 im Konfidenzintervall): erklärt, keine versteckte Annahme. Ob sie stimmt, sagen
    # statistiktests.py und kopf.py.
    # =PROGNOSE(13;H18:H29;B18:B29) über die Monate 1 bis 12: Die 13 ist die Periode, für die man vorhersagt, keine
    # Annahme. Nur, wenn sie auf dem Raster der bekannten x-Werte liegt. Ebenso =13-B11 mit einem Monat in B11.
    for literal in list(used):
        if all(_forecast_point(cells, cells[formel.position(at)].formula, literal) for at in where[literal]):
            used.pop(literal)
    for literal in list(used):
        if all(_explained(cells, plan, literal, formel.position(at)) for at in where[literal]):
            used.pop(literal)
    for literal in list(used):
        if all(_tolerance(cells[formel.position(at)].formula, literal) or _date_part(cells[formel.position(at)].formula, literal)
               for at in where[literal]):
            used.pop(literal)
    for literal, count in used.most_common(3):
        if count < 2:
            break
        text = f"Die Zahl {_text(literal)} steht fest in {count} Formeln ({', '.join(where[literal][:6])})"
        mentioned = []
        for (row, col), cell in sorted(cells.items()):
            if cell.formula is None and cell.value and any(abs(float(n.replace(",", ".")) - literal) < 1e-9 for n in _IN_TEXT.findall(cell.value)):
                mentioned.append(f"{column_letters(col)}{row} „{cell.value}“")
        if mentioned:
            text += f" und auf dem Blatt in {', '.join(mentioned[:3])}"
        result.append(Fact(f"constant:{_text(literal)}", text + ".", f"Woher kommt die {_text(literal)} in deinen Formeln?", 70))
    return result


def material(tracker) -> Material | None:
    """Das aktive Blatt (workbook.WorkbookTracker) mit allem, was Gegenprobe daran nachrechnet."""
    state = tracker.active_state()
    if state is None:
        return None
    sheet, cells, plan = state
    return Material(tracker.name, sheet, tracker.overview(), grid(cells), facts(cells, plan), list(plan.headers.values()))


