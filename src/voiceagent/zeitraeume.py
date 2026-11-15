"""Alter und Zeiträume, die nicht genau gerechnet sind.

- Ganze Jahre (Alter, Dienstjahre) als (Ende − Beginn)/365 oder /365,25 mit GANZZAHL, KÜRZEN, ABRUNDEN, RUNDEN oder
  QUOTIENT, und als JAHR(Ende) − JAHR(Beginn): Um Geburtstage herum liegt das ein Jahr daneben. Genau ist
  DATEDIF(Beginn;Ende;"Y"). Ganze Monate als (Ende − Beginn)/30: genau ist DATEDIF(Beginn;Ende;"M").
- HEUTE() in einer Tabelle von Fällen mit Eintritt oder Austritt: Das Alter oder die Dauer ändert sich jeden Tag, und
  wer die Mappe später öffnet, bekommt andere Zahlen. Gemeint ist meist das Alter beim Eintritt oder die Dauer bis
  zum Austritt. Steht „bei Eintritt“ schon im Namen, ist es ein Fehler.
- Tage „inklusive“ ohne +1, „Nächte“ oder „exklusive“ mit +1, und „Arbeitstage“ als einfache Differenz mit
  Wochenenden statt NETTOARBEITSTAGE.

Jede Korrektur ist für jede Zeile nachgerechnet und ergibt genau das, was ein Kalender sagt. HEUTE() rechnet der
Prüfer dafür als das heutige Datum. Alles ohne KI.
"""

from __future__ import annotations

import math
from datetime import date, timedelta

from . import einsetzen, formel, woerter
from .analyse import Finding, Sheet, address
from .bereiche import _german_formula, _kept, _paths, _put, _text, _words
from .formel import Node
from .sprache import T
from .statistik import _name
from .statistiktests import _column
from .workbook import FUNCTIONS, pattern
from .xlsx import number_text

Pos = tuple[int, int]

MAX_FINDINGS = 6
MAX_LISTED = 3
_EPOCH = date(1899, 12, 30)
_YEAR = {365.0, 365.25, 365.24, 365.242, 365.2425}
_MONTH = {30.0, 30.4, 30.41, 30.42, 30.44, 30.437, 30.5, 31.0}
_FLOOR = {"INT", "TRUNC", "ROUNDDOWN", "FLOOR", "FLOOR.MATH", "ROUND", "ROUNDUP"}
_TODAY = {"TODAY", "NOW"}
_BIRTH = {"geburt", "geburtsdatum", "geburtstag", "geb", "gebdat", "geburtsdat", "dob", "birth", "birthday",
          "birthdate", "geboren", "jahrgang"} | woerter.BIRTH_WORDS
_START = {"eintritt", "eintrittsdatum", "aufnahme", "aufnahmedatum", "admission", "eintrittstag", "spitaleintritt",
          "diagnose", "diagnosedatum", "op", "operation", "opdatum", "beginn", "therapiebeginn", "einschluss",
          "randomisierung", "behandlungsbeginn"} | woerter.START_WORDS
_END = {"austritt", "austrittsdatum", "entlassung", "entlassungsdatum", "discharge", "austrittstag", "spitalaustritt"} | woerter.END_WORDS
_AGE = {"alter", "age", "lebensalter", "lebensjahre"} | woerter.AGE_WORDS
# „Alter heute“: Dann ist HEUTE() gemeint.
_NOW = {"heute", "aktuell", "aktuelles", "derzeit", "jetzt", "today", "current", "stand"}
_INCLUSIVE = {"inkl", "inklusive", "einschließlich", "einschliesslich", "inclusive", "incl", "inkl."}
_EXCLUSIVE = {"exkl", "exklusive", "ausschließlich", "ausschliesslich", "exclusive", "excl"}
_NIGHTS = {"nächte", "naechte", "übernachtungen", "uebernachtungen", "nights", "nacht"}
_WORKDAYS = {"arbeitstage", "werktage", "arbeitstag", "werktag", "geschäftstage", "workdays", "weekdays", "businessdays",
             "arbeitstagen", "werktagen"}


def _today() -> date:
    return date.today()


def _day(serial: float) -> date:
    return _EPOCH + timedelta(days=math.floor(serial))


def _shown(d: date) -> str:
    return d.strftime("%d.%m.%Y")


def _freeze(node: Node) -> Node:
    """HEUTE() und JETZT() als das heutige Datum, damit sich nachrechnen lässt, was Excel heute zeigt."""
    if node.kind == "call" and node.value in _TODAY and not node.args:
        t = _today()
        return Node("call", "DATE", (Node("num", float(t.year)), Node("num", float(t.month)), Node("num", float(t.day))))
    if not node.args:
        return node
    return Node(node.kind, node.value, tuple(_freeze(a) if isinstance(a, Node) else a for a in node.args))


class _Calc:
    """Der Rechner des Blatts, mit HEUTE() als heutigem Datum."""

    def __init__(self, sheet: Sheet):
        self.calc, _ = sheet.evaluator()

    def run(self, node: Node):
        return self.calc.run(_freeze(node))

    def value(self, node: Node):
        try:
            return self.run(node)
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError, IndexError, KeyError):
            return None


def _number(value) -> bool:
    return isinstance(value, float) and not isinstance(value, bool) and math.isfinite(value)


def _datedif(start: date, end: date, unit: str) -> int | None:
    """DATEDIF wie in Excel: volle Monate, ein Monat ist erst voll, wenn der Tag erreicht ist (29.02. → 01.03.)."""
    if end < start:
        return None
    months = (end.year - start.year) * 12 + end.month - start.month - (end.day < start.day)
    return months // 12 if unit == "Y" else months


def _anniversary(start: date, years: int) -> date:
    """Der Tag, an dem DATEDIF(start; …; "Y") auf years springt: 29.02. in einem Jahr ohne Schalttag ist der 01.03."""
    try:
        return start.replace(year=start.year + years)
    except ValueError:
        return date(start.year + years, 3, 1)


# --- Datumswerte in Formeln ---


def _kind(sheet: Sheet, node: Node) -> str | None:
    """Was ein Datum in der Formel ist: „heute“, „geburt“, „start“, „ende“ oder „datum“, sonst None."""
    if node.kind == "call" and node.value in _TODAY and not node.args:
        return "heute"
    if node.kind != "ref" or not isinstance(node.value, tuple):
        return None
    shown = sheet.shown.get(node.value)
    if not isinstance(shown, str) or formel.serial_of(shown) is None or not _number(sheet.inputs.get(node.value)):
        return None
    words = set(_header_words(sheet, node.value))
    if words & _BIRTH:
        return "geburt"
    if words & _START:
        return "start"
    if words & _END:
        return "ende"
    return "datum"


def _header_words(sheet: Sheet, pos: Pos) -> list[str]:
    r, c = pos
    if sheet.plan.in_table(r, c):
        return _words(sheet.plan.headers.get(c, ""))
    left = sheet.cells.get((r, c - 1))
    return _words(left.value) if left is not None and left.formula is None and isinstance(left.value, str) else []


def _span(sheet: Sheet, node: Node) -> tuple[Node, Node] | None:
    """Ende − Beginn mit zwei Datumswerten → (Beginn, Ende)."""
    if node.kind == "op" and node.value == "-" and len(node.args) == 2:
        end, start = node.args
        if _kind(sheet, end) and _kind(sheet, start) and _kind(sheet, start) != "heute":
            return start, end
    return None


def _whole(sheet: Sheet, node: Node):
    """Ganze Jahre oder Monate von Hand: (Einheit, Beginn, Ende, Art, Teiler) mit Art „teilen“, „runden“ oder
    „jahre“ (JAHR(Ende) − JAHR(Beginn), ohne Teiler)."""
    if node.kind == "call" and node.value == "QUOTIENT" and len(node.args) == 2:
        inner, k = node.args
        span = _span(sheet, inner)
        if span and k.kind == "num":
            unit = "Y" if k.value in _YEAR else "M" if k.value in _MONTH else None
            if unit:
                return unit, *span, "teilen", k.value
    if node.kind == "call" and node.value in _FLOOR and node.args:
        inner, rest = node.args[0], node.args[1:]
        exact = node.value in ("INT",) or (node.value == "FLOOR.MATH" and not rest)
        exact = exact or (node.value in ("TRUNC", "ROUNDDOWN", "ROUND", "ROUNDUP") and
                          (not rest and node.value == "TRUNC" or rest and rest[0].kind == "num" and rest[0].value == 0))
        exact = exact or (node.value == "FLOOR" and len(rest) == 1 and rest[0].kind == "num" and rest[0].value == 1)
        if exact and inner.kind == "op" and inner.value == "/" and len(inner.args) == 2 and inner.args[1].kind == "num":
            span = _span(sheet, inner.args[0])
            k = inner.args[1].value
            unit = "Y" if k in _YEAR else "M" if k in _MONTH else None
            if span and unit:
                form = {"ROUND": "runden", "ROUNDUP": "aufrunden"}.get(node.value, "teilen")
                return unit, *span, form, k
    if node.kind == "op" and node.value == "-" and len(node.args) == 2:
        end, start = node.args
        # Nur fürs Alter: Sonst kann die Differenz der Kalenderjahre gemeint sein (2024 bis 2026 sind zwei).
        if all(a.kind == "call" and a.value == "YEAR" and len(a.args) == 1 for a in (end, start)):
            if _kind(sheet, end.args[0]) and _kind(sheet, start.args[0]) == "geburt":
                return "Y", start.args[0], end.args[0], "jahre", None
    return None


def _cases(sheet: Sheet, r: int) -> dict[str, Pos]:
    """Die Datumsspalten der Tabelle in Zeile r, nach Art: {"start": (r, 4), "ende": (r, 5)}."""
    found = {}
    if sheet.plan.header_row is None:
        return found
    for c in sorted(sheet.plan.headers):
        kind = _kind(sheet, Node("ref", (r, c)))
        if kind in ("start", "ende", "geburt") and kind not in found:
            found[kind] = (r, c)
    return found


def _german(name: str) -> str:
    return FUNCTIONS.get(name, name)


def _rows(rows: list[int]) -> str:
    shown = ", ".join(map(str, rows[:MAX_LISTED]))
    return shown + (T(f" und {len(rows) - MAX_LISTED} weitere", f" and {len(rows) - MAX_LISTED} more")
                    if len(rows) > MAX_LISTED else "")


def _ordinal(n: int) -> str:
    """3 → „3rd“, für „the 3rd birthday“."""
    suffix = "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def _fix(sheet: Sheet, calc: _Calc, run: list[Pos], build, expected: dict[Pos, float]) -> dict | None:
    """Die Korrektur für alle Zellen von run: build(pos) gibt die neue englische Formel, expected den Wert, den sie
    ergeben muss. Nur, wenn jede Zelle genau ihn ergibt."""
    changes = {}
    for pos in run:
        text = build(pos)
        if text is None:
            return None
        try:
            got = calc.run(formel.parse(text))
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError, IndexError, KeyError):
            return None
        if pos in expected and not (_number(got) and abs(got - expected[pos]) < 1e-9):
            return None
        changes[pos] = text
    return einsetzen.formulas(sheet, changes, calc=calc)


def _block(sheet: Sheet, pos: Pos) -> list[Pos]:
    """Die ganze heruntergezogene Spalte um pos, von ihrer ersten Zelle an."""
    shape = pattern(sheet.cells[pos].formula or "", pos)
    top = pos
    while (above := (top[0] - 1, top[1])) in sheet.nodes and sheet.cells[above].formula and \
            pattern(sheet.cells[above].formula, above) == shape:
        top = above
    return _column(sheet, top)


def _tip(fix: dict | None, run: list[Pos], rows: int, calendar: bool = False, question: str = "") -> str:
    """„Korrigiert in H2: …, bis H41 herunterziehen. Für alle 40 Zeilen nachgerechnet.“, oder mit einer Frage vorn
    (question, ohne Fragezeichen): „Ist das Alter bei Eintritt gemeint? Dann in H2: …“."""
    if fix is None:
        return ""
    head = T(f"{question}? Dann", f"{question}? Then") if question else T("Korrigiert", "Corrected")
    check = T(f"Für {'die eine Zeile' if rows == 1 else f'alle {rows} Zeilen'} nachgerechnet" +
              (", genau wie im Kalender." if calendar else "."),
              f"Recalculated for {'the one row' if rows == 1 else f'all {rows} rows'}" +
              (", exactly as the calendar says." if calendar else "."))
    if len(run) > 1:
        return T(f"{head} in {address(run[0])}: {fix['german']}, bis {address(run[-1])} herunterziehen. {check}",
                 f"{head} in {address(run[0])}: {fix['german']}, then fill down to {address(run[-1])}. {check}")
    return f"{head}: {fix['german']}. {check}"


# --- ganze Jahre und Monate ---


def _wholes(sheet: Sheet, calc: _Calc, done: set[Pos]) -> list[Finding]:  # noqa: C901 - je Form ein Satz
    found, seen = [], set()
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in done or pos in seen:
            continue
        match = next(((path, shape) for path, part in _paths(node) if (shape := _whole(sheet, part))), None)
        if match is None:
            continue
        path, (unit, start, end, form, k) = match
        run = _column(sheet, pos)
        seen |= set(run)
        label = set(_name(sheet, pos)) | set(_words(sheet.figure_label(pos)))
        birth = _kind(sheet, start) == "geburt" or bool(label & _AGE)
        wrong, expected, rows, most = [], {}, 0, 0
        for cell in run:
            mine = sheet.nodes.get(cell)
            part = _part(mine, path)
            shape = _whole(sheet, part) if part is not None else None
            if shape is None:
                continue
            a, b, now = calc.value(shape[1]), calc.value(shape[2]), calc.value(part)
            if not (_number(a) and _number(b) and _number(now)):
                continue
            exact = _datedif(_day(a), _day(b), unit)
            if exact is None:
                continue
            rows += 1
            most = max(most, exact)
            expected[cell] = float(exact)
            if abs(now - exact) > 1e-9:
                wrong.append((cell, now, exact, _day(a), _day(b)))
        if not rows:
            continue
        what = T(("das Alter" if birth else "ganze Jahre") if unit == "Y" else "ganze Monate",
                 ("the age" if birth else "whole years") if unit == "Y" else "whole months")
        shown = _german_formula(_text(part_of(node, path), _kept(sheet.cells[pos].formula or "")) or "")
        unit_word = T("Jahre" if unit == "Y" else "Monate", "years" if unit == "Y" else "months")
        divisor = formel_number(k) if k is not None else ""
        if form == "jahre":
            if birth:
                why = T(f"zieht nur die Jahreszahlen ab ({shown}): Vor dem Geburtstag im selben Jahr ist das ein Jahr zu viel",
                        f"only subtracts the year numbers (={shown}): before the birthday in the same year, that is one "
                        "year too many")
            else:
                why = T(f"zieht nur die Jahreszahlen ab ({shown}), so zählt ein Jahr schon, bevor es um ist",
                        f"only subtracts the year numbers (={shown}), so a year counts before it is over")
        elif form == "runden":
            why = T(f"rundet {shown}: Ab einem halben Jahr steht schon das nächste da",
                    f"rounds ={shown}: from half a year on, the next one already shows") if unit == "Y" else \
                T(f"rundet {shown}: Ab einem halben Monat steht schon der nächste da",
                  f"rounds ={shown}: from half a month on, the next one already shows")
        elif form == "aufrunden":
            why = T(f"rundet {shown} auf: Schon einen Tag danach steht {'das nächste Jahr' if unit == 'Y' else 'der nächste Monat'} da",
                    f"rounds ={shown} up: just one day later, the next {'year' if unit == 'Y' else 'month'} already shows")
        elif unit == "Y":
            why = T(f"teilt die Tage durch {divisor} ({shown}). Ein Jahr hat aber 365 oder 366 Tage, so "
                    f"springt {'das Alter' if birth else 'die Zahl'} um den {'Geburtstag' if birth else 'Jahrestag'} "
                    "herum einen oder mehrere Tage zu früh oder zu spät",
                    f"divides the days by {divisor} (={shown}). But a year has 365 or 366 days, so "
                    f"{'the age' if birth else 'the number'} changes one or more days too early or too late around the "
                    f"{'birthday' if birth else 'anniversary'}")
        else:
            why = T(f"teilt die Tage durch {divisor} ({shown}). Ein Monat hat aber 28 bis 31 Tage, so "
                    "liegt es um den Monatstag herum daneben",
                    f"divides the days by {divisor} (={shown}). But a month has 28 to 31 days, so "
                    "it is off around the day of the month")
        if wrong:
            cell, now, exact, a, b = wrong[0]
            jump = _anniversary(a, exact + 1) if now > exact else _anniversary(a, exact) if unit == "Y" else None
            when = ""
            if unit == "Y" and jump is not None:
                word = T("Geburtstag" if birth else "Jahrestag", "birthday" if birth else "anniversary")
                if now > exact:
                    when = T(f", der {exact + 1}. {word} ist erst am {_shown(jump)}",
                             f", the {_ordinal(exact + 1)} {word} is not until {_shown(jump)}")
                else:
                    when = T(f", der {exact}. {word} war schon am {_shown(jump)}",
                             f", the {_ordinal(exact)} {word} was already on {_shown(jump)}")
            example = T(f"{formel_number(now)} statt {exact} {unit_word}{when}",
                        f"{formel_number(now)} instead of {exact} {unit_word}{when}")
            if len(wrong) == 1:
                where = T(f"Zeile {cell[0]}: {example}", f"row {cell[0]}: {example}")
            else:
                where = T(f"Zeilen {_rows([w[0][0] for w in wrong])}; in Zeile {cell[0]}: {example}",
                          f"rows {_rows([w[0][0] for w in wrong])}; in row {cell[0]}: {example}")
            text = T(f"{sheet.figure_label(pos)} rechnet {what} von Hand und {why}. In {len(wrong)} von {rows} Zeilen stimmt "
                     f"das Ergebnis darum nicht ({where}).",
                     f"{sheet.figure_label(pos)} calculates {what} by hand and {why}. So the result is wrong in "
                     f"{len(wrong)} of {rows} rows ({where}).")
            level = "fehler"
        else:
            text = T(f"{sheet.figure_label(pos)} rechnet {what} von Hand und {why}. Im Moment stimmt es in allen {rows} "
                     "Zeilen, mit anderen Daten aber nicht.",
                     f"{sheet.figure_label(pos)} calculates {what} by hand and {why}. Right now it is correct in "
                     f"{'the one row' if rows == 1 else f'all {rows} rows'}, but not with other data.")
            level = "pruefen"
        if len(run) > 1:
            text += _all_cells(run)

        def build(cell, path=path, unit=unit):
            mine = sheet.nodes.get(cell)
            part = _part(mine, path)
            shape = _whole(sheet, part) if part is not None else None
            if shape is None:
                return None
            new = Node("call", "DATEDIF", (shape[1], shape[2], Node("str", unit)))
            return _text(_put(mine, path, new), _kept(sheet.cells[cell].formula or ""))

        fix = _fix(sheet, calc, run, build, expected)
        tip = _tip(fix, run, len(expected), calendar=True)
        found.append(Finding(f"{'age' if unit == 'Y' else 'months'}:{address(pos)}", text, level, tip, fix=fix))
    return found


def _all_cells(run: list[Pos]) -> str:
    return T(f" Das gilt für alle {len(run)} Zellen {address(run[0])}:{address(run[-1])}.",
             f" This applies to all {len(run)} cells {address(run[0])}:{address(run[-1])}.")


def formel_number(value: float) -> str:
    return number_text(round(value, 4))


def _part(node: Node | None, path: tuple) -> Node | None:
    try:
        return part_of(node, path) if node is not None else None
    except (IndexError, AttributeError):
        return None


def part_of(node: Node, path: tuple) -> Node:
    for i in path:
        node = node.args[i]
    return node


# --- HEUTE() in Fällen ---


def _todays(sheet: Sheet, calc: _Calc, done: set[Pos]) -> list[Finding]:  # noqa: C901 - Alter und Dauer
    found, seen = [], set()
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in done or pos in seen or not any(c.value in _TODAY for c in formel.calls(node)):
            continue
        if not sheet.plan.in_table(*pos):
            continue
        block = _block(sheet, pos)
        # Ist ein Fall der Spalte noch offen (kein Austritt), rechnet HEUTE() mit Absicht: dann die ganze Spalte nicht.
        ended = all("ende" in _cases(sheet, cell[0]) for cell in block)
        cases = _cases(sheet, pos[0])
        label = set(_name(sheet, pos)) | set(_words(sheet.figure_label(pos)))
        if label & _NOW:
            continue
        match = None
        for path, part in _paths(node):
            shape = _whole(sheet, part)
            if shape and _kind(sheet, shape[2]) == "heute" and (_kind(sheet, shape[1]) == "geburt" or label & _AGE) \
                    and "start" in cases:
                match = ("alter", path, shape)
                break
            if part.kind == "call" and part.value == "DATEDIF" and len(part.args) == 3 and \
                    _kind(sheet, part.args[1]) == "heute" and _kind(sheet, part.args[0]) == "geburt" and "start" in cases:
                match = ("alter", path, ("Y", part.args[0], part.args[1], "datedif", None))
                break
            span = _span(sheet, part)
            if span and _kind(sheet, span[1]) == "heute" and _kind(sheet, span[0]) == "start" and ended:
                match = ("dauer", path, span)
                break
        if match is None:
            continue
        what, path, shape = match
        run = _column(sheet, pos)
        seen |= set(block) | set(run)
        other = "start" if what == "alter" else "ende"
        expected, missing = {}, 0
        for cell in run:
            mine = _part(sheet.nodes.get(cell), path)
            row = _cases(sheet, cell[0])
            if mine is None or other not in row:
                missing += 1
                continue
            if what == "alter":
                inner = _whole(sheet, mine) if shape[3] != "datedif" else ("Y", mine.args[0], mine.args[1], "datedif", None)
                if inner is None:
                    missing += 1
                    continue
                a, b = calc.value(inner[1]), calc.value(Node("ref", row[other]))
                exact = _datedif(_day(a), _day(b), "Y") if _number(a) and _number(b) else None
            else:
                a, b = calc.value(_span(sheet, mine)[0]) if _span(sheet, mine) else None, calc.value(Node("ref", row[other]))
                exact = b - a if _number(a) and _number(b) else None
            if exact is None:
                missing += 1
                continue
            expected[cell] = float(exact)
        # Ist eine Zeile noch offen (kein Austritt), rechnet HEUTE() dort mit Absicht.
        if missing or not expected:
            continue
        header = sheet.plan.headers.get(cases[other][1], "")
        said = bool(label & (_START if what == "alter" else _END))
        if what == "alter":
            text = T(f"{sheet.figure_label(pos)} rechnet das Alter bis HEUTE(). Das ändert sich jeden Tag, und wer die Mappe "
                     "später öffnet, bekommt andere Zahlen. " +
                     (f"Im Namen steht aber „bei {header}“." if said else
                      f"In einer Tabelle von Fällen mit „{header}“ ist meist das Alter bei {header} gemeint."),
                     f"{sheet.figure_label(pos)} calculates the age up to TODAY(). That changes every day, and anyone who "
                     "opens the workbook later gets different numbers. " +
                     (f"But the name says “{header}”." if said else
                      f"In a table of cases with “{header}”, the age at {header} is usually meant."))
            if shape[3] != "datedif":
                text += T(" Außerdem rechnet es das Alter von Hand statt mit DATEDIF, so liegt es um den Geburtstag herum daneben.",
                          " It also calculates the age by hand instead of with DATEDIF, so it is off around the birthday.")
        else:
            text = T(f"{sheet.figure_label(pos)} rechnet die Dauer bis HEUTE(), obwohl in jeder Zeile „{header}“ steht. "
                     "Das ändert sich jeden Tag, und wer die Mappe später öffnet, bekommt andere Zahlen.",
                     f"{sheet.figure_label(pos)} calculates the duration up to TODAY(), although every row has “{header}”. "
                     "That changes every day, and anyone who opens the workbook later gets different numbers.")
        if len(run) > 1:
            text += _all_cells(run)

        def build(cell, path=path, what=what, other=other):
            mine = _part(sheet.nodes.get(cell), path)
            row = _cases(sheet, cell[0])
            if mine is None or other not in row:
                return None
            here = Node("ref", row[other])
            if what == "alter":
                inner = _whole(sheet, mine)
                first = inner[1] if inner else mine.args[0]
                new = Node("call", "DATEDIF", (first, here, Node("str", "Y")))
            else:
                new = Node("op", "-", (here, _span(sheet, mine)[0]))
            return _text(_put(sheet.nodes[cell], path, new), _kept(sheet.cells[cell].formula or ""))

        fix = _fix(sheet, calc, run, build, expected)
        question = "" if said else T(f"Ist das Alter bei {header} gemeint", f"Is the age at {header} meant") \
            if what == "alter" else T(f"Ist die Dauer bis {header} gemeint", f"Is the duration until {header} meant")
        tip = _tip(fix, run, len(expected), question=question)
        found.append(Finding(f"today:{address(pos)}", text, "fehler" if said else "pruefen", tip, fix=fix))
    return found


# --- Tage mit und ohne den ersten Tag ---


def _plus(node: Node) -> tuple[Node, int] | None:
    """Ende − Beginn, Ende − Beginn + 1, 1 + (Ende − Beginn), 1 + Ende − Beginn, Ende − Beginn − 1 → (Differenz,
    Zusatz)."""
    if node.kind == "op" and node.value in "+-" and len(node.args) == 2:
        a, b = node.args
        if b.kind == "num" and b.value == 1 and a.kind == "op" and a.value == "-":
            return a, (1 if node.value == "+" else -1)
        if node.value == "+" and a.kind == "num" and a.value == 1 and b.kind == "op" and b.value == "-":
            return b, 1
        # 1+D2-C2 liest Excel als (1+D2)-C2.
        if node.value == "-" and a.kind == "op" and a.value == "+" and len(a.args) == 2 and a.args[0] == Node("num", 1.0):
            return Node("op", "-", (a.args[1], b)), 1
    if node.kind == "op" and node.value == "-":
        return node, 0
    return None


def _days(sheet: Sheet, calc: _Calc, done: set[Pos]) -> list[Finding]:  # noqa: C901 - je Wort eine Regel
    found, seen = [], set()
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in done or pos in seen:
            continue
        # Zuerst die Form (Ende − Beginn mit zwei Daten), dann der Name: Die meisten Formeln sind keine Tage.
        shape = _plus(node)
        span = _span(sheet, shape[0]) if shape else None
        if span is None or _kind(sheet, span[1]) == "heute":
            continue
        label = set(_name(sheet, pos)) | set(_words(sheet.figure_label(pos)))
        work = bool(label & _WORKDAYS) or {"working", "days"} <= label or {"business", "days"} <= label
        inclusive, exclusive = bool(label & _INCLUSIVE), bool(label & (_EXCLUSIVE | _NIGHTS))
        if not (work or inclusive or exclusive) or (inclusive and exclusive):
            continue
        extra = shape[1]
        if work:
            want = "work"
        elif inclusive and extra == 0:
            want = 1
        elif exclusive and extra == 1:
            want = 0
        else:
            continue
        run = _column(sheet, pos)
        seen |= set(run)
        expected, sample = {}, None
        for cell in run:
            mine = sheet.nodes.get(cell)
            got = _plus(mine) if mine is not None else None
            pair = _span(sheet, got[0]) if got else None
            if pair is None:
                continue
            a, b = calc.value(pair[0]), calc.value(pair[1])
            if not (_number(a) and _number(b)):
                continue
            a, b = math.floor(a), math.floor(b)
            expected[cell] = float(formel.workdays(a, b) if want == "work" else b - a + want)
            if sample is None:
                sample = (cell, _day(a), _day(b), calc.value(mine))
        if not expected or sample is None:
            continue
        cell, a, b, now = sample
        if want == "work":
            why = T(f"zählt alle Kalendertage mit den Wochenenden: von {_shown(a)} bis {_shown(b)} sind es "
                    f"{formel_number(now) if _number(now) else '?'}, Arbeitstage (Montag bis Freitag, beide mitgezählt) "
                    f"aber {int(expected[cell])}",
                    f"counts all calendar days including weekends: from {_shown(a)} to {_shown(b)} that is "
                    f"{formel_number(now) if _number(now) else '?'}, but working days (Monday to Friday, both counted) "
                    f"{int(expected[cell])}")
        elif want == 1:
            why = T(f"heißt „inklusive“, zählt den ersten Tag aber nicht mit: von {_shown(a)} bis {_shown(b)} sind es mit "
                    f"beiden Tagen {int(expected[cell])}, nicht {formel_number(now)}",
                    f"is called “inclusive” but does not count the first day: from {_shown(a)} to {_shown(b)} it is "
                    f"{int(expected[cell])} with both days, not {formel_number(now)}")
        else:
            word = T("Nächte" if label & _NIGHTS else "Tage ohne den ersten Tag",
                     "nights" if label & _NIGHTS else "days without the first day")
            why = T(f"zählt mit +1 beide Tage, „{word}“ sind aber einer weniger: von {_shown(a)} bis {_shown(b)} sind es "
                    f"{int(expected[cell])}, nicht {formel_number(now)}",
                    f"counts both days with +1, but “{word}” are one fewer: from {_shown(a)} to {_shown(b)} it is "
                    f"{int(expected[cell])}, not {formel_number(now)}")
        text = f"{sheet.figure_label(pos)} {why}."
        if len(run) > 1:
            text += _all_cells(run)

        def build(cell, want=want):
            mine = sheet.nodes.get(cell)
            got = _plus(mine) if mine is not None else None
            pair = _span(sheet, got[0]) if got else None
            if pair is None:
                return None
            if want == "work":
                new = Node("call", "NETWORKDAYS", (pair[0], pair[1]))
            elif want == 1:
                new = Node("op", "+", (got[0], Node("num", 1.0)))
            else:
                new = got[0]
            return _text(new, _kept(sheet.cells[cell].formula or ""))

        fix = _fix(sheet, calc, run, build, expected)
        tip = _tip(fix, run, len(expected))
        if fix and want == "work":
            tip += T(" Feiertage zieht NETTOARBEITSTAGE auch ab, wenn du ihren Bereich als drittes Argument angibst.",
                     " NETWORKDAYS also subtracts holidays if you give their range as the third argument.")
        found.append(Finding(f"days:{address(pos)}", text, "fehler", tip, fix=fix))
    return found


def _taken(sheet: Sheet, found: list[Finding]) -> set[Pos]:
    """Die Zellen, die ein Hinweis schon abdeckt: seine heruntergezogene Spalte."""
    return {cell for f in found for cell in _column(sheet, formel.position(f.key.partition(":")[2]))}


def checks(sheet: Sheet) -> list[Finding]:
    calc = _Calc(sheet)
    found = _todays(sheet, calc, set())
    found += _wholes(sheet, calc, _taken(sheet, found))
    found += _days(sheet, calc, _taken(sheet, found))
    return found[:MAX_FINDINGS]
