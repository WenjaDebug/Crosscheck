"""Monatsende, Geschäftsjahr und Quartal, die nicht zu dem passen, was über der Spalte steht.

Ohne KI, nur aus Texten des Blatts und der Form der Formel:
- Unter „Monatsende“ steht =MONATSENDE(A16;-1): Das ist das Ende des Vormonats.
- Das Blatt sagt „Geschäftsjahr Oktober bis September“, die Formel =JAHR(A16)+(MONAT(A16)>10) zählt aber erst ab
  November zum nächsten Geschäftsjahr. Richtig ist >9 (oder >=10).
- Unter „Quartal GJ“ rechnet =AUFRUNDEN(MONAT(A16)/3;0) das Kalenderquartal, nicht das des Geschäftsjahrs.
"""

from __future__ import annotations

import re

from . import einsetzen
from .analyse import Finding, Sheet, address
from .sprache import T, english

Pos = tuple[int, int]

_MONTHS = {
    "januar": 1, "jänner": 1, "january": 1, "februar": 2, "february": 2, "märz": 3, "maerz": 3, "march": 3,
    "april": 4, "mai": 5, "may": 5, "juni": 6, "june": 6, "juli": 7, "july": 7, "august": 8, "september": 9,
    "oktober": 10, "october": 10, "november": 11, "dezember": 12, "december": 12,
}
_MONTH = "|".join(sorted(_MONTHS, key=len, reverse=True))
_MONTHS_EN = ("January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
              "November", "December")
_FISCAL = re.compile(rf"(?:geschäftsjahr|geschaeftsjahr|wirtschaftsjahr|fiscal year|\bgj\b|\bfy\b)[^.;\n]{{0,20}}?"
                     rf"(?:von\s+|from\s+|vom\s+)?({_MONTH})\s*(?:bis|to|–|-)\s*({_MONTH})", re.IGNORECASE)
_MONTH_END = re.compile(r"monatsende|monatsletzt|month.?end|end of month|ultimo", re.IGNORECASE)
_NOT_THIS = re.compile(r"vormonat|vorher|previous|prior|folgemonat|nächst|next|vorjahr", re.IGNORECASE)
_FY_HEAD = re.compile(r"geschäftsjahr|geschaeftsjahr|wirtschaftsjahr|\bgj\b|fiscal|\bfy\b", re.IGNORECASE)
_QUARTER = re.compile(r"quartal|quarter|\bq\b", re.IGNORECASE)

_EOMONTH = re.compile(r"EOMONTH\(([^(),]+),\s*(-?\d+)\)")
_FY_PLUS = re.compile(r"YEAR\(([^()]+)\)\s*\+\s*\(MONTH\(\1\)\s*(>=?)\s*(\d+)\)")
_CAL_QUARTER = re.compile(r"ROUNDUP\(MONTH\(([^()]+)\)/3,\s*0\)|INT\(\(MONTH\(([^()]+)\)\+2\)/3\)")
MAX_ROWS = 2000


def _fiscal_start(sheet: Sheet) -> int | None:
    starts = set()
    for cell in sheet.cells.values():
        if cell.formula or not isinstance(cell.value, str):
            continue
        for m in _FISCAL.finditer(cell.value):
            first, last = _MONTHS[m.group(1).casefold()], _MONTHS[m.group(2).casefold()]
            if (last - first) % 12 == 11:
                starts.add(first)
    return starts.pop() if len(starts) == 1 else None


def _head(sheet: Sheet, pos: Pos) -> str:
    """Der Text über der Spalte: die nächste Textzelle darüber (die Formeln dazwischen gehören zur selben Spalte)."""
    for r in range(pos[0] - 1, max(0, pos[0] - MAX_ROWS), -1):
        cell = sheet.cells.get((r, pos[1]))
        if cell is None:
            return ""
        if cell.formula:
            continue
        if isinstance(sheet.inputs.get((r, pos[1])), (int, float)):
            return ""
        return str(cell.value).strip()
    return ""


def _runs(cells: list[Pos]) -> list[list[Pos]]:
    runs: list[list[Pos]] = []
    for pos in sorted(cells, key=lambda p: (p[1], p[0])):
        if runs and runs[-1][-1][1] == pos[1] and runs[-1][-1][0] + 1 == pos[0]:
            runs[-1].append(pos)
        else:
            runs.append([pos])
    return runs


def _fmt_date(value) -> str:
    from datetime import datetime, timedelta

    from .analyse import _fmt

    if isinstance(value, (int, float)) and not isinstance(value, bool) and 1 < value < 2958466:
        return (datetime(1899, 12, 30) + timedelta(days=int(value))).strftime("%d.%m.%Y")
    return _fmt(value)


def _for_month(sheet: Sheet, pos: Pos) -> str:
    from . import formel

    m = _EOMONTH.search(sheet.cells[pos].formula or "")
    try:
        ref = formel.position(m.group(1).replace("$", "").strip()) if m else None
    except (ValueError, AttributeError, KeyError):
        ref = None
    if ref is None or not isinstance(sheet.inputs.get(ref), (int, float)):
        return ""
    return T(f" für den Monat {_fmt_date(sheet.inputs[ref])} ({address(ref)})",
             f" for the month {_fmt_date(sheet.inputs[ref])} ({address(ref)})")


def _fix(sheet: Sheet, changes: dict[Pos, str]) -> dict | None:
    """Die Korrektur; ein Fehlerwert nur dort, wo schon die alte Formel beim Nachrechnen einen gibt (ein Datum als
    Text, das Excel selbst umwandelt, der Prüfer aber nicht)."""
    from . import formel

    calc, _ = sheet.evaluator()
    inherited = True
    for pos, text in changes.items():
        try:
            new = calc.run(formel.parse(text))
            old = calc.run(sheet.nodes[pos]) if sheet.nodes.get(pos) is not None else None
        except Exception:  # noqa: BLE001 - nachrechnen geht nicht: keine Korrektur
            return None
        if isinstance(new, formel.Err) and not isinstance(old, formel.Err):
            inherited = False
    return einsetzen.formulas(sheet, changes, errors=inherited)


def checks(sheet: Sheet) -> list[Finding]:
    found: list[Finding] = []
    start = _fiscal_start(sheet)
    ends, years, quarters = [], [], []
    for pos, cell in sheet.cells.items():
        formula = cell.formula or ""
        if not formula:
            continue
        head = _head(sheet, pos)
        if (m := _EOMONTH.search(formula)) and m.group(2) != "0" and _MONTH_END.search(head) \
                and not _NOT_THIS.search(head):
            ends.append(pos)
        if start and _FY_HEAD.search(head) and not _QUARTER.search(head) and (m := _FY_PLUS.search(formula)):
            k = int(m.group(3))
            if (m.group(2) == ">" and k != start - 1) or (m.group(2) == ">=" and k != start):
                years.append(pos)
        if start and start != 1 and _QUARTER.search(head) and _FY_HEAD.search(head) and _CAL_QUARTER.search(formula):
            quarters.append(pos)
    for run in _runs(ends)[:2]:
        changes = {p: _EOMONTH.sub(r"EOMONTH(\1,0)", sheet.cells[p].formula.lstrip("=")) for p in run}
        fix = _fix(sheet, changes)
        first = run[0]
        n = _EOMONTH.search(sheet.cells[first].formula).group(2)
        where = address(first) if len(run) == 1 else f"{address(first)}:{address(run[-1])}"
        k = abs(int(n))
        found.append(Finding(
            f"monthend:{address(first)}",
            T(f"{where} steht unter „{_head(sheet, first)}“, rechnet aber MONATSENDE(…;{n}): das Monatsende "
              f"{k} {'Monat' if k == 1 else 'Monate'} {'früher' if int(n) < 0 else 'später'}. "
              f"So zeigt {address(first)} {_fmt_date(sheet.inputs.get(first))}"
              f"{_for_month(sheet, first)}.",
              f"{where} is under “{_head(sheet, first)}”, but calculates EOMONTH(…, {n}): the end of the month "
              f"{k} month{'' if k == 1 else 's'} {'earlier' if int(n) < 0 else 'later'}. "
              f"So {address(first)} shows {_fmt_date(sheet.inputs.get(first))}{_for_month(sheet, first)}."),
            "fehler", T("Mit 0 als zweitem Argument ist es das Ende desselben Monats.",
                        "With 0 as the second argument, it is the end of the same month."), fix=fix))
    names = {v: k.capitalize() for k, v in _MONTHS.items() if k in (
        "januar", "februar", "märz", "april", "mai", "juni", "juli", "august", "september", "oktober", "november",
        "dezember")} if not english() else dict(enumerate(_MONTHS_EN, start=1))
    for run in _runs(years)[:2]:
        first = run[0]
        m = _FY_PLUS.search(sheet.cells[first].formula)
        changes = {p: _FY_PLUS.sub(lambda x: f"YEAR({x.group(1)})+(MONTH({x.group(1)})>{start - 1})",
                                   sheet.cells[p].formula.lstrip("=")) for p in run}
        fix = _fix(sheet, changes)
        k = int(m.group(3)) - (0 if m.group(2) == ">" else 1)  # ab Monat k+1 zählt die Formel zum neuen Jahr
        if k >= start:
            months, side = list(range(start, k + 1)), T("noch zum alten", "still")
        else:
            months, side = list(range(k + 1, start)), T("schon zum neuen", "already")
        one = len(months) == 1
        span = names[months[0]] if one else T(f"{names[months[0]]} bis {names[months[-1]]}",
                                              f"{names[months[0]]} to {names[months[-1]]}")
        where = address(first) if len(run) == 1 else f"{address(first)}:{address(run[-1])}"
        text = T(f"Das Blatt sagt Geschäftsjahr {names[start]} bis {names[(start + 10) % 12 + 1]}. {where} rechnet aber "
                 f"=JAHR(…)+(MONAT(…){m.group(2)}{m.group(3)}): So {'zählt' if len(months) == 1 else 'zählen'} {span} "
                 f"{side} Geschäftsjahr",
                 f"The sheet says fiscal year {names[start]} to {names[(start + 10) % 12 + 1]}. But {where} calculates "
                 f"=JAHR(…)+(MONAT(…){m.group(2)}{m.group(3)}): so {span} {side} {'counts' if one else 'count'} "
                 f"toward the {'old' if k >= start else 'new'} fiscal year")
        found.append(Finding(f"fiscal:{address(first)}", text + ".", "fehler",
                             T(f"Das neue Geschäftsjahr fängt im {names[start]} an: MONAT(…)>{start - 1}.",
                               f"The new fiscal year starts in {names[start]}: MONTH(…)>{start - 1}."), fix=fix))
    for run in _runs(quarters)[:2]:
        first = run[0]

        def fiscal(x: re.Match) -> str:
            ref = x.group(1) or x.group(2)
            return f"ROUNDUP((MOD(MONTH({ref})-{start},12)+1)/3,0)"

        changes = {p: _CAL_QUARTER.sub(fiscal, sheet.cells[p].formula.lstrip("=")) for p in run}
        fix = _fix(sheet, changes)
        where = address(first) if len(run) == 1 else f"{address(first)}:{address(run[-1])}"
        found.append(Finding(
            f"fquarter:{address(first)}",
            T(f"{where} steht unter „{_head(sheet, first)}“, rechnet aber das Kalenderquartal (MONAT/3). Das "
              f"Geschäftsjahr fängt laut Blatt im {names[start]} an, dann ist {names[start]} bis "
              f"{names[(start + 1) % 12 + 1]} das erste Quartal.",
              f"{where} is under “{_head(sheet, first)}”, but calculates the calendar quarter (MONTH/3). According to "
              f"the sheet, the fiscal year starts in {names[start]}, so {names[start]} to "
              f"{names[(start + 1) % 12 + 1]} is the first quarter."),
            "fehler", T(f"Quartal des Geschäftsjahrs: AUFRUNDEN((REST(MONAT(…)-{start};12)+1)/3;0).",
                        f"Fiscal-year quarter: ROUNDUP((MOD(MONTH(…)-{start}, 12)+1)/3, 0)."), fix=fix))
    return found
