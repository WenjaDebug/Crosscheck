"""Ausreißer innerhalb einer Gruppe: Laborwerte im Langformat (je Zeile Patient, Parameter, Wert) mischen in einer
Spalte Kalium um 4, Natrium um 140 und Thrombozyten um 250. Kalium 45 statt 4,5 fällt dort gegen die ganze Spalte
nicht auf, gegen die übrigen Kaliumwerte aber sehr. Ebenso Preise je Artikel, Dosen je Medikament, Löhne je Stufe.

Eine Gruppe zählt nur, wo sie die Spalte wirklich ordnet: In den Gruppen streuen die Werte höchstens ein Viertel so
weit wie in der ganzen Spalte, und Gruppen mit mindestens acht Werten tragen fast alle Zeilen. Dann sagt ein Ausreißer
gegen die ganze Spalte nichts (der größte Natriumwert ist nur Natrium), und es zählt der Vergleich in der Gruppe. Passt
ein Wert mal 10, 100 oder 1000 genau zu seiner Gruppe, steht die richtige Zahl zum Einsetzen dabei."""

from __future__ import annotations

import math
import re
import statistics
from collections import Counter, defaultdict

from .analyse import Finding, Sheet, _fmt, _outliers, address
from .formel import position
from .sprache import T

Pos = tuple[int, int]
GROUP_MIN = 8
# So weit dürfen die Werte in den Gruppen höchstens streuen, gemessen an der ganzen Spalte.
SHARE = 0.25
# So viele Zeilen der Spalte müssen in Gruppen mit mindestens GROUP_MIN Werten stehen.
COVERED = 0.8
MAX_GROUPS = 200
_UNIT_HEAD = re.compile(r"einheit|\bunit|\bdimension|\bmasseinheit", re.I)
# Vorsilben, mit denen sich eine Einheit nur um eine Zehnerpotenz ändert (g/l und g/dl, mmol/l und µmol/l).
_PREFIX = {"G": 9, "M": 6, "k": 3, "h": 2, "d": -1, "c": -2, "m": -3, "µ": -6, "μ": -6, "u": -6, "n": -9, "p": -12}
_BASES = ("mol", "kat", "IU", "Pa", "g", "l", "L", "m", "U", "s")


def checks(sheet: Sheet, columns: dict, rows: list[int]) -> tuple[list[Finding], set[str]]:
    """Die Hinweise und die Schlüssel der Ausreißer gegen die ganze Spalte, die in einer geordneten Spalte nichts
    sagen (analyse.py)."""
    from .auswertung import is_naming
    from .kategorien import _categories, _norm

    keys, _ = _categories(sheet, columns, rows)
    if not keys:
        return [], set()
    found: list[Finding] = []
    covered: set[str] = set()
    for col, info in columns.items():
        header = sheet.plan.headers.get(col, "")
        if col in keys or info.kind != "zahl" or is_naming(header):
            continue
        numbers = {r: v for r, v in info.numbers.items()
                   if math.isfinite(v) and (r, col) in sheet.cells and sheet.cells[(r, col)].formula is None}
        if len(numbers) < 2 * GROUP_MIN:
            continue
        best = _grouping(numbers, keys, col, _norm)
        if best is None:
            continue
        key_col, groups = best
        covered |= {f"{kind}:{address((r, col))}" for r in numbers for kind in ("outlier", "scale")}
        names = keys[key_col]
        group_header = sheet.plan.headers.get(key_col, "")
        units = next((c for c in keys if c not in (key_col, col) and _UNIT_HEAD.search(sheet.plan.headers.get(c, ""))), None)
        for members in groups.values():
            if len(members) < GROUP_MIN:
                continue
            shown = names[next(iter(members))]
            where = T(f"in {header}, wo {group_header} „{shown.strip()}“ ist", f"in {header} where {group_header} is “{shown.strip()}”")
            for item in _outliers(sheet, col, header, members, where):
                if units is not None:
                    item = _other_unit(sheet, item, col, units, keys[units], members, where, covered)
                found.append(item)
    return sorted(found, key=lambda f: position(f.key.partition(":")[2])), covered


def _other_unit(sheet: Sheet, item: Finding, col: int, units: int, texts: dict[int, str], members: dict[int, float],
                where: str, covered: set[str]) -> Finding:
    """Hämoglobin 142 mit „g/l“ in der Zeile, die übrigen in „g/dl“: Der Wert stimmt wohl, nur in einer anderen
    Einheit. Unterscheiden sich die Einheiten nur in der Vorsilbe, ist die Umrechnung genau (142 g/l sind 14,2 g/dl),
    und Einsetzen schreibt Wert und Einheit wie in den übrigen Zeilen."""
    from . import einsetzen

    pos = position(item.key.partition(":")[2])
    r = pos[0]
    given = {m: texts[m].strip() for m in members if m in texts}
    if r not in given or len(given) < GROUP_MIN:
        return item
    usual, count = Counter(given.values()).most_common(1)[0]
    if given[r] == usual or count < 0.8 * len(given):
        return item
    value, unit_cell = members[r], (r, units)
    others = sorted(v for m, v in members.items() if m != r and given.get(m) == usual)
    low, high = others[len(others) // 10], others[-1 - len(others) // 10]
    lead = T(f"{sheet.label(pos)} = {_fmt(value)} liegt weit weg von den übrigen Werten {where} (sonst {_fmt(low)} bis "
             f"{_fmt(high)}). In {address(unit_cell)} steht dazu „{given[r]}“, sonst „{usual}“.",
             f"{sheet.label(pos)} = {_fmt(value)} is far away from the other values {where} (otherwise {_fmt(low)} to "
             f"{_fmt(high)}). {address(unit_cell)} gives the unit as “{given[r]}”, the others “{usual}”.")
    # Was kategorien.py zur Einheit allein sagt (gehört sonst zu „mmol/l“), ginge mit dem Wert in die Irre.
    covered |= {f"{kind}:{address(unit_cell)}" for kind in ("mapping", "typo")}
    covered |= {f"pair:{address((r, c))}" for c in range(1, max(col, units) + 1)}
    factor = _factor(given[r], usual)
    if factor is None:
        return Finding(f"rowunit:{address(pos)}", T(f"{lead} Für Mittelwerte, Pivot-Tabellen und Vergleiche gehören alle "
                                                    f"Werte in dieselbe Einheit.",
                                                    f"{lead} For averages, pivot tables and comparisons, all values need to "
                                                    f"be in the same unit."), "pruefen",
                       T(f"Rechne {address(pos)} genau in {usual} um, oder korrigiere die Einheit, wenn sie falsch dasteht.",
                         f"Convert {address(pos)} exactly to {usual}, or correct the unit if it is wrong."))
    right = float(f"{value * factor:.12g}")
    if not low <= right <= high:
        return Finding(f"rowunit:{address(pos)}", T(f"{lead} Umgerechnet wären es {_fmt(right)} {usual}, das passt auch "
                                                    "nicht zu den übrigen.",
                                                    f"{lead} Converted, it would be {_fmt(right)} {usual}, which does not fit "
                                                    "the others either."), "pruefen",
                       T("Prüfe Wert und Einheit an der Quelle.", "Check the value and the unit at the source."))
    like = next(((m, col) for m in members if m != r), None)
    fix = einsetzen.values(sheet, {pos: right}, like, T(f"{_fmt(right)} und „{usual}“", f"{_fmt(right)} and “{usual}”"))
    if fix is not None:
        cell = sheet.cells.get(unit_cell)
        if cell is None or cell.formula is not None:
            fix = None
        else:
            fix["cells"][address(unit_cell)] = usual
            fix["was"][address(unit_cell)] = {"text": cell.text if cell.text is not None else cell.value}
            fix["range"] = f"{address(pos)}:{address(unit_cell)}" if units == col + 1 else fix["range"]
    return Finding(
        f"rowunit:{address(pos)}",
        T(f"{lead} {_fmt(value)} {given[r]} sind {_fmt(right)} {usual}. Für Mittelwerte, Pivot-Tabellen und Vergleiche "
          "gehören alle Werte in dieselbe Einheit.",
          f"{lead} {_fmt(value)} {given[r]} is {_fmt(right)} {usual}. For averages, pivot tables and comparisons, "
          "all values need to be in the same unit."),
        "pruefen", T(f"Einsetzen schreibt {_fmt(right)} in {address(pos)} und „{usual}“ in {address(unit_cell)}.",
                     f"Applying the fix writes {_fmt(right)} to {address(pos)} and “{usual}” to {address(unit_cell)}."), fix=fix)


def _factor(source: str, target: str) -> float | None:
    """Womit ein Wert in source malzunehmen ist, damit er in target dasteht, wenn sich die Einheiten nur in den
    Vorsilben unterscheiden (g/l → g/dl: 0,1; mmol/l → µmol/l: 1000), sonst None."""
    a, b = _unit(source), _unit(target)
    if a is None or b is None or a[0] != b[0]:
        return None
    return 10.0 ** (a[1] - b[1])


def _unit(text: str) -> tuple[tuple, int] | None:
    """„mmol/l“ → ((("mol",), ("l",)), -3): die Grundeinheiten über und unter dem Bruchstrich und die Zehnerpotenz."""
    text = text.replace(" ", "")
    if not text or text.count("/") > 1:
        return None
    top, _, bottom = text.partition("/")
    power, bases = 0, []
    for sign, part in ((1, top), (-1, bottom)):
        if not part:
            bases.append(())
            continue
        parsed = _part(part)
        if parsed is None:
            return None
        bases.append(parsed[0])
        power += sign * parsed[1]
    return tuple(bases), power


def _part(part: str) -> tuple[tuple, int] | None:
    for base in _BASES:
        if part.endswith(base):
            rest = part[: -len(base)]
            if rest == "":
                return (base.lower() if base in "lL" else base,), 0
            if rest in _PREFIX:
                return (base.lower() if base in "lL" else base,), _PREFIX[rest]
    # Nur eine Vorsilbe: eine Anzahl (G/l, Milliarden je Liter, ist dasselbe wie /nl).
    if part in _PREFIX and part not in ("m", "d", "c", "u"):
        return (), _PREFIX[part]
    return None


def _grouping(numbers: dict[int, float], keys: dict[int, dict[int, str]], col: int, norm):
    """Die Spalte, nach der die Zahlen am deutlichsten geordnet sind, und die Zahlen je Gruppe, oder None."""
    values = list(numbers.values())
    scale = (lambda v: math.log10(v)) if min(values) > 0 else (lambda v: v)
    scaled = {r: scale(v) for r, v in numbers.items()}
    center = statistics.median(scaled.values())
    spread = statistics.median(abs(v - center) for v in scaled.values())
    if spread == 0:
        return None
    best = None
    for key_col, texts in keys.items():
        if key_col == col:
            continue
        groups: dict[str, dict[int, float]] = defaultdict(dict)
        for r, v in numbers.items():
            if r in texts:
                groups[norm(texts[r])][r] = v
        big = {k: g for k, g in groups.items() if len(g) >= GROUP_MIN}
        if len(big) < 2 or len(groups) > MAX_GROUPS or sum(map(len, big.values())) < COVERED * len(numbers):
            continue
        within = []
        for g in big.values():
            middle = statistics.median(scaled[r] for r in g)
            within += [abs(scaled[r] - middle) for r in g]
        ratio = statistics.median(within) / spread
        if ratio <= SHARE and (best is None or ratio < best[0]):
            best = (ratio, key_col, dict(groups))
    return None if best is None else best[1:]
