"""Statistische Tests und Intervalle, die falsch aufgesetzt sind, und Prozent, die eigentlich Prozentpunkte sind.

- T.TEST gepaart (Typ 1), obwohl die Werte in verschiedenen Zeilen stehen, also nicht zu denselben Fällen gehören,
  und ungepaart, obwohl jede Zeile derselbe Fall zweimal gemessen ist (vorher und nachher, 2024 und 2025).
- T.TEST mit gleichen Varianzen (Typ 2), wo die Streuungen weit auseinanderliegen: Dann gilt Welch (Typ 3).
- T.TEST einseitig (Seiten 1), ohne dass es dasteht: Das halbiert den p-Wert.
- KONFIDENZ.NORM und KONFIDENZ.T mit 0,95 statt 0,05, mit dem Standardfehler oder der Varianz statt der
  Standardabweichung, mit einer Anzahl, die nicht zu den Werten passt, und mit der Normalverteilung bei wenigen Werten.
- Ein Konfidenzintervall von Hand: Mittelwert ± 1,96 × STABW ohne Wurzel aus n (das ist der Bereich der einzelnen
  Werte), durch n statt durch die Wurzel aus n, 1,96 bei wenigen Werten (dort gilt t) und ein z, das nicht zum
  Niveau im Namen passt (1,645 bei „95 %“).
- Die Differenz zweier Prozentwerte heißt „Veränderung in %“ (es sind Prozentpunkte), oder „Prozentpunkte“
  rechnen die relative Veränderung.

Jede Korrektur ist auf dem Blatt nachgerechnet und ergibt genau, was der richtige Test oder das richtige Intervall
ergibt. Alles ohne KI.
"""

from __future__ import annotations

import math
import re
import statistics

from . import einsetzen, formel
from .analyse import Finding, Sheet, _fmt, address
from .bereiche import _kept, _paths, _put, _text, _words
from .formel import Node
from .sprache import T
from .statistik import _name
from .workbook import FUNCTIONS, pattern
from .xlsx import number_text

Pos = tuple[int, int]

MAX_FINDINGS = 6
_TTEST = {"T.TEST", "TTEST"}
_CONFIDENCE = {"CONFIDENCE", "CONFIDENCE.NORM", "CONFIDENCE.T"}
_SD = {"STDEV", "STDEV.S", "STDEV.P", "STDEVP"}
_VAR = {"VAR", "VAR.S", "VAR.P", "VARP"}
_KINDS = {1: "gepaart (Typ 1)", 2: "ungepaart mit gleichen Varianzen (Typ 2)", 3: "ungepaart nach Welch (Typ 3)"}
_KINDS_EN = {1: "paired", 2: "unpaired with equal variances", 3: "unpaired using Welch's method"}
_BEFORE = {"vorher", "vor", "prä", "prae", "pre", "baseline", "basis", "beginn", "start", "anfang", "eintritt",
           "aufnahme", "erstmessung", "ausgangswert", "initial"}
_AFTER = {"nachher", "nach", "post", "danach", "ende", "schluss", "follow", "followup", "austritt", "entlassung",
          "zweitmessung", "final", "verlauf"}
_SIDES = ({"links", "rechts"}, {"left", "right"})
# „Messung 1“ und „Messung 2“ sind derselbe Fall zweimal, „Gruppe 1“ und „Gruppe 2“ zwei verschiedene.
_TIMES = {"messung", "messungen", "zeitpunkt", "visite", "termin", "tag", "woche", "monat", "quartal", "jahr", "runde",
          "durchgang", "wiederholung", "zyklus", "sitzung", "untersuchung", "erhebung", "welle", "befragung", "t",
          "time", "visit", "week", "day", "month", "year", "wave", "session"}
_GROUPS = {"gruppe", "gruppen", "arm", "kohorte", "station", "klinik", "spital", "standort", "team", "filiale", "region",
           "abteilung", "zentrum", "center", "praxis", "land", "kanton", "firma", "produkt", "marke", "variante",
           "methode", "verfahren", "behandlung", "therapie", "medikament", "dosis", "labor", "gerät", "schule",
           "klasse", "group", "site", "cohort"}
_ONE_SIDED = {"einseitig", "einseitige", "einseitiger", "einseitiges", "gerichtet", "gerichtete", "gerichteter",
              "onesided", "onetailed"}
_TWO_SIDED = {"zweiseitig", "zweiseitige", "zweiseitiger", "zweiseitiges", "ungerichtet", "ungerichtete",
              "ungerichteter", "twosided", "twotailed"}
_CI = {"konfidenzintervall", "konfidenzintervalls", "konfidenz", "ki", "ci", "vertrauensintervall", "vertrauensbereich",
       "konfidenzgrenze", "konfidenzbereich", "confidence", "fehlermarge"}
# Dort ist Mittelwert ± 1,96 × STABW richtig: der Bereich, in dem die einzelnen Werte liegen.
_SPREAD_RANGE = {"referenzbereich", "normbereich", "streubereich", "normalbereich", "vorhersageintervall", "prediction",
                 "toleranzbereich", "einzelwerte", "referenz", "prädiktionsintervall", "prognoseintervall",
                 "toleranzintervall", "normwert", "normwerte", "eingriffsgrenze", "warngrenze", "regelgrenze"}
# z für ein zweiseitiges Intervall, wie man es in Excel tippt, und das Niveau dazu.
_Z = {1.64: 90, 1.645: 90, 1.65: 90, 1.96: 95, 2.0: 95, 2.57: 99, 2.576: 99, 2.58: 99}
_Z_OF = {90: 1.645, 95: 1.96, 99: 2.576}
_POINTS = {"prozentpunkte", "prozentpunkt", "prozentpunkten", "pp", "ppt", "punkte", "punkten"}
# „Differenz“ und „Unterschied“ sagen selbst, dass es eine Differenz ist: Dort sind Prozentpunkte gemeint.
_CHANGE = {"veränderung", "änderung", "wachstum", "zuwachs", "zunahme", "steigerung", "anstieg", "abnahme", "rückgang",
           "change", "growth", "increase"}
_RELATIVE = {"relativ", "relative", "relativer", "prozentual", "prozentuale", "prozentualer"}


def _german(name: str) -> str:
    return FUNCTIONS.get(name, name)


def _number(value) -> bool:
    return isinstance(value, float) and not isinstance(value, bool) and math.isfinite(value)


def _close(a: float, b: float) -> bool:
    return abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))


def _p(value: float) -> str:
    """Ein p-Wert so genau, dass man ihn mit 0,05 vergleichen kann: 0,031, 0,00012."""
    if value >= 0.001:
        return number_text(round(value, 3))
    if value < 1e-12:
        return T("fast 0", "almost 0")
    return number_text(float(f"{value:.2g}"))


def _num(value: float) -> Node:
    return Node("num", float(round(value, 12)))


class _Calc:
    def __init__(self, sheet: Sheet):
        self.sheet = sheet
        self.calc, _ = sheet.evaluator()

    def run(self, node: Node):
        try:
            return self.calc.run(node)
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError, IndexError):
            return None

    def area(self, node: Node) -> list | None:
        """Die Werte eines Bereichs auf diesem Blatt, Zeile für Zeile (None, wenn es keiner ist)."""
        if node.kind != "range" or node.value[2] >= formel.MAX_ROW:
            return None
        try:
            return self.calc.values_of(node)
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError, IndexError):
            return None


def _numbers(values: list) -> list[float]:
    return [v for v in values if _number(v)]


def _fixed(sheet: Sheet, pos: Pos, node: Node, changes: list[tuple[tuple, Node]]) -> str | None:
    """Die Formel der Zelle mit neuen Teilen (Pfad → Knoten), englisch und mit den $ wie vorher."""
    for path, new in changes:
        node = _put(node, path, new)
    return _text(node, _kept(sheet.cells[pos].formula or ""))


def _verified(sheet: Sheet, calc: _Calc, pos: Pos, node: Node, changes, path, expected: float) -> dict | None:
    """Die Korrektur zum Einsetzen, wenn der geänderte Teil (bei path) genau expected ergibt."""
    text = _fixed(sheet, pos, node, changes)
    if text is None:
        return None
    try:
        new = formel.parse(text)
    except formel.Unknown:
        return None
    part = new
    for i in path:
        part = part.args[i]
    got = calc.run(part)
    if not _number(got) or not _close(got, expected):
        return None
    return einsetzen.formulas(sheet, {pos: text})


def _joined(why: list[str]) -> str:
    """Die Gründe als Sätze: „bekommt … . Außerdem rechnet es …“ (jeder Grund fängt mit seinem Verb an)."""
    out = why[0] + "."
    for more in why[1:]:
        verb, _, rest = more.partition(" ")
        out += T(f" Außerdem {verb} es {rest}.", f" It also {verb} {rest}.")
    return out


def _finding(kind: str, pos: Pos, text: str, level: str, fix: dict | None, said: str) -> Finding:
    tip = T(f"Korrigiert: {fix['german']}, nachgerechnet {said}.",
            f"Corrected: {fix['german']}, recalculated {said}.") if fix else ""
    return Finding(f"{kind}:{address(pos)}", text, level, tip, fix=fix)


# --- T.TEST ---


def _header(sheet: Sheet, area: tuple) -> str:
    r1, c1, _, _ = area
    if sheet.plan.in_table(r1, c1):
        return sheet.plan.headers.get(c1, "")
    above = sheet.cells.get((r1 - 1, c1))
    return above.value if above is not None and above.formula is None and isinstance(above.value, str) else ""


def _repeated(one: str, two: str) -> bool:
    """Zwei Spaltenköpfe für dieselbe Größe zweimal gemessen: „Blutdruck vorher“ und „Blutdruck nachher“,
    „Messung 1“ und „Messung 2“, „Umsatz 2024“ und „Umsatz 2025“, „Auge links“ und „Auge rechts“."""
    a, b = _words(one), _words(two)
    if not a or not b or a == b:
        return False
    left, right = set(a) - set(b), set(b) - set(a)
    if left & _BEFORE and right & _AFTER or right & _BEFORE and left & _AFTER:
        return True
    if any(left | right <= pair for pair in _SIDES) and left and right:
        return True
    same = {w for w in a if w in b}
    if not same or len(left) != 1 or len(right) != 1 or same & _GROUPS:
        return False
    if all(re.fullmatch(r"t\d+|(19|20)\d\d", w) for w in left | right):
        return True
    return bool(same & _TIMES) and all(re.fullmatch(r"\d+", w) for w in left | right)


def _sides(words: set[str], one: set[str], number: str, word: str) -> bool:
    """Steht „einseitig“ (oder „zweiseitig“) im Namen, auch als „1-seitig“ oder „one-sided“?"""
    return bool(words & one) or {number, "seitig"} <= words or {word, "sided"} <= words or {word, "tailed"} <= words


def _ttests(sheet: Sheet, calc: _Calc) -> list[Finding]:  # noqa: C901 - je Fehler eine Regel
    found = []
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown or not any(c.value in _TTEST for c in formel.calls(node)):
            continue
        for path, call in _paths(node):
            if call.kind != "call" or call.value not in _TTEST or len(call.args) != 4:
                continue
            one, two = call.args[0], call.args[1]
            first, second = calc.area(one), calc.area(two)
            tails, kind, now = calc.run(call.args[2]), calc.run(call.args[3]), calc.run(call)
            if first is None or second is None or not _number(tails) or not _number(kind):
                continue
            tails, kind = math.trunc(tails), math.trunc(kind)
            if tails not in (1, 2) or kind not in (1, 2, 3):
                continue
            # Gepaart über verschieden viele Werte zeigt Excel #NV: Auch das ist zu erklären.
            uneven = kind == 1 and len(first) != len(second) and isinstance(now, formel.Err) and now.code == "#N/A"
            if not _number(now) and not uneven:
                continue
            x, y = _numbers(first), _numbers(second)
            if len(x) < 2 or len(y) < 2:
                continue
            a, b = one.value, two.value
            same_rows = (a[0], a[2]) == (b[0], b[2]) and (a[1], a[3]) != (b[1], b[3])
            complete = len(x) == len(first) and len(y) == len(second)
            words = set(_name(sheet, pos)) | set(_words(sheet.figure_label(pos)))
            want_kind, want_tails, why, level = kind, tails, [], "pruefen"
            shown = f"{address(a[:2])}:{address(a[2:])}{T(' und ', ' and ')}{address(b[:2])}:{address(b[2:])}"
            if uneven:
                level = "fehler"
                why.append(T(f"rechnet gepaart (Typ 1), aber {shown} sind verschieden viele Werte ({len(x)} und "
                             f"{len(y)}), darum zeigt Excel #NV. Gepaart heißt: Zeile für Zeile derselbe Fall zweimal "
                             "gemessen. Zwei Gruppen vergleicht der ungepaarte Test, sicher ist Welch (Typ 3)",
                             f"calculates a paired test (type 1), but {shown} have different numbers of values "
                             f"({len(x)} and {len(y)}), so Excel shows #N/A. Paired means: row by row, the same case "
                             "measured twice. Two groups are compared by the unpaired test; the safe choice is Welch "
                             "(type 3)"))
                want_kind = 3
            elif kind == 1 and not same_rows:
                if a[1] == b[1] and a[3] == b[3]:
                    level = "fehler"
                    why.append(T(f"rechnet gepaart (Typ 1), aber {shown} stehen in derselben Spalte untereinander, also "
                                 "sind es verschiedene Fälle. Gepaart heißt: jede Zeile ist derselbe Fall zweimal gemessen. "
                                 "Zwei Gruppen vergleicht der ungepaarte Test, sicher ist Welch (Typ 3)",
                                 f"calculates a paired test (type 1), but {shown} are one below the other in the same "
                                 "column, so they are different cases. Paired means: each row is the same case measured "
                                 "twice. Two groups are compared by the unpaired test; the safe choice is Welch (type 3)"))
                else:
                    why.append(T(f"rechnet gepaart (Typ 1), aber {shown} stehen in verschiedenen Zeilen. Gepaart heißt: "
                                 "Zeile für Zeile derselbe Fall zweimal gemessen. Gehören die Werte nicht so zusammen, ist "
                                 "es der ungepaarte Test, sicher ist Welch (Typ 3)",
                                 f"calculates a paired test (type 1), but {shown} are in different rows. Paired means: "
                                 "row by row, the same case measured twice. If the values do not belong together like "
                                 "that, it is the unpaired test; the safe choice is Welch (type 3)"))
                want_kind = 3
            elif kind in (2, 3) and same_rows and complete and sheet.plan.in_table(a[0], a[1]) and \
                    _repeated(_header(sheet, a), _header(sheet, b)):
                why.append(T(f"vergleicht „{_header(sheet, a)}“ und „{_header(sheet, b)}“ {_KINDS[kind].split(' (')[0]}, "
                             "nach den Spaltenköpfen ist aber jede Zeile derselbe Fall zweimal gemessen. Dann gehört der "
                             "gepaarte Test (Typ 1) dazu: Er vergleicht die Veränderung je Fall",
                             f"compares “{_header(sheet, a)}” and “{_header(sheet, b)}” {_KINDS_EN[kind]}, but according "
                             "to the column headers each row is the same case measured twice. Then the paired test "
                             "(type 1) applies: it compares the change per case"))
                want_kind = 1
            elif kind == 2:
                v1, v2 = statistics.variance(x), statistics.variance(y)
                ratio = max(v1, v2) / min(v1, v2) if min(v1, v2) > 0 else math.inf
                if ratio >= 2 and (len(x) != len(y) or ratio >= 4):
                    welch = _test(x, y, first, second, tails, 3)
                    if welch is not None and (abs(welch - now) >= 0.1 * max(welch, now) or (welch < 0.05) != (now < 0.05)):
                        why.append(T(f"nimmt gleiche Varianzen an (Typ 2), die eine Varianz ist aber {number_text(round(ratio, 1))}-mal "
                                     f"so groß wie die andere ({_fmt(v1)} und {_fmt(v2)}, n = {len(x)} und {len(y)}). Dann "
                                     "gilt Welch (Typ 3), der keine gleichen Varianzen braucht",
                                     f"assumes equal variances (type 2), but one variance is {number_text(round(ratio, 1))} "
                                     f"times as large as the other ({_fmt(v1)} and {_fmt(v2)}; n = {len(x)} and {len(y)}). "
                                     "Then Welch (type 3) applies, which does not need equal variances"))
                        want_kind = 3
            if tails == 1 and not _sides(words, _ONE_SIDED, "1", "one"):
                if _sides(words, _TWO_SIDED, "2", "two"):
                    level = "fehler"
                    why.append(T("rechnet einseitig (Seiten 1), obwohl „zweiseitig“ dabeisteht",
                                 "calculates one-tailed (tails 1), although the name says “two-tailed”"))
                else:
                    why.append(T("rechnet einseitig (Seiten 1): Das halbiert den p-Wert. Einseitig ist nur richtig, wenn vor "
                                 "der Auswertung feststand, in welche Richtung der Unterschied gehen muss; üblich ist zweiseitig",
                                 "calculates one-tailed (tails 1): that halves the p-value. One-tailed is only right if, "
                                 "before the analysis, it was clear in which direction the difference must go; two-tailed "
                                 "is usual"))
                want_tails = 2
            if not why:
                continue
            expected = _test(x, y, first, second, want_tails, want_kind)
            if expected is None:
                continue
            changes = [(path + (2,), _num(want_tails)), (path + (3,), _num(want_kind))]
            fix = _verified(sheet, calc, pos, node, changes, path, expected)
            text = T(f"T.TEST in {sheet.figure_label(pos)} {_joined(why)} Richtig aufgesetzt ergibt es "
                     f"p = {_p(expected)} statt {'#NV' if uneven else _p(now)}.",
                     f"T.TEST in {sheet.figure_label(pos)} {_joined(why)} Set up correctly, it gives "
                     f"p = {_p(expected)} instead of {'#N/A' if uneven else _p(now)}.")
            if uneven:
                pass
            elif expected < 0.05 <= now:
                text += T(" Richtig gerechnet ist der Unterschied bei 5 % signifikant, die jetzige Formel verdeckt das.",
                          " Calculated correctly, the difference is significant at 5 %; the current formula hides that.")
            elif now < 0.05 <= expected:
                text += T(" Richtig gerechnet ist der Unterschied bei 5 % nicht signifikant, die jetzige Formel zeigt ihn "
                          "fälschlich als signifikant.",
                          " Calculated correctly, the difference is not significant at 5 %; the current formula wrongly "
                          "shows it as significant.")
            found.append(_finding("ttest", pos, text, level, fix, f"p = {_p(expected)}"))
            break
    return found


def _test(x: list[float], y: list[float], first: list, second: list, tails: int, kind: int) -> float | None:
    try:
        t, df = formel.t_statistic(first, second, kind) if kind == 1 else formel.t_statistic(x, y, kind)
    except (formel.Unknown, formel._Fail, ZeroDivisionError, ValueError, statistics.StatisticsError):  # noqa: SLF001
        return None
    return tails * formel.t_tail(abs(t), df)


# --- Konfidenzintervalle ---


def _spread(sheet: Sheet, calc: _Calc, node: Node, depth: int = 0) -> tuple[str, Node, Node, list[float], Node] | None:
    """Was eine Streuung ist: ("sd" | "var" | "sem", Aufruf der Streuung, Bereich, Zahlen darin, die Streuung so, wie
    sie in der Formel steht). Auch über eine Zelle, in der sie steht (=STABW.S(B2:B21) in F3), und für den
    Standardfehler STABW/WURZEL(n): Dort ist das Letzte die Standardabweichung über dem Bruchstrich (F3)."""
    while node.kind == "call" and node.value in ("ROUND", "ROUNDUP", "ROUNDDOWN") and node.args:
        node = node.args[0]
    if node.kind == "ref" and depth < 2 and node.value in sheet.nodes and sheet.nodes[node.value] is not None:
        inner = _spread(sheet, calc, sheet.nodes[node.value], depth + 1)
        return None if inner is None else (*inner[:4], inner[4] if inner[0] == "sem" else node)
    if node.kind == "call" and node.value in _SD | _VAR and len(node.args) == 1:
        values = calc.area(node.args[0])
        if values is None:
            return None
        return ("sd" if node.value in _SD else "var"), node, node.args[0], _numbers(values), node
    if node.kind == "op" and node.value == "/" and len(node.args) == 2:
        inner = _spread(sheet, calc, node.args[0], depth)
        bottom = node.args[1]
        if inner is not None and inner[0] == "sd" and bottom.kind == "call" and bottom.value == "SQRT":
            n = calc.run(bottom.args[0])
            if _number(n) and _close(n, len(inner[3])):
                return "sem", inner[1], inner[2], inner[3], node.args[0]
    return None


def _sample(call: Node) -> bool:
    return call.value in ("STDEV", "STDEV.S", "VAR", "VAR.S")


def _level(sheet: Sheet, pos: Pos) -> int | None:
    """Das Niveau im Namen: „95-%-KI“, „KI 99 %“."""
    r, c = pos
    texts = [sheet.figure_label(pos)]
    for near in ((r, c - 1), (r - 1, c)):
        cell = sheet.cells.get(near)
        if cell is not None and cell.formula is None and isinstance(cell.value, str):
            texts.append(cell.value)
    for text in texts:
        match = re.search(r"(?<![\d,.])(90|95|99)\s*-?\s*%", text)
        if match:
            return int(match.group(1))
    return None


def _confidences(sheet: Sheet, calc: _Calc, done: set[Pos]) -> list[Finding]:  # noqa: C901 - je Argument eine Regel
    found = []
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown or pos in done or not any(c.value in _CONFIDENCE for c in formel.calls(node)):
            continue
        for path, call in _paths(node):
            if call.kind != "call" or call.value not in _CONFIDENCE or len(call.args) != 3:
                continue
            alpha, sd, size, now = (calc.run(a) for a in (*call.args, call))
            if not all(_number(v) for v in (alpha, sd, size, now)):
                continue
            spread = _spread(sheet, calc, call.args[1])
            changes, why, level = [], [], "pruefen"
            new_alpha, name = alpha, call.value
            if 0.8 <= alpha < 1:
                new_alpha = round(1 - alpha, 12)
                new = _num(new_alpha) if call.args[0].kind == "num" else Node("op", "-", (Node("num", 1.0), call.args[0]))
                changes.append((path + (0,), new))
                level = "fehler"
                why.append(T(f"bekommt als Alpha {_fmt(alpha)}. Alpha ist die Irrtumswahrscheinlichkeit: Für ein "
                             f"{number_text(round(alpha * 100, 4))}-%-Intervall gehört {_fmt(new_alpha)} hinein",
                             f"gets {_fmt(alpha)} as alpha. Alpha is the error probability: for a "
                             f"{number_text(round(alpha * 100, 4))} % interval, it should be {_fmt(new_alpha)}"))
            n = math.trunc(size)
            values = spread[3] if spread else None
            if spread is not None and spread[0] in ("sem", "var"):
                level = "fehler"
                if spread[0] == "sem":
                    changes.append((path + (1,), spread[4]))
                    why.append(T("bekommt den Standardfehler statt der Standardabweichung. KONFIDENZ teilt selbst durch die "
                                 "Wurzel aus n, so teilt es zweimal, und das Intervall wird zu schmal",
                                 "gets the standard error instead of the standard deviation. CONFIDENCE divides by the "
                                 "square root of n itself, so it divides twice, and the interval becomes too narrow"))
                else:
                    changes.append((path + (1,), Node("call", "SQRT", (call.args[1],))))
                    why.append(T("bekommt die Varianz statt der Standardabweichung (ihrer Wurzel)",
                                 "gets the variance instead of the standard deviation (its square root)"))
            if values is not None and len(values) >= 2 and n != len(values):
                changes.append((path + (2,), Node("call", "COUNT", (spread[2],))))
                level = "fehler"
                span = f"{address(spread[2].value[:2])}:{address(spread[2].value[2:])}"
                why.append(T(f"rechnet mit n = {n}, in {span} stehen "
                             f"aber {len(values)} Werte",
                             f"calculates with n = {n}, but {span} contains {len(values)} values"))
                n = len(values)
            if name != "CONFIDENCE.T" and spread is not None and _sample(spread[1]) and 2 <= n < 30:
                name = "CONFIDENCE.T"
                why.append(T(f"rechnet mit der Normalverteilung (z), bei nur {n} Werten und einer geschätzten "
                             "Standardabweichung gilt aber die t-Verteilung (KONFIDENZ.T)",
                             f"uses the normal distribution (z), but with only {n} values and an estimated standard "
                             "deviation, the t-distribution applies (CONFIDENCE.T)"))
            if not why:
                continue
            if spread is not None and spread[0] in ("sem", "var", "sd") and len(values or []) >= 2:
                deviation = (statistics.stdev if _sample(spread[1]) else statistics.pstdev)(values)
            elif spread is None and not changes[1:] and _number(sd):
                deviation = sd
            else:
                continue
            if not (0 < new_alpha < 1) or n < 1 or (name == "CONFIDENCE.T" and n < 2):
                continue
            factor = (formel.t_inverse(new_alpha, n - 1) if name == "CONFIDENCE.T"
                      else statistics.NormalDist().inv_cdf(1 - new_alpha / 2))
            expected = factor * deviation / math.sqrt(n)
            if name != call.value:
                base = Node("call", name, call.args)
                for where, new in changes:
                    base = _put(base, where[len(path):], new)
                changes = [(path, base)]
            fix = _verified(sheet, calc, pos, node, changes, path, expected)
            text = T(f"{_german(call.value)} in {sheet.figure_label(pos)} {_joined(why)} Richtig ist die halbe "
                     f"Breite ±{_fmt(expected)} statt ±{_fmt(now)}.",
                     f"{call.value} in {sheet.figure_label(pos)} {_joined(why)} The correct half-width is "
                     f"±{_fmt(expected)} instead of ±{_fmt(now)}.")
            found.append(_finding("confidence", pos, text, level, fix, f"±{_fmt(expected)}"))
            break
    return found


def _factors(node: Node) -> tuple[list[tuple[tuple, Node]], list[tuple[tuple, Node]]]:
    """Ein Produkt als Faktoren oben und unten, je mit Pfad: 1,96*F3/WURZEL(F4) → ([1,96, F3], [WURZEL(F4)])."""
    top, bottom = [], []

    def walk(item, path, up):
        if item.kind == "op" and item.value == "*" and len(item.args) == 2:
            walk(item.args[0], path + (0,), up)
            walk(item.args[1], path + (1,), up)
        elif item.kind == "op" and item.value == "/" and len(item.args) == 2:
            walk(item.args[0], path + (0,), up)
            walk(item.args[1], path + (1,), not up)
        else:
            (top if up else bottom).append((path, item))

    walk(node, (), True)
    return top, bottom


def _margin(sheet: Sheet, calc: _Calc, node: Node):
    """z × Streuung [/ Teiler]: (Pfad und Wert von z, die Streuung, Teiler-Pfad und -Knoten oder None)."""
    top, bottom = _factors(node)
    z = [(p, n) for p, n in top if n.kind == "num" and round(n.value, 3) in _Z]
    spreads = [(p, n, s) for p, n in top if (s := _spread(sheet, calc, n)) is not None and s[0] == "sd"]
    if len(z) != 1 or len(spreads) != 1 or len(top) != 2 or len(bottom) > 1:
        return None
    return z[0], spreads[0], (bottom[0] if bottom else None)


def _has_z(node: Node) -> bool:
    """Steht irgendwo eine Zahl wie 1,96 (nicht als Hochzahl, ^2 im BMI ist kein z)?"""
    if node.kind == "num":
        return round(node.value, 3) in _Z
    args = node.args[:1] if node.kind == "op" and node.value == "^" else node.args
    return any(_has_z(a) for a in args)


def _intervals(sheet: Sheet, calc: _Calc, done: set[Pos]) -> list[Finding]:  # noqa: C901 - je Fehler eine Regel
    """Ein Konfidenzintervall von Hand: Mittelwert ± z × Standardabweichung / Wurzel aus n."""
    found = []
    for pos, node in sorted(sheet.nodes.items()):
        # Ohne ein z wie 1,96 in der Formel ist es kein Intervall von Hand (schnell, vor dem Namen).
        if node is None or pos in sheet.unknown or pos in done or not _has_z(node):
            continue
        words = set(_name(sheet, pos)) | set(_words(sheet.figure_label(pos)))
        if not words & _CI or words & _SPREAD_RANGE:
            continue
        parts = [((), node)]
        if node.kind == "op" and node.value in "+-" and len(node.args) == 2:
            parts = [((1,), node.args[1]), ((0,), node.args[0])]
        for where, part in parts:
            shape = _margin(sheet, calc, part)
            if shape is None:
                continue
            (zpath, znode), (spath, snode, spread), divisor = shape
            values = spread[3]
            n = len(values)
            if n < 2:
                continue
            now = calc.run(part)
            if not _number(now):
                continue
            meant = _level(sheet, pos)
            z = znode.value
            want_z, changes, why, level = z, [], [], "pruefen"
            count = Node("call", "COUNT", (spread[2],))
            if divisor is None:
                changes.append((where + spath, Node("op", "/", (snode, Node("call", "SQRT", (count,))))))
                level = "fehler"
                why.append(T(f"rechnet {_fmt(z)} × Standardabweichung ohne die Wurzel aus n. Das ist der Bereich, in dem "
                             f"etwa {_Z[round(z, 3)]} % der einzelnen Werte liegen, nicht das Konfidenzintervall des "
                             f"Mittelwerts. Dafür gehört die Standardabweichung durch die Wurzel aus {n} (Anzahl Werte)",
                             f"calculates {_fmt(z)} × standard deviation without the square root of n. That is the "
                             f"range in which about {_Z[round(z, 3)]} % of the individual values lie, not the confidence "
                             f"interval of the mean. For that, the standard deviation must be divided by the square root "
                             f"of {n} (number of values)"))
            else:
                dpath, dnode = divisor
                d = calc.run(dnode)
                if not _number(d):
                    continue
                if _close(d, n) and not (dnode.kind == "call" and dnode.value == "SQRT"):
                    changes.append((where + dpath, Node("call", "SQRT", (dnode,))))
                    level = "fehler"
                    why.append(T(f"teilt die Standardabweichung durch n ({n}) statt durch die Wurzel aus n",
                                 f"divides the standard deviation by n ({n}) instead of by the square root of n"))
                elif not _close(d, math.sqrt(n)):
                    continue
            wrong = meant is not None and _Z[round(z, 3)] != meant
            if wrong:
                want_z = _Z_OF[meant]
                level = "fehler"
            alpha = round(1 - (meant or _Z[round(z, 3)]) / 100, 12)
            factor = want_z
            small = n < 30 and _sample(spread[1])
            if small:
                factor = formel.t_inverse(alpha, n - 1)
                changes.append((where + zpath, Node("call", "T.INV.2T", (_num(alpha), Node("op", "-", (count, Node("num", 1.0)))))))
            elif wrong:
                changes.append((where + zpath, _num(want_z)))
            # Englisch mit Komma zwischen den Argumenten: Alpha (0,1, 0,05 oder 0,01) darum gleich mit Punkt.
            t = T(f"t = {_fmt(factor)} (T.INV.2S({_fmt(alpha)};{n - 1}))",
                  f"t = {_fmt(factor)} (T.INV.2T({alpha:g},{n - 1}))")
            if wrong:
                why.append(T(f"nimmt z = {_fmt(z)}, das gehört zu {_Z[round(z, 3)]} %, im Namen steht aber {meant} %"
                             + (f"; bei nur {n} Werten gilt dafür {t}" if small else f" (z = {_fmt(want_z)})"),
                             f"uses z = {_fmt(z)}; that belongs to {_Z[round(z, 3)]} %, but the name says {meant} %"
                             + (f"; with only {n} values, {t} applies" if small else f" (z = {_fmt(want_z)})")))
            elif small:
                why.append(T(f"nimmt z = {_fmt(z)} aus der Normalverteilung, bei nur {n} Werten gilt aber {t}",
                             f"uses z = {_fmt(z)} from the normal distribution, but with only {n} values, {t} applies"))
            if not why:
                continue
            deviation = (statistics.stdev if _sample(spread[1]) else statistics.pstdev)(values)
            expected = factor * deviation / math.sqrt(n)
            fix = _verified(sheet, calc, pos, node, changes, where, expected)
            text = T(f"{sheet.figure_label(pos)} {_joined(why)} Richtig ist die halbe Breite ±{_fmt(expected)} "
                     f"statt ±{_fmt(now)}.",
                     f"{sheet.figure_label(pos)} {_joined(why)} The correct half-width is ±{_fmt(expected)} "
                     f"instead of ±{_fmt(now)}.")
            found.append(_finding("interval", pos, text, level, fix, f"±{_fmt(expected)}"))
            break
    return found


# --- Prozent und Prozentpunkte ---


def _pct(share: float, digits: int = 2) -> str:
    return f"{number_text(round(share * 100, digits))} %"


def _pp(points: float) -> str:
    return number_text(round(points * 100, 2))


def _percent(sheet: Sheet, node: Node) -> bool:
    cell = sheet.cells.get(node.value) if node.kind == "ref" else None
    return cell is not None and isinstance(cell.value, str) and cell.value.strip().endswith("%")


def _difference(node: Node) -> tuple[Node, Node] | None:
    """B5-B4 → (neu, alt)."""
    if node.kind == "op" and node.value == "-" and len(node.args) == 2 and all(a.kind == "ref" for a in node.args):
        return node.args[0], node.args[1]
    return None


def _relative(node: Node) -> tuple[Node, Node] | None:
    """(B5-B4)/B4 oder B5/B4-1 → (neu, alt)."""
    if node.kind == "op" and node.value == "/" and len(node.args) == 2:
        pair = _difference(node.args[0])
        if pair is not None and node.args[1] == pair[1]:
            return pair
    if node.kind == "op" and node.value == "-" and len(node.args) == 2 and node.args[1] == Node("num", 1.0):
        top = node.args[0]
        if top.kind == "op" and top.value == "/" and all(a.kind == "ref" for a in top.args):
            return top.args[0], top.args[1]
    return None


def _column(sheet: Sheet, pos: Pos) -> list[Pos]:
    """pos und die Zellen darunter mit derselben Formel, heruntergezogen."""
    shape = pattern(sheet.cells[pos].formula or "", pos)
    run = [pos]
    while (below := (run[-1][0] + 1, pos[1])) in sheet.nodes and sheet.cells[below].formula and \
            pattern(sheet.cells[below].formula, below) == shape and len(run) < einsetzen.MAX_CELLS:
        run.append(below)
    return run


def _points(sheet: Sheet, calc: _Calc, done: set[Pos]) -> list[Finding]:
    found, seen = [], set()
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown or pos in done or pos in seen:
            continue
        diff, rel = _difference(node), _relative(node)
        pair = diff or rel
        if pair is None or not all(_percent(sheet, p) for p in pair):
            continue
        words = set(_name(sheet, pos)) | set(_words(sheet.figure_label(pos)))
        new, old = (calc.run(p) for p in pair)
        if not _number(new) or not _number(old):
            continue
        run = _column(sheet, pos)
        seen |= set(run)
        if rel is not None and words & _POINTS and not words & _RELATIVE:
            level, kind = "fehler", "points"
            what = T(f"heißt Prozentpunkte, rechnet aber die relative Veränderung: von {_pct(old)} auf {_pct(new)} sind "
                     f"{_pp(new - old)} Prozentpunkte" + (f", relativ {_pct((new - old) / old, 1)}" if old else ""),
                     f"is called percentage points, but calculates the relative change: from {_pct(old)} to "
                     f"{_pct(new)} is {_pp(new - old)} percentage points"
                     + (f", {_pct((new - old) / old, 1)} in relative terms" if old else ""))
            make = lambda n, o: Node("op", "-", (n, o))  # noqa: E731
        elif diff is not None and words & _CHANGE and ("%" in words or "prozent" in words or words & _RELATIVE) \
                and not words & _POINTS and old != 0:
            level, kind = ("fehler" if words & _RELATIVE else "pruefen"), "points"
            what = T(f"zieht zwei Prozentwerte voneinander ab: von {_pct(old)} auf {_pct(new)} sind {_pp(new - old)} "
                     f"Prozentpunkte. Mit „%“ im Namen liest man das als relative Veränderung, die ist aber "
                     f"{_pct((new - old) / old, 1)}",
                     f"subtracts two percentages: from {_pct(old)} to {_pct(new)} is {_pp(new - old)} percentage "
                     f"points. With “%” in the name, this reads as a relative change, but that is "
                     f"{_pct((new - old) / old, 1)}")
            make = lambda n, o: Node("op", "/", (Node("op", "-", (n, o)), o))  # noqa: E731
        else:
            continue
        changes = {}
        for cell in run:
            mine = sheet.nodes.get(cell)
            mpair = (_difference(mine) or _relative(mine)) if mine is not None else None
            text = _text(make(*mpair), _kept(sheet.cells[cell].formula or "")) if mpair else None
            if text is None:
                changes = {}
                break
            changes[cell] = text
        expected = new - old if rel is not None else (new - old) / old
        fix = None
        got = calc.run(formel.parse(changes[pos])) if changes else None
        if _number(got) and _close(got, expected):
            fix = einsetzen.formulas(sheet, changes)
        span = f"{address(run[0])}:{address(run[-1])}"
        where = T(f" Das gilt für alle {len(run)} Zellen {span}.", f" This applies to all {len(run)} cells {span}.") \
            if len(run) > 1 else ""
        text = f"{sheet.figure_label(pos)} {what}.{where}"
        if level == "pruefen":
            text += T(" Meinst du Prozentpunkte, schreib es so in den Namen.",
                      " If you mean percentage points, say so in the name.")
        said = _pct(expected, 1) if rel is None else T(f"{_pp(expected)} Prozentpunkte", f"{_pp(expected)} percentage points")
        item = _finding(kind, pos, text, level, fix, said)
        if fix is not None and len(run) > 1:
            item.tip = T(f"Korrigiert in {address(pos)}: {fix['german']}, dann bis {address(run[-1])} herunterziehen.",
                         f"Corrected in {address(pos)}: {fix['german']}, then fill it down to {address(run[-1])}.")
        if rel is not None and fix is not None:
            # Das Einsetzen lässt das Zahlenformat, wie es ist: 0,05 im Prozentformat zeigt Excel als 5 %.
            item.tip += T(f" Im Prozentformat zeigt Excel das als {_pp(expected)} %, gemeint sind {_pp(expected)} Prozentpunkte.",
                          f" In percent format, Excel shows this as {_pp(expected)} %, but {_pp(expected)} percentage "
                          "points are meant.")
        found.append(item)
    return found


def checks(sheet: Sheet) -> list[Finding]:
    calc = _Calc(sheet)
    found = _ttests(sheet, calc)
    done = {formel.position(f.key.partition(":")[2]) for f in found}
    found += _confidences(sheet, calc, done)
    done |= {formel.position(f.key.partition(":")[2]) for f in found}
    found += _intervals(sheet, calc, done)
    done |= {formel.position(f.key.partition(":")[2]) for f in found}
    found += _points(sheet, calc, done)
    return found[:MAX_FINDINGS]
