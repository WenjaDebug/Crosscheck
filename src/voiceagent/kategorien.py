"""Textspalten, deren Werte zusammenpassen müssen: Namen, Abteilungen, Codes und ihre Bezeichnungen.

Ein Name, der einmal anders geschrieben ist („Dr. Shmidt“ neben 16-mal „Dr. Schmidt“, ein Leerzeichen am Ende), ist
für eine Pivot-Tabelle, ZÄHLENWENN und SUMMEWENN ein eigener Eintrag; „kardiologie“ neben „Kardiologie“ zählt Excel
zusammen, R, Python und SPSS aber nicht. Und wo eine Spalte die andere bestimmt (ein ICD-Code seine Diagnose, ein Arzt
seine Abteilung), fällt die Zeile auf, in der das nicht gilt: E11.9 mit „Essentielle Hypertonie“. Welche Zelle der
Zeile falsch ist, sagen die übrigen Spalten: Passt nur eine Korrektur zu allen, ist es diese. Können sie es nicht sagen
(ein Code und seine Bezeichnung widersprechen sich, sonst nichts), stehen beide Korrekturen zur Wahl.

Dasselbe gilt für Zahlenspalten mit wenigen, wiederholten Werten: Zu „Kalium“ gehört der Referenzwert 5,1, zur
Lohnklasse 4 derselbe Grundlohn, zur Kostenstelle 410 die Buchhaltung. Messwerte (Gewicht, Betrag) wiederholen sich
nicht so und sind darum keine Kategorie.

Eine Regel „X bestimmt Y“ gilt nur, wo sie sich nicht zufällig ergibt: Jeder häufige Wert von X steht fast immer mit
demselben Y, und ohne Zusammenhang wären mindestens zehnmal so viele Zeilen anders. Ein oder zwei Ausreißer in einer
kurzen Liste sind Tippfehler, kein loser Zusammenhang: Dort reicht sechsmal so viele. Zwei Werte, die zufällig meist
zusammen vorkommen (Geschlecht und Raucher in zehn Zeilen), sind keine Regel.
"""

from __future__ import annotations

import math
import re
from collections import Counter, defaultdict

from . import formel
from .analyse import Finding, Sheet, _fmt, address, more_formulas, said
from .formel import Err
from .sprache import T

Pos = tuple[int, int]

MIN_ROWS = 10
MAX_LISTED = 3
MAX_FORMULAS = 20000
# Höchstens so viele Spalten werden paarweise verglichen (je Paar ein Durchgang über die Zeilen).
MAX_COLUMNS = 30


def checks(sheet: Sheet, columns: dict, rows: list[int]) -> list[Finding]:
    """Tippfehler, anders geschriebene Namen und Zeilen, die einer Regel zwischen zwei Spalten widersprechen."""
    raw, numbers = _categories(sheet, columns, rows)
    if not raw:
        return []
    keys = {col: {r: _norm(t) for r, t in texts.items()} for col, texts in raw.items()}
    counts = {col: Counter(k.values()) for col, k in keys.items()}
    spelling = {col: _dominant(raw[col], keys[col]) for col in raw}
    rules = _rules(keys)
    data = _Data(sheet, raw, keys, counts, spelling, rules, numbers)
    found: list[Finding] = []
    taken: set[Pos] = set()
    for r, broken in sorted(_broken(keys, rules).items()):
        item = _row(data, r, broken)
        if item is not None:
            found.append(item[0])
            taken |= item[1]
    for col in sorted(c for c in raw if c not in numbers):
        for item, cells in _typos(data, col, taken) + _variants(data, col, taken):
            found.append(item)
            taken |= cells
    return found


class _Data:
    def __init__(self, sheet, raw, keys, counts, spelling, rules, numbers=None) -> None:
        self.sheet, self.raw, self.keys, self.counts, self.spelling, self.rules = sheet, raw, keys, counts, spelling, rules
        # Zahlenspalten: je Zeile die Zahl selbst, zum Einsetzen und Nachrechnen.
        self.numbers: dict[int, dict[int, float]] = numbers or {}

    def header(self, col: int) -> str:
        return self.sheet.plan.headers.get(col, "")

    def shown(self, col: int, key: str) -> str:
        """Ein Wert so, wie er meistens geschrieben ist."""
        return self.spelling[col].get(key, key)

    def q(self, col: int, text: str) -> str:
        """„Kalium“ in Anführungszeichen, eine Zahl wie 5,1 ohne."""
        return text if col in self.numbers else T(f"„{text}“", f"“{text}”")

    def value(self, col: int, key: str):
        """Was eingesetzt wird: der Text, wie er meistens geschrieben ist, oder die Zahl, die meistens dasteht."""
        if col not in self.numbers:
            return self.shown(col, key)
        found = Counter(v for r, v in self.numbers[col].items() if self.keys[col].get(r) == key)
        return found.most_common(1)[0][0]

    def like(self, col: int, key: str) -> Pos | None:
        """Eine Zelle der Spalte mit diesem Wert: Ihr Zahlenformat bekommt die Korrektur."""
        return next(((r, col) for r, k in sorted(self.keys[col].items()) if k == key), None)


def _norm(text: str) -> str:
    return " ".join(text.split()).casefold()


def _categories(sheet: Sheet, columns: dict, rows: list[int]) -> tuple[dict[int, dict[int, str]], dict[int, dict[int, float]]]:
    """Textspalten mit wenigen, oft wiederholten Werten (Arzt, Station, Diagnose, Code), nicht freier Text. Dazu
    Zahlenspalten dieser Art (Kostenstelle, Referenzwert), als Text so, wie der Prüfer Zahlen schreibt, und die Zahlen."""
    # Auch gruppen.py fragt danach, mit denselben Spalten.
    known = sheet.memo.get("kategorien.categories")
    if known is not None and known[0] is columns and known[1] == rows:
        return known[2]
    out, numbers = {}, {}
    wanted = set(rows)
    for col, info in columns.items():
        if info.formulas:
            continue
        values = None
        if info.kind == "text":
            texts = {}
            for r in rows:
                if r not in info.texts:
                    continue
                cell = sheet.cells.get((r, col))
                text = cell.text if cell is not None and cell.text is not None else info.texts[r]
                if text.strip():
                    texts[r] = text
        elif info.kind == "zahl":
            values = {r: v for r, v in info.numbers.items() if r in wanted and math.isfinite(v)}
            texts = {r: _fmt(v) for r, v in values.items()}
        else:
            continue
        if len(texts) < MIN_ROWS or any(len(t) > 80 or "\n" in t for t in texts.values()):
            continue
        counts = Counter(_norm(t) for t in texts.values())
        repeated = sum(n for n in counts.values() if n >= 2)
        if len(counts) < 2 or len(counts) > max(2, len(texts) // 3) or repeated < 0.7 * len(texts):
            continue
        out[col] = texts
        if values is not None:
            numbers[col] = values
        if len(out) >= MAX_COLUMNS:
            break
    sheet.memo["kategorien.categories"] = (columns, rows, (out, numbers))
    return out, numbers


def _dominant(texts: dict[int, str], keys: dict[int, str]) -> dict[str, str]:
    """Je Wert die Schreibweise, die am häufigsten vorkommt."""
    ways: dict[str, Counter] = defaultdict(Counter)
    for r, t in texts.items():
        ways[keys[r]][t] += 1
    return {k: c.most_common(1)[0][0] for k, c in ways.items()}


def _rules(keys: dict[int, dict[int, str]]) -> dict[tuple[int, int], dict[str, tuple[str, int]]]:
    """(X, Y) → {Wert von X: (der Wert von Y, der dazu gehört, wie oft)} für jede Regel „X bestimmt Y“."""
    out = {}
    for x in sorted(keys):
        for y in sorted(keys):
            if x != y:
                rule = _rule(keys[x], keys[y])
                if rule:
                    out[(x, y)] = rule
    return out


def _rule(kx: dict[int, str], ky: dict[int, str]) -> dict[str, tuple[str, int]] | None:
    both = [r for r in kx if r in ky]
    n = len(both)
    if n < MIN_ROWS:
        return None
    by: dict[str, Counter] = defaultdict(Counter)
    for r in both:
        by[kx[r]][ky[r]] += 1
    marginal = Counter(ky[r] for r in both)
    mapping, bad, expected, covered = {}, 0, 0.0, 0
    for x, ys in by.items():
        total = sum(ys.values())
        if total < 3:
            continue
        ranked = ys.most_common(2)
        top, k = ranked[0]
        if len(ranked) > 1 and ranked[1][1] == k:
            return None
        off = total - k
        # Höchstens jede siebte Zeile anders; zwei Ausreißer neben mindestens sechs passenden auch.
        if k < 3 or (off > max(1.0, 0.15 * total) and not (off <= 2 and k >= 3 * off)):
            return None
        mapping[x] = (top, k)
        covered += total
        bad += total - k
        expected += total * (1 - marginal[top] / n)
    if len(mapping) < 2 or covered < 0.8 * n or len({y for y, _ in mapping.values()}) < 2:
        return None
    # Ohne Zusammenhang wären mindestens zehnmal so viele Zeilen anders als die Regel sagt, bei ein oder zwei
    # Ausreißern sechsmal so viele.
    if expected < 8 or (bad * 10 > expected and not (bad <= 2 and bad * 6 <= expected)):
        return None
    return mapping


def _broken(keys, rules) -> dict[int, list[tuple[int, int]]]:
    out: dict[int, list[tuple[int, int]]] = defaultdict(list)
    for (x, y), mapping in rules.items():
        for r, kx in keys[x].items():
            if kx in mapping and r in keys[y] and keys[y][r] != mapping[kx][0]:
                out[r].append((x, y))
    return out


def _still_broken(data: _Data, r: int, trial: dict[int, str]) -> bool:
    for (x, y), mapping in data.rules.items():
        if x in trial and y in trial and trial[x] in mapping and trial[y] != mapping[trial[x]][0]:
            return True
    return False


def _row(data: _Data, r: int, broken: list[tuple[int, int]]) -> tuple[Finding, set[Pos]] | None:
    """Eine Zeile, die einer oder mehreren Regeln widerspricht: die Zelle, deren Korrektur alle Regeln erfüllt."""
    here = {col: k[r] for col, k in data.keys.items() if r in k}
    involved = sorted({c for pair in broken for c in pair})
    repairs = []
    for c in involved:
        implied = {x: mapping[here[x]] for (x, y), mapping in data.rules.items()
                   if y == c and x in here and here[x] in mapping}
        wanted = {v for v, _ in implied.values()}
        if len(wanted) != 1:
            continue
        value = wanted.pop()
        if value == here[c] or _still_broken(data, r, {**here, c: value}):
            continue
        similar = c not in data.numbers and _typo(data.raw[c][r], data.shown(c, value))
        repairs.append((len(implied) + similar, c, value, implied, similar))
    if not repairs:
        return None
    repairs.sort(key=lambda t: (-t[0], t[1]))
    sheet = data.sheet
    if len(repairs) == 1 or repairs[0][0] > repairs[1][0]:
        support, c, value, implied, similar = repairs[0]
        pos = (r, c)
        now, right, put = data.raw[c][r], data.shown(c, value), data.value(c, value)
        if c not in data.numbers and _risky(right):
            return None
        reasons = _reasons(data, r, c, implied, value)
        others = data.counts[c].get(here[c], 0)
        cell, head = address(pos), data.header(c)
        if similar:
            lead = T(f"{cell} ({head}) „{now.strip()}“ ist wohl ein Tippfehler für „{right}“: {reasons}.",
                     f"{cell} ({head}) contains “{now.strip()}”, probably a typo for “{right}”: {reasons}.")
        else:
            # „kardiologie“ statt „Kardiologie“ und dazu die falsche Abteilung: beides sagen.
            spelled = data.shown(c, here[c])
            if now == spelled or _norm(now) != _norm(spelled):
                how = ""
            elif now == now.lower():
                how = T(" ist klein geschrieben und", " is in lower case and")
            else:
                how = T(" ist anders geschrieben und", " is spelled differently and")
            also = T(" auch sonst", " also") if how else ""
            lead = T(f"{cell} ({head}) {data.q(c, now.strip())}{how} passt{also} nicht zum Rest der "
                     f"Zeile: {reasons}.",
                     # „contains“ vorn: So liest das Fenster (ansicht._where) die Zelle davor wie auf Deutsch.
                     f"{cell} ({head}) {data.q(c, now.strip())}{how}{also} does not fit the rest of the row: {reasons}."
                     if c in data.numbers else
                     f"{cell} ({head}) contains {data.q(c, now.strip())}, which{how}{also} does not fit the rest of the "
                     f"row: {reasons}.")
        if c in data.numbers:
            effect = ""  # Eine Zahl zählt keine Pivot-Tabelle für sich; was sie ändert, sagen die Formeln.
        elif others > 1 and here[c] in data.spelling[c]:
            effect = T(f" Eine Pivot-Tabelle, ZÄHLENWENN oder SUMMEWENN nach {head} zählt die Zeile darum bei "
                       f"„{data.shown(c, here[c])}“ statt bei „{right}“.",
                       f" A pivot table, COUNTIF or SUMIF by {head} therefore counts the row under "
                       f"“{data.shown(c, here[c])}” instead of “{right}”.")
        else:
            effect = T(f" Eine Pivot-Tabelle, ZÄHLENWENN oder SUMMEWENN nach {head} zählt die Zeile darum "
                       "für sich.",
                       f" A pivot table, COUNTIF or SUMIF by {head} therefore counts the row separately.")
        effect += _effects(sheet, {pos: put})
        text = lead + effect + _should(data.q(c, right), data.q(c, now.strip()))
        level = "fehler" if support >= 2 else "pruefen"
        fix = _fix(sheet, {pos: put}, data.like(c, value), right)
        tip = T(f"Korrigiert in {cell}: {right}.", f"Corrected in {cell}: {right}.")
        return Finding(f"{'typo' if similar else 'mapping'}:{address(pos)}", text, level, tip, fix=fix), {pos}
    # Zwei Korrekturen passen gleich gut: Code und Bezeichnung widersprechen sich, und sonst sagt nichts, welche stimmt.
    if len(repairs) > 2 and repairs[2][0] == repairs[0][0]:
        return None
    (_, c1, v1, imp1, _), (_, c2, v2, imp2, _) = sorted(repairs[:2], key=lambda t: t[1])
    p1, p2 = (r, c1), (r, c2)
    # Erst, was zum Wert links gehört („Zu „E11.9“ gehört sonst immer „Diabetes mellitus Typ 2““), dann umgekehrt.
    first = _because(data, r, c2, imp2, v2)
    second = _because(data, r, c1, imp1, v1, again=True)
    one, two = f"{address(p1)} ({data.header(c1)})", f"{address(p2)} ({data.header(c2)})"
    text = T(f"{one} und {two} passen nicht zusammen: "
             f"{first}, {second}. Eins von beiden ist falsch; welches, sagt der Rest der Zeile nicht.",
             f"{one} and {two} do not match: {first}, and {second}. One of them is wrong; the rest of the row "
             "does not say which.")
    choices = []
    for pos, c, value in ((p2, c2, v2), (p1, c1, v1)):
        right = data.shown(c, value)
        fix = _fix(sheet, {pos: data.value(c, value)}, data.like(c, value), right)
        if fix is None or (c not in data.numbers and _risky(right)):
            return None
        choices.append({"label": f"{address(pos)}: {right}", "fix": fix})
    to2, to1 = data.q(c2, data.shown(c2, v2)), data.q(c1, data.shown(c1, v1))
    tip = T(f"Entweder {address(p2)} auf {to2} oder {address(p1)} auf "
            f"{to1} korrigieren, je nachdem, was in der Akte steht.",
            # Englisch ohne Komma nach der Zahl (5,1 bleibt sonst deutsch geschrieben, sprache._numbers).
            f"Depending on what the record says, correct either {address(p2)} to {to2} or {address(p1)} to {to1}.")
    return Finding(f"pair:{address(p1)}", text, "fehler", tip, cell=f"{address(p1)},{address(p2)}",
                   choices=choices), {p1, p2}




def _because(data: _Data, r: int, col: int, implied: dict, value: str, again: bool = False) -> str:
    """„Zu „E11.9“ gehört sonst immer „Diabetes mellitus Typ 2“ (8-mal)“; again: „zu „I10“ sonst immer …“"""
    x, (_, k) = sorted(implied.items(), key=lambda t: (-t[1][1], t[0]))[0]
    have, want = data.q(x, data.raw[x][r].strip()), data.q(col, data.shown(col, value))
    if again:
        return T(f"zu {have} sonst immer {want} ({k}-mal)", f"{have} always with {want} ({_times(k)})")
    return T(f"Zu {have} gehört sonst immer {want} ({k}-mal)",
             f"everywhere else, {have} goes with {want} ({_times(k)})")


def _times(k: int) -> str:
    return "once" if k == 1 else f"{k} times"


def _should(right: str, now: str) -> str:
    """„ Richtig wäre „X“ statt „Y“.“"""
    return T(f" Richtig wäre {right} statt {now}.", f" It should be {right} instead of {now}.")


def _reasons(data: _Data, r: int, col: int, implied: dict, value: str) -> str:
    """„Zu „Dr. Müller“ (Arzt) gehört sonst immer „Innere Medizin“ (27-mal), zu „J06.9“ (ICD-10) ebenso“"""
    ordered = sorted(implied.items(), key=lambda t: (-t[1][1], t[0]))
    (x, (_, k)), rest = ordered[0], ordered[1:]
    have, want = f"{data.q(x, data.raw[x][r].strip())} ({data.header(x)})", data.q(col, data.shown(col, value))
    said = T(f"Zu {have} gehört sonst immer "
             f"{want} ({k}-mal)",
             f"everywhere else, {have} goes with {want} ({_times(k)})")
    if rest:
        also = _listed([f'{data.q(o, data.raw[o][r].strip())} ({data.header(o)})' for o, _ in rest])
        said += T(f", zu {also} ebenso", f", as {'does' if len(rest) == 1 else 'do'} {also}")
    return said


def _listed(parts: list[str]) -> str:
    return parts[0] if len(parts) == 1 else ", ".join(parts[:-1]) + T(" und ", " and ") + parts[-1]


def _osa(a: str, b: str) -> int:
    """Wie viele Buchstaben man ändern, einfügen, löschen oder vertauschen muss (Damerau-Levenshtein)."""
    if abs(len(a) - len(b)) > 2:
        return 3
    prev2, prev = None, list(range(len(b) + 1))
    for i in range(1, len(a) + 1):
        cur = [i] + [0] * len(b)
        for j in range(1, len(b) + 1):
            cost = a[i - 1] != b[j - 1]
            cur[j] = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost)
            if i > 1 and j > 1 and a[i - 1] == b[j - 2] and a[i - 2] == b[j - 1]:
                cur[j] = min(cur[j], prev2[j - 2] + 1)
        prev2, prev = prev, cur
    return prev[-1]


def _typo(rare: str, often: str) -> bool:
    """„dr. shmidt“ und „dr. schmidt“: ein Buchstabe anders, nicht eine andere Zahl („Station 3“, „Typ 2“) und nicht
    ein anderes kurzes Wort („Gruppe A“, „Gruppe B“)."""
    a, b = _norm(rare), _norm(often)
    if a == b or min(len(a), len(b)) < 4 or not re.search(r"[^\W\d_]", a):
        return False
    if re.sub(r"\D", "", a) != re.sub(r"\D", "", b):
        return False
    wa, wb = a.split(), b.split()
    if len(wa) == len(wb) and any(x != y and min(len(x), len(y)) <= 2 for x, y in zip(wa, wb)):
        return False
    return _osa(a, b) <= (1 if min(len(a), len(b)) < 9 else 2)


def _typos(data: _Data, col: int, taken: set[Pos]) -> list[tuple[Finding, set[Pos]]]:
    """Ein seltener Wert, der fast so aussieht wie ein häufiger: „Dr. Shmidt“ (1-mal) neben „Dr. Schmidt“ (16-mal)."""
    counts = data.counts[col]
    n = sum(counts.values())
    often = [k for k, c in counts.items() if c >= 3]
    out = []
    for rare, c in sorted(counts.items()):
        if c > 2 or c > n / 10:
            continue
        near = [k for k in often if counts[k] >= 3 * c and _typo(rare, k)]
        if len(near) != 1:
            continue
        target = near[0]
        cells = [(r, col) for r, k in data.keys[col].items() if k == rare and (r, col) not in taken]
        if len(cells) != c:
            continue
        # Sagen die übrigen Spalten der Zeile etwas anderes als den häufigen Wert, ist es wohl ein anderer Eintrag.
        if any(_implied(data, r, col) not in (None, target) for r, _ in cells):
            continue
        now, right = data.raw[col][cells[0][0]].strip(), data.shown(col, target)
        if _risky(right):
            continue
        where = f"{address(cells[0])} ({data.header(col)})" if c == 1 else \
            f"{T(' und ', ' and ').join(address(p) for p in cells)} ({data.header(col)})"
        n = counts[target]
        lead = T(f"{where} „{now}“ steht nur {'hier' if c == 1 else 'dort'}, „{right}“ dagegen {n}-mal: "
                 "wohl ein Tippfehler.",
                 f"{where} contains “{now}”, which appears only {'here' if c == 1 else 'there'}, while “{right}” "
                 f"appears {_times(n)}: probably a typo.")
        effect = T(f" Eine Pivot-Tabelle, ZÄHLENWENN oder SUMMEWENN nach {data.header(col)} zählt "
                   f"{'die Zeile' if c == 1 else 'die Zeilen'} darum für sich.",
                   f" A pivot table, COUNTIF or SUMIF by {data.header(col)} therefore counts "
                   f"{'the row' if c == 1 else 'the rows'} separately.")
        effect += _effects(data.sheet, {p: right for p in cells})
        text = lead + effect + _should(T(f"„{right}“", f"“{right}”"), T(f"„{now}“", f"“{now}”"))
        tip = T(f"Korrigiert in {address(cells[0])}: {right}.", f"Corrected in {address(cells[0])}: {right}.") \
            if c == 1 else T(f"Korrigiert: {right}.", f"Corrected: {right}.")
        out.append((Finding(f"typo:{address(cells[0])}", text, "pruefen", tip,
                            fix=_fix(data.sheet, {p: right for p in cells})), set(cells)))
    return out


def _implied(data: _Data, r: int, col: int) -> str | None:
    """Was die übrigen Spalten der Zeile für col sagen (über die Regeln), wenn sie es einig sagen."""
    said = {mapping[data.keys[x][r]][0] for (x, y), mapping in data.rules.items()
            if y == col and r in data.keys[x] and data.keys[x][r] in mapping}
    return said.pop() if len(said) == 1 else None


def _variants(data: _Data, col: int, taken: set[Pos]) -> list[tuple[Finding, set[Pos]]]:
    """Derselbe Wert anders geschrieben: groß und klein („kardiologie“), Leerzeichen am Rand oder doppelt."""
    counts = Counter(data.raw[col].values())
    out = []
    for raw, c in sorted(counts.items()):
        key = _norm(raw)
        right = data.spelling[col][key]
        if raw == right or counts[right] < 3 or counts[right] < 2 * c:
            continue
        cells = [(r, col) for r, t in data.raw[col].items() if t == raw and (r, col) not in taken]
        if not cells or _risky(right):
            continue
        spaces = raw != " ".join(raw.split())
        # how: wie der Satz für eine Zelle es sagt; tail: dasselbe hinter dem Doppelpunkt für mehrere Zellen.
        if spaces:
            if raw.rstrip() != raw and raw.lstrip() == raw:
                how, tail = T("hat ein Leerzeichen am Ende", "has a trailing space"), \
                    T("mit ein Leerzeichen am Ende", "with a trailing space")
            elif raw.lstrip() != raw and raw.rstrip() == raw:
                how, tail = T("hat ein Leerzeichen am Anfang", "has a leading space"), \
                    T("mit ein Leerzeichen am Anfang", "with a leading space")
            else:
                how, tail = T("hat Leerzeichen, wo keine hingehören", "has spaces where they do not belong"), \
                    T("mit Leerzeichen, wo keine hingehören", "with spaces where they do not belong")
            why = T(f" Eine Pivot-Tabelle, ZÄHLENWENN und SUMMEWENN zählen {'die Zeile' if len(cells) == 1 else 'die Zeilen'}"
                    " darum für sich.",
                    f" A pivot table, COUNTIF and SUMIF therefore count {'the row' if len(cells) == 1 else 'the rows'}"
                    " separately.")
        else:
            if raw == raw.lower():
                how, tail = T("ist klein geschrieben", "is in lower case"), T("ist klein geschrieben", "in lower case")
            elif raw == raw.upper():
                how, tail = T("ist groß geschrieben", "is in upper case"), T("ist groß geschrieben", "in upper case")
            else:
                how, tail = T("ist anders geschrieben", "is spelled differently"), \
                    T("ist anders geschrieben", "spelled differently")
            why = T(" Excel zählt beides zusammen (Pivot-Tabelle, ZÄHLENWENN), R, Python und SPSS zählen es als eigene "
                    "Gruppe.",
                    " Excel counts both together (pivot table, COUNTIF), but R, Python and SPSS count it as a separate "
                    "group.")
        n = counts[right]
        if len(cells) == 1:
            lead = T(f"{address(cells[0])} ({data.header(col)}) „{raw}“ {how}, sonst steht {n}-mal „{right}“.",
                     f"{address(cells[0])} ({data.header(col)}) contains “{raw}”, which {how}; elsewhere “{right}” "
                     f"appears {_times(n)}.")
        else:
            listed = ", ".join(address(p) for p in cells[:MAX_LISTED]) + (
                T(f" und {len(cells) - MAX_LISTED} weitere", f" and {len(cells) - MAX_LISTED} more")
                if len(cells) > MAX_LISTED else "")
            lead = T(f"In {data.header(col)} steht {len(cells)}-mal „{raw}“ ({listed}), sonst {n}-mal "
                     f"„{right}“: {tail}.",
                     f"In {data.header(col)}, “{raw}” appears {_times(len(cells))} ({listed}), elsewhere “{right}” "
                     f"appears {_times(n)}: {tail}.")
        effect = _effects(data.sheet, {p: right for p in cells}) if spaces else ""
        text = lead + why + effect + _should(T(f"„{right}“", f"“{right}”"), T(f"„{raw}“", f"“{raw}”"))
        out.append((Finding(f"spelling:{address(cells[0])}", text, "fehler",
                            T(f"Korrigiert in {address(cells[0])}: {right}.", f"Corrected in {address(cells[0])}: {right}.")
                            if len(cells) == 1 else T(f"Korrigiert: {right}.", f"Corrected: {right}."),
                            fix=_fix(data.sheet, {p: right for p in cells})), set(cells)))
    return out


# Text, den Excel beim Eintippen in eine Zahl oder ein Datum verwandelt: nicht einsetzen.
_NUMBERISH = re.compile(r"[\d\s.,:/+\-%€$]*")


def _risky(text: str) -> bool:
    return bool(_NUMBERISH.fullmatch(text)) or text.startswith(("=", "+", "-", "@"))


def _fix(sheet: Sheet, changes: dict[Pos, str | float], like: Pos | None = None, said: str = "") -> dict | None:
    """Zum Einsetzen: der Text, wie er in den übrigen Zeilen steht, oder die Zahl im Format der Spalte (like)."""
    from . import einsetzen

    if not changes or len(changes) > einsetzen.MAX_CELLS:
        return None
    if any(isinstance(v, float) for v in changes.values()):
        return einsetzen.values(sheet, changes, like, said)
    targets = sorted(changes, key=lambda p: (p[1], p[0]))
    was = {}
    for pos in targets:
        cell = sheet.cells.get(pos)
        if cell is None or cell.formula:
            return None
        was[address(pos)] = {"text": cell.text if cell.text is not None else cell.value}
    first = changes[targets[0]]
    return {"range": einsetzen._area(targets), "cells": {address(p): changes[p] for p in targets}, "was": was,
            "german": first, "text": True}


def _effects(sheet: Sheet, fixed: dict[Pos, str]) -> str:
    """„ Darum ergibt M5 (Behandlungen Dr. Schmidt) 15 statt 16.“: Formeln, die mit dem richtigen Text anders rechnen."""
    chain = {pos for pos in sheet.nodes if pos not in sheet.unknown}
    if not chain or len(chain) > MAX_FORMULAS:
        return ""
    _, now = sheet.evaluator(chain=chain)
    _, right = sheet.evaluator(overrides=fixed, chain=chain)
    changed = []
    for pos in sorted(chain, key=lambda p: (sheet.plan.in_table(*p), p)):
        try:
            before, after = now(pos), right(pos)
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, IndexError, RecursionError):
            continue
        if isinstance(before, float) and isinstance(after, float):
            if math.isclose(before, after, rel_tol=1e-9, abs_tol=1e-9):
                continue
        elif before == after:
            continue
        changed.append((pos, before, after))
    if not changed:
        return ""
    shown = [T(f"{sheet.label(p)} {_value(b)} statt {_value(a)}", f"{sheet.label(p)} gives {_value(b)} instead of "
               f"{_value(a)}") for p, b, a in changed[:MAX_LISTED]]
    extra = len(changed) - MAX_LISTED
    more = T(more_formulas(extra), " and one more formula" if extra == 1 else f" and {extra} more formulas") \
        if extra > 0 else ""
    return T(f" Darum ergibt {', '.join(shown)}{more}.", f" As a result, {'; '.join(shown)}{more}.")


def _value(value) -> str:
    if isinstance(value, str):
        return T(f"„{value}“", f"“{value}”")
    if isinstance(value, bool):
        return T(said(value), "TRUE" if value else "FALSE")
    if isinstance(value, Err):
        return T(said(value), value.code)
    return said(value)
