"""Was der Name sagt, gegen das, was die Formel rechnet: die Fehler, die in der ganzen Spalte gleich falsch sind.

Ein Musterbruch fällt auf, weil eine Zelle anders rechnet als ihre Nachbarn. Hier rechnen alle gleich, aber anders,
als der Spaltenkopf oder das Wort daneben sagt: „Abweichung Ist − Plan“ rechnet Plan − Ist, „Netto nach Rabatt“
zählt den Rabatt dazu, „Dauer (min)“ ergibt Stunden, „Ø 3 Monate“ nimmt 4, „Index (Jan = 100)“ teilt durch Februar.

Ohne KI: Der Prüfer liest Wörter und Einheiten aus den Köpfen, die Form der Formel und die Zahlen. Wo die Mappe
nicht sagt, was gemeint ist, heißt es „prüfen“; wo sie es sagt, „Fehler“ mit der Korrektur, nachgerechnet.
"""

from __future__ import annotations

import logging
import re

from . import formel, woerter
from .analyse import Finding, Sheet, _fmt, address
from .sprache import T, english
from .workbook import _left_label, pattern
from .xlsx import column_letters

Pos = tuple[int, int]
log = logging.getLogger(__name__)
MAX_EACH = 2

_FILL = {"der", "die", "das", "des", "den", "dem", "von", "vom", "in", "im", "für", "aus", "und", "nach", "zum", "zur",
         "je", "pro", "the", "of", "for", "and", "per", "eur", "euro", "chf", "usd", "fr", "t", "tsd", "k", "mio", "h",
         "min", "%", "€"}


# Sprache: Ein Kopf heißt „Ist-Plan“, „Ist − Plan“ oder „IST minus PLAN“, „Fälle“, „Faelle“ oder „Fall“,
# „Gesamtumsatz“ oder „Umsatz gesamt“. Verglichen wird darum ohne Groß/klein, Umlaut gleich Grundlaut (ä = ae = a),
# ohne Leerzeichen und Bindestriche und mit dem Wortstamm, und ein Wort zählt auch in einem zusammengesetzten.


def _plain(text: str, keep_spaces: bool = False) -> str:
    text = woerter.spaced(text).casefold().replace("ß", "ss")
    for a, b in (("ä", "a"), ("ö", "o"), ("ü", "u"), ("ae", "a"), ("oe", "o"), ("ue", "u")):
        text = text.replace(a, b)
    return re.sub(r"\s+", " ", text) if keep_spaces else re.sub(r"[^0-9a-zø%€$£]+", "", text)


def _stem(word: str) -> str:
    w = _plain(word)
    for end in ("en", "er", "es", "e", "n", "s"):
        if len(w) - len(end) >= 4 and w.endswith(end):
            return w[: -len(end)]
    return w


def _words(text: str) -> list[str]:
    return re.findall(r"[^\W_]+", woerter.spaced(text).casefold())


def _has(text: str, words) -> bool:
    """Steckt eines der Wörter im Text, auch in einem zusammengesetzten Wort („Gesamtumsatz“) oder anders
    geschrieben? Kurze Wörter (VK, EK) nur als ganzes Wort."""
    whole = _plain(text)
    tokens = {_stem(w) for w in _words(text)}
    for word in words:
        stem = _stem(word)
        if (len(stem) >= 4 and stem in whole) or stem in tokens:
            return True
    return False


def _key(text: str) -> set[str]:
    return {_stem(w) for w in _words(re.sub(r"\([^)]*\)", " ", text or "")) if w not in _FILL and len(w) > 1}


def _covers(words: set[str], text: str) -> bool:
    """Alle Wörter (Stämme) kommen im Text vor, auch zusammengesetzt."""
    return bool(words) and all(_has(text, [w]) for w in words)


def _rx(pattern_: str) -> re.Pattern:
    """Ein Muster, das so liest wie _plain: Umlaute als Grundlaut, Leerzeichen dürfen fehlen oder ein Bindestrich sein."""
    folded = _plain(pattern_, keep_spaces=True).replace(" ", r"[\s\-_.]*")
    return re.compile(folded, re.IGNORECASE)


def _search(rx: re.Pattern, text: str):
    return rx.search(_plain(text, keep_spaces=True))


class _Ctx:
    def __init__(self, sheet: Sheet, rows: list[int]):
        self.sheet, self.rows, self.data = sheet, rows, set(rows)
        self.plan = sheet.plan
        self.calc, self.now = sheet.evaluator()
        self.runs = _runs(sheet, self.data)

    def name(self, pos: Pos) -> str:
        if pos[0] in self.data and pos[1] in self.plan.headers:
            return self.plan.headers[pos[1]]
        # Eine Spalte neben der Tabelle mit eigenem Kopf in derselben Zeile („Ø 3 Monate“ in F4).
        top = self.sheet.cells.get((self.plan.header_row, pos[1])) if self.plan.header_row else None
        if pos[0] in self.data and top is not None and not top.formula and isinstance(top.value, str) and \
                re.search(r"[^\W\d_]", top.value):
            return top.value.strip()
        return _left_label(self.sheet.cells, *pos)

    def head(self, col: int) -> str:
        return self.plan.headers.get(col, "")

    def value(self, pos: Pos):
        try:
            return self.now(pos)
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
            return None

    def run(self, text: str):
        try:
            return self.calc.run(formel.parse(text))
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, IndexError, KeyError, RecursionError):
            return None

    def fix(self, changes: dict[Pos, str], **extra) -> dict | None:
        from . import einsetzen

        found = einsetzen.formulas(self.sheet, changes)
        if found is not None:
            found.update(extra)
        return found

    def label(self, cells: list[Pos]) -> str:
        where = address(cells[0]) if len(cells) == 1 else f"{address(cells[0])}:{address(cells[-1])}"
        return f"{where} ({self.name(cells[0])})"

    def german(self, text: str) -> str:
        from .bereiche import _german_formula

        return _german_formula(text)


def _runs(sheet: Sheet, data: set[int]) -> list[list[Pos]]:
    """Formeln untereinander mit derselben Form; eine Formel allein (Kennzahl) ist ein Lauf aus einer Zelle."""
    by_col: dict[int, list[int]] = {}
    for (r, c), cell in sheet.cells.items():
        if cell.formula and (r, c) in sheet.nodes and (r, c) not in sheet.unknown and sheet.nodes[(r, c)] is not None:
            by_col.setdefault(c, []).append(r)
    runs = []
    for c, rows in sorted(by_col.items()):
        run: list[Pos] = []
        for r in sorted(rows):
            here = pattern(sheet.cells[(r, c)].formula, (r, c))
            if run and r == run[-1][0] + 1 and here == pattern(sheet.cells[run[-1]].formula, run[-1]):
                run.append((r, c))
            else:
                if run:
                    runs.append(run)
                run = [(r, c)]
        if run:
            runs.append(run)
    return runs


_REF = r"(?<![A-Za-z0-9_.$!:])(\$?){col}(\$?){row}(?![0-9A-Za-z(!:])"


def _swap_ref(text: str, old: Pos, new: str) -> str | None:
    """Den Bezug old (mit oder ohne $) einmal durch new ersetzen; None, wenn er nicht genau einmal dasteht."""
    pattern_ = re.compile(_REF.format(col=column_letters(old[1]), row=old[0]))
    if len(pattern_.findall(text)) != 1:
        return None
    return pattern_.sub(lambda m: new.replace("{d1}", m.group(1)).replace("{d2}", m.group(2)), text)


def _close(a, b) -> bool:
    return isinstance(a, float) and isinstance(b, float) and abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))


def _changes(ctx: _Ctx, run: list[Pos], make) -> dict[Pos, str] | None:
    """make(text, pos) → neue Formel je Zelle des Laufs, oder None, wenn es für eine nicht geht."""
    out = {}
    for pos in run:
        new = make(ctx.sheet.cells[pos].formula or "", pos)
        if new is None:
            return None
        out[pos] = new
    return out


def _wrong(ctx: _Ctx, changes: dict[Pos, str]) -> tuple[Pos, object, object] | None:
    """Die erste Zelle, die mit der Korrektur etwas anderes ergibt: (Zelle, jetzt, richtig)."""
    for pos, text in changes.items():
        now, right = ctx.value(pos), ctx.run(text)
        if isinstance(now, float) and isinstance(right, float) and not _close(now, right):
            return pos, now, right
    return None


def _shown(ctx: _Ctx, pos: Pos, value) -> str:
    cell = ctx.sheet.cells.get(pos)
    if isinstance(value, float) and cell is not None and str(cell.value).strip().endswith("%"):
        return f"{_fmt(round(value * 100, 2))} %"
    return _fmt(round(value, 4)) if isinstance(value, float) else _fmt(value)


def _finding(kind: str, ctx: _Ctx, run: list[Pos], text: str, level: str, changes: dict[Pos, str] | None,
             said: tuple[str, str] = ("", ""), **extra) -> Finding:
    """said: ein Satz vor dem Tipp, (in der Sprache der Anfrage, auf Deutsch)."""
    fix = ctx.fix(changes, **extra) if changes else None
    first = changes[run[0]] if changes else ""
    tip = german = ""
    if changes:
        german = (f"Schreib in {address(run[0])} ={ctx.german(first)}"
                  + (f" und zieh es herunter bis {address(run[-1])}." if len(run) > 1 else "."))
        tip = T(german,
                f"In {address(run[0])}, enter ={ctx.german(first)}"
                + (f" and fill it down to {address(run[-1])}." if len(run) > 1 else "."))
    if said[0]:
        tip, german = f"{said[0]} {tip}".strip(), f"{said[1]} {german}".strip()
    item = Finding(f"{kind}:{address(run[0])}", text, level, tip, fix=fix)
    _fix_from_tip(ctx.sheet, item, german)
    return item


def _tip_pair(german: str, english_text: str) -> tuple[str, str]:
    """Ein Tipp in der Sprache der Anfrage und auf Deutsch (für _fix_from_tip)."""
    return T(german, english_text), german


def _tip_plus(a: tuple[str, str], b: tuple[str, str]) -> tuple[str, str]:
    return a[0] + b[0], a[1] + b[1]


def _fix_from_tip(sheet: Sheet, item: Finding, german: str) -> None:
    """Ohne eigene Korrektur liest einsetzen.attach sie später aus dem Tipp, aber nur aus einem deutschen
    („Korrigiert: =…“, „Schreib in … und zieh es herunter“). Auf Englisch darum gleich hier aus dem deutschen Tipp:
    dieselbe Korrektur wie auf Deutsch."""
    if english() and not item.fix and "=" in german:
        from . import einsetzen

        item.fix = einsetzen.fix(sheet, Finding(item.key, item.text, item.level, german))


def _two(node):
    """X − Y (auch mal 100 oder geteilt): die beiden Bezüge, sonst None."""
    if node.kind == "op" and node.value in ("*", "/") and len(node.args) == 2 and node.args[1].kind in ("num", "ref"):
        node = node.args[0]
    if node.kind == "op" and node.value == "-" and len(node.args) == 2 and all(a.kind == "ref" for a in node.args):
        return node.args[0].value, node.args[1].value
    return None


# --- 26: Vorzeichen gegen „A − B“ im Kopf ---

_MINUS = re.compile(r"([^()\-−–:=]+?)\s*(?:−|–|-|minus)\s*([^()\-−–:=]+)")


def _sign_order(ctx: _Ctx) -> list[Finding]:
    found = []
    for run in ctx.runs:
        first = run[0]
        pair = _two(ctx.sheet.nodes[first])
        if pair is None:
            continue
        name = ctx.name(first)
        inner = re.search(r"\(([^)]*)\)", name)
        for part in ([inner.group(1)] if inner else []) + [name]:
            m = _MINUS.search(part)
            if not m:
                continue
            # Das Wort direkt vor und direkt nach dem Minus: „Abweichung Ist-Plan“ → Ist, Plan.
            left_words, right_words = _words(m.group(1)), _words(m.group(2))
            if not left_words or not right_words:
                continue
            a, b = {_stem(left_words[-1])}, {_stem(right_words[0])}
            na, nb = ctx.name(pair[0]), ctx.name(pair[1])
            if _covers(a, nb) and _covers(b, na) and not (_covers(a, na) or _covers(b, nb)):
                break
        else:
            continue
        left, right = pair

        def make(text, pos, left=left, right=right):
            dr = pos[0] - first[0]
            lpos, rpos = (left[0] + dr, left[1]), (right[0] + dr, right[1])
            swapped = _swap_ref(text, lpos, "\0L")
            swapped = swapped and _swap_ref(swapped, rpos, "\0R")
            if not swapped:
                return None
            return swapped.replace("\0L", address(rpos)).replace("\0R", address(lpos))

        changes = _changes(ctx, run, make)
        bad = changes and _wrong(ctx, changes)
        if not bad:
            continue
        pos, now, right_value = bad
        text = T(f"{ctx.label(run)} heißt „{m.group(0).strip()}“, die Formel rechnet aber umgekehrt "
                 f"{ctx.name(left)} − {ctx.name(right)}: Das Vorzeichen ist vertauscht. Richtig wäre in {address(pos)} "
                 f"{_shown(ctx, pos, right_value)} statt {_shown(ctx, pos, now)}.",
                 f"{ctx.label(run)} is called “{m.group(0).strip()}”, but the formula calculates the other way round, "
                 f"{ctx.name(left)} − {ctx.name(right)}: the sign is reversed. In {address(pos)}, it should be "
                 f"{_shown(ctx, pos, right_value)} instead of {_shown(ctx, pos, now)}.")
        found.append(_finding("sign", ctx, run, text, "fehler", changes))
        if len(found) >= MAX_EACH:
            break
    return found


# --- 27: Kosten schon negativ und noch einmal abgezogen; 28: Rabatt zugezählt ---

_COSTS = ("kosten", "aufwand", "ausgaben", "ausgabe", "costs", "cost", "expenses", "abzüge", "spesen")
_DISCOUNT = ("rabatt", "skonto", "nachlass", "abzug", "discount", "preisnachlass", "gutschrift")
_AFTER = ("netto", "nach", "abzüglich", "abzgl", "zahlbetrag", "endpreis", "net", "after", "zu zahlen", "total",
          "gesamt", "betrag")


def _signed_terms(node, sign=1):
    """B5 − C5 + D5 → [(+1, B5), (−1, C5), (+1, D5)]; None bei anderem als + und − zwischen Bezügen."""
    if node.kind == "ref":
        return [(sign, node.value)]
    if node.kind == "op" and node.value in ("+", "-") and len(node.args) == 2:
        left = _signed_terms(node.args[0], sign)
        right = _signed_terms(node.args[1], sign if node.value == "+" else -sign)
        return left + right if left is not None and right is not None else None
    return None


def _column_values(ctx: _Ctx, col: int) -> list[float]:
    return [v for r in ctx.rows if isinstance(v := ctx.value((r, col)), float) and not isinstance(v, bool)]


def _plus_minus(ctx: _Ctx) -> list[Finding]:
    found = []
    for run in ctx.runs:
        first = run[0]
        if first[0] not in ctx.data:
            continue
        terms = _signed_terms(ctx.sheet.nodes[first])
        if not terms or len(terms) < 2:
            continue
        name = ctx.name(first)
        for sign, ref in terms:
            if ref[0] != first[0]:
                continue
            head = ctx.name(ref)
            values = [v for v in _column_values(ctx, ref[1]) if v != 0]
            if len(values) < 3:
                continue
            negative = sum(v < 0 for v in values) >= 0.8 * len(values)
            positive = all(v > 0 for v in values)
            # „Netto nach Rabatt“ darf „Rabatt“ im Namen haben, „Rabatt total“ nicht.
            discount = _has(head, _DISCOUNT) and positive and sign > 0 and _has(name, _AFTER) and \
                (not _has(name, _DISCOUNT) or _has(name, ("nach", "abzüglich", "abzgl", "after", "less", "minus")))
            twice = _has(head, _COSTS) and negative and sign < 0 and not _has(name, _COSTS)
            if not (discount or twice):
                continue
            op = "+" if twice else "-"

            def make(text, pos, ref=ref, op=op):
                target = (pos[0], ref[1])
                pat = re.compile(r"([+\-])(\s*)" + _REF.format(col=column_letters(target[1]), row=target[0]))
                hits = pat.findall(text)
                if len(hits) != 1:
                    return None
                return pat.sub(lambda m: f"{op}{m.group(2)}{m.group(3)}{column_letters(target[1])}{m.group(4)}{target[0]}",
                               text)

            changes = _changes(ctx, run, make)
            bad = changes and _wrong(ctx, changes)
            if not bad:
                continue
            pos, now, right = bad
            col = column_letters(ref[1])
            if twice:
                text = T(f"{ctx.label(run)} zieht {col} „{ctx.name(ref)}“ ab, aber die Kosten stehen dort schon negativ "
                         f"(etwa {_fmt(ctx.value((pos[0], ref[1])))}). Minus mal minus zählt sie dazu statt ab. Richtig "
                         f"wäre in {address(pos)} {_fmt(right)} statt {_fmt(now)}.",
                         f"{ctx.label(run)} subtracts {col} “{ctx.name(ref)}”, but the costs there are already negative "
                         f"(e.g. {_fmt(ctx.value((pos[0], ref[1])))}). Minus times minus adds them instead of subtracting "
                         f"them. In {address(pos)}, it should be {_fmt(right)} instead of {_fmt(now)}.")
                kind = "negcost"
            else:
                text = T(f"{ctx.label(run)} zählt {col} „{ctx.name(ref)}“ dazu, statt ihn abzuziehen. Richtig wäre in "
                         f"{address(pos)} {_fmt(right)} statt {_fmt(now)}.",
                         f"{ctx.label(run)} adds {col} “{ctx.name(ref)}” instead of subtracting it. In "
                         f"{address(pos)}, it should be {_fmt(right)} instead of {_fmt(now)}.")
                kind = "rabatt"
            found.append(_finding(kind, ctx, run, text, "fehler", changes))
            break
        if len(found) >= MAX_EACH:
            break
    return found


# --- 7: „inkl. MwSt.“ oder „Brutto“, aber ohne MwSt. gerechnet ---

_GROSS = _rx(r"brutto|inkl\.?\s*(?:mwst|ust|mehrwertsteuer)|incl\.?\s*vat|gross|mit mwst")
_NET = _rx(r"netto|exkl\.?|ohne mwst|zzgl|excl|net\b")
_VAT = _rx(r"mwst|ust\b|mehrwertsteuer|vat|tva|iva")


def _vat_rate(ctx: _Ctx) -> Pos | None:
    """Die Zelle mit dem Steuersatz: eine Zahl unter 1 (oder in %) neben „MwSt“."""
    for pos, cell in ctx.sheet.cells.items():
        if cell.formula or pos[0] in ctx.data:
            continue
        v = ctx.sheet.inputs.get(pos)
        if isinstance(v, float) and 0 < v < 0.3 and _search(_VAT, _left_label(ctx.sheet.cells, *pos) or ""):
            return pos
    return None


def _gross(ctx: _Ctx) -> list[Finding]:
    found = []
    for run in ctx.runs:
        first = run[0]
        name = ctx.name(first)
        if not _search(_GROSS, name) or _search(_NET, name):
            continue
        node = ctx.sheet.nodes[first]
        if node.kind != "ref" or node.value[0] != first[0] or not _search(_NET, ctx.name(node.value)):
            continue
        rate = _vat_rate(ctx)
        net = node.value
        changes = None
        if rate is not None:
            fixed = f"${column_letters(rate[1])}${rate[0]}"
            changes = {pos: f"{address((pos[0], net[1]))}*(1+{fixed})" for pos in run}
        text = T(f"{ctx.label(run)} übernimmt nur {column_letters(net[1])} „{ctx.name(net)}“: Der Betrag ist ohne "
                 f"MwSt., der Name sagt aber mit MwSt."
                 + (f" Mit dem Satz in {address(rate)} ({_left_label(ctx.sheet.cells, *rate)})." if rate else
                    " Rechne mal (1 + Steuersatz)."),
                 f"{ctx.label(run)} only takes {column_letters(net[1])} “{ctx.name(net)}”: the amount excludes "
                 f"VAT, but the name says including VAT."
                 + (f" Use the rate in {address(rate)} ({_left_label(ctx.sheet.cells, *rate)})." if rate else
                    " Multiply by (1 + tax rate)."))
        found.append(_finding("gross", ctx, run, text, "fehler" if changes else "pruefen", changes))
    return found[:MAX_EACH]


# --- 4: Prozent doppelt (*100 und Prozentformat); 32: Anteil ohne Prozentformat ---


def _times(node, factor: float):
    """X*100 oder 100*X: (X, True); sonst (node, False)."""
    if node.kind == "op" and node.value == "*" and len(node.args) == 2:
        a, b = node.args
        if b.kind == "num" and b.value == factor:
            return a, True
        if a.kind == "num" and a.value == factor:
            return b, True
    return node, False


_PCT = _rx(r"%|prozent|percent|anteil|quote")
_PCT_SIGN = _rx(r"%|prozent|percent")


def _percent(ctx: _Ctx) -> list[Finding]:
    found = []
    like = next((p for p, c in sorted(ctx.sheet.cells.items()) if c.formula and str(c.value).strip().endswith("%")), None)
    for run in ctx.runs:
        first = run[0]
        cell = ctx.sheet.cells[first]
        node = ctx.sheet.nodes[first]
        inner, hundred = _times(node, 100.0)
        ratio = inner.kind == "op" and inner.value == "/"
        shows_percent = str(cell.value).strip().endswith("%")
        if shows_percent and hundred and ratio:
            def make(t, pos):
                new = re.sub(r"\s*\*\s*100(?![0-9.,])", "", t, count=1)
                if new == t:
                    new = re.sub(r"(?<![0-9.,])100\s*\*\s*", "", t, count=1)
                return new if new != t else None

            changes = _changes(ctx, run, make)
            bad = changes and _wrong(ctx, changes)
            if not bad:
                continue
            pos, now, right = bad
            msg = T(f"{ctx.label(run)} rechnet mal 100 und zeigt das Ergebnis auch noch als Prozent: So steht dort "
                    f"{_fmt(round(now * 100, 2))} % statt {_fmt(round(right * 100, 2))} %. Das Prozentformat rechnet "
                    "selbst mal 100.",
                    f"{ctx.label(run)} multiplies by 100 and also shows the result as a percentage, so it shows "
                    f"{_fmt(round(now * 100, 2))} % instead of {_fmt(round(right * 100, 2))} %. The percent format "
                    "already multiplies by 100 itself.")
            found.append(_finding("percent2", ctx, run, msg, "fehler", changes))
            continue
        name = ctx.name(first)
        if shows_percent or hundred or not ratio or not _search(_PCT_SIGN, name):
            continue
        values = [ctx.value(p) for p in run]
        if not values or not all(isinstance(v, float) and 0 < abs(v) <= 1.5 for v in values):
            continue
        if like is not None:
            changes = {p: ctx.sheet.cells[p].formula for p in run}
            msg = T(f"{ctx.label(run)} heißt in Prozent, zeigt aber {_fmt(round(values[0], 4))} statt "
                    f"{_fmt(round(values[0] * 100, 2))} %: Der Zelle fehlt das Prozentformat.",
                    f"{ctx.label(run)} is named in percent, but shows {_fmt(round(values[0], 4))} instead of "
                    f"{_fmt(round(values[0] * 100, 2))} %: the cell is missing the percent format.")
            item = _finding("percentfmt", ctx, run, msg, "fehler", changes,
                            said=(T("Mit dem Prozentformat wie in " + address(like) + ".",
                                    f"With the percent format as in {address(like)}."),
                                  "Mit dem Prozentformat wie in " + address(like) + "."),
                            format={address(p): address(like) for p in run})
            if item.fix and "format" in item.fix:  # nicht die aus dem Tipp gelesene Formel (_fix_from_tip)
                item.fix["german"] = T(f"Prozentformat wie {address(like)}", f"Percent format like {address(like)}")
            found.append(item)
        else:
            changes = _changes(ctx, run, lambda t, pos: f"({t})*100")
            msg = T(f"{ctx.label(run)} heißt in Prozent, zeigt aber den Anteil {_fmt(round(values[0], 4))} statt "
                    f"{_fmt(round(values[0] * 100, 2))}.",
                    f"{ctx.label(run)} is named in percent, but shows the fraction {_fmt(round(values[0], 4))} instead "
                    f"of {_fmt(round(values[0] * 100, 2))}.")
            found.append(_finding("percentfmt", ctx, run, msg, "fehler", changes))
        if len(found) >= MAX_EACH:
            break
    return found


# --- 5: Zeitdifferenz in der falschen Einheit ---

_TIME = re.compile(r"^\d{1,2}:\d{2}(:\d{2})?$|^\d{2}\.\d{2}\.\d{4}( \d{1,2}:\d{2}(:\d{2})?)?$")
_UNITS = [(_rx(r"\((?:in\s+)?(?:min|minuten)\.?\)|\bminuten\b|\bmin\.?$|\bin\s*min\b"), 1440.0, "Minuten"),
          (_rx(r"\((?:in\s+)?(?:h|std\.?|stunden)\)|\bstunden\b|\bstd\.?(?:\b|$)|\bin\s*h\b"), 24.0, "Stunden"),
          (_rx(r"\((?:in\s+)?(?:s|sek\.?|sekunden)\)|\bsekunden\b"), 86400.0, "Sekunden"),
          (_rx(r"\((?:in\s+)?(?:tage|d)\)|\btage\b"), 1.0, "Tage")]
_UNIT_EN = {"Minuten": "minutes", "Stunden": "hours", "Sekunden": "seconds", "Tage": "days"}


def _time_unit(ctx: _Ctx) -> list[Finding]:
    found = []
    for run in ctx.runs:
        first = run[0]
        cell = ctx.sheet.cells[first]
        if _TIME.match(str(cell.value).strip()):
            continue  # Die Zelle zeigt selbst eine Uhrzeit: kein Faktor nötig.
        node = ctx.sheet.nodes[first]
        factor = 1.0
        if node.kind == "op" and node.value == "*" and len(node.args) == 2 and node.args[1].kind == "num":
            factor, node = node.args[1].value, node.args[0]
        if node.kind != "op" or node.value != "-" or not all(a.kind == "ref" for a in node.args):
            continue
        if not all(_TIME.match(str((ctx.sheet.cells.get(a.value) or formel.Node("num")).value if
                                   ctx.sheet.cells.get(a.value) else "").strip()) for a in node.args):
            continue
        name = ctx.name(first)
        units = [(f, w) for rx, f, w in _UNITS if _search(rx, name)]
        if len(units) != 1 or units[0][0] == factor:
            continue
        want, word = units[0]
        if factor not in (1.0, 24.0, 60.0, 1440.0, 86400.0):
            continue

        def make(t, pos, want=want, factor=factor):
            base = re.sub(r"\s*\*\s*" + re.escape(_fmt(factor).replace(",", ".")) + r"(?![0-9.,])", "", t, count=1) \
                if factor != 1.0 else t
            if factor != 1.0 and base == t:
                return None
            if want == 1.0:
                return base
            return f"{base}*{int(want)}" if re.fullmatch(r"\([^()]*\)", base.strip()) else f"({base})*{int(want)}"

        changes = _changes(ctx, run, make)
        bad = changes and _wrong(ctx, changes)
        if not bad:
            continue
        pos, now, right = bad
        got = {1.0: "Tage", 24.0: "Stunden", 60.0: "ein Sechzigstel Tag", 1440.0: "Minuten", 86400.0: "Sekunden"}[factor]
        got_en = {1.0: "days", 24.0: "hours", 60.0: "sixtieths of a day", 1440.0: "minutes", 86400.0: "seconds"}[factor]
        msg = T(f"{ctx.label(run)} soll {word} zeigen, die Formel ergibt aber {got}"
                + (" (Excel rechnet Zeiten in Tagen)" if factor == 1.0 else "")
                + f". Richtig wäre in {address(pos)} {_fmt(round(right, 4))} statt {_fmt(round(now, 4))}.",
                f"{ctx.label(run)} should show {_UNIT_EN[word]}, but the formula gives {got_en}"
                + (" (Excel calculates times in days)" if factor == 1.0 else "")
                + f". In {address(pos)}, it should be {_fmt(round(right, 4))} instead of {_fmt(round(now, 4))}.")
        found.append(_finding("timeunit", ctx, run, msg, "fehler", changes))
        if len(found) >= MAX_EACH:
            break
    return found


# --- 8: pro Monat gegen pro Jahr ---

_MONTHLY = _rx(r"pro monat|je monat|/\s*monat|monatlich|per month|monthly|\bmtl\.?")
_YEARLY = _rx(r"pro jahr|je jahr|/\s*jahr|jährlich|jahres(?!-?ø)|p\.\s*a\.|per year|annual|yearly")


def _period_unit(ctx: _Ctx) -> list[Finding]:
    found = []
    for run in ctx.runs:
        first = run[0]
        node = ctx.sheet.nodes[first]
        if node.kind != "ref":
            continue
        mine, theirs = ctx.name(first), ctx.name(node.value)
        if _search(_MONTHLY, mine) and _search(_YEARLY, theirs) and not _search(_MONTHLY, theirs):
            op, say = "/12", T(("pro Monat", "pro Jahr", "durch 12"), ("per month", "per year", "dividing by 12"))
        elif _search(_YEARLY, mine) and _search(_MONTHLY, theirs) and not _search(_MONTHLY, mine):
            op, say = "*12", T(("pro Jahr", "pro Monat", "mal 12"), ("per year", "per month", "multiplying by 12"))
        else:
            continue
        changes = _changes(ctx, run, lambda t, pos, op=op: f"{t}{op}")
        bad = changes and _wrong(ctx, changes)
        if not bad:
            continue
        pos, now, right = bad
        msg = T(f"{ctx.label(run)} soll {say[0]} sein, übernimmt aber {address(node.value)} „{theirs}“ ({say[1]}) "
                f"ohne {say[2]}. Richtig wäre in {address(pos)} {_fmt(round(right, 4))} statt {_fmt(round(now, 4))}.",
                f"{ctx.label(run)} should be {say[0]}, but takes {address(node.value)} “{theirs}” ({say[1]}) "
                f"without {say[2]}. In {address(pos)}, it should be {_fmt(round(right, 4))} instead of "
                f"{_fmt(round(now, 4))}.")
        found.append(_finding("period", ctx, run, msg, "fehler", changes))
    return found[:MAX_EACH]


# --- 14: Index mit der falschen Basis ---

_INDEX = re.compile(r"([^\W\d_][\w.]*|\d{4})\s*=\s*100", re.IGNORECASE)


def _index_base(ctx: _Ctx) -> list[Finding]:
    found = []
    names = {v.casefold().strip(): r for r, v in ctx.plan.names.items()}
    for run in ctx.runs:
        first = run[0]
        m = _INDEX.search(ctx.name(first))
        if not m or m.group(1).casefold() not in names:
            continue
        base_row = names[m.group(1).casefold()]
        inner, hundred = _times(ctx.sheet.nodes[first], 100.0)
        if not hundred or inner.kind != "op" or inner.value != "/" or not all(a.kind == "ref" for a in inner.args):
            continue
        top, bottom = inner.args[0].value, inner.args[1].value
        if top[1] != bottom[1] or bottom[0] == base_row or top[0] != first[0]:
            continue

        # Die Basis steht mit $ da ($B$6) und in jeder Zeile gleich; nur sie wird ersetzt, nicht B6 in Zeile 6.
        held = re.search(rf"(?<![A-Za-z0-9_.$!:])(\$?){column_letters(bottom[1])}\${bottom[0]}(?![0-9A-Za-z(!:])",
                         ctx.sheet.cells[first].formula or "")
        if not held:
            continue
        token = held.group(0)

        def make(t, pos, token=token, held=held, bottom=bottom):
            rx = re.compile(rf"(?<![A-Za-z0-9_.$!:]){re.escape(token)}(?![0-9A-Za-z(!:])")
            if len(rx.findall(t)) != 1:
                return None
            return rx.sub(f"{held.group(1)}{column_letters(bottom[1])}${base_row}", t)

        changes = _changes(ctx, run, make)
        bad = changes and _wrong(ctx, changes)
        if not bad:
            continue
        pos, now, right = bad
        msg = T(f"{ctx.label(run)} heißt „{m.group(0)}“, teilt aber durch {address(bottom)} "
                f"({ctx.plan.names.get(bottom[0], 'Zeile ' + str(bottom[0]))}) statt durch {address((base_row, bottom[1]))} "
                f"({ctx.plan.names[base_row]}). Richtig wäre in {address(pos)} {_fmt(round(right, 4))} statt "
                f"{_fmt(round(now, 4))}.",
                f"{ctx.label(run)} is called “{m.group(0)}”, but divides by {address(bottom)} "
                f"({ctx.plan.names.get(bottom[0], 'row ' + str(bottom[0]))}) instead of by "
                f"{address((base_row, bottom[1]))} ({ctx.plan.names[base_row]}). In {address(pos)}, it should be "
                f"{_fmt(round(right, 4))} instead of {_fmt(round(now, 4))}.")
        found.append(_finding("indexbase", ctx, run, msg, "fehler", changes))
    return found[:MAX_EACH]


# --- 16: gleitendes Fenster mit der falschen Länge ---

_SPAN = _rx(r"(?<!\d)(\d{1,2})\s*-?\s*(?:mon(?:ate?|\.)?|m\b|tage?|wochen?|jahre?|quartale?|perioden?|werte|"
                   r"zeilen|months?|days?|weeks?|years?)|(?:letzten?|last)\s+(\d{1,2})")
_MOVING = _rx(r"gleitend|rollierend|rollend|rolling|moving|trailing|ø|durchschnitt|mittel|summe|average|sum")


def _window_length(ctx: _Ctx) -> list[Finding]:
    found = []
    for run in ctx.runs:
        first = run[0]
        if len(run) < 2 or first[0] not in ctx.data:
            continue
        name = ctx.name(first)
        m = _search(_SPAN, name)
        if not m or not _search(_MOVING, name):
            continue
        want = int(m.group(1) or m.group(2))
        node = ctx.sheet.nodes[first]
        if node.kind != "call" or len(node.args) != 1 or node.args[0].kind != "range":
            continue
        r1, c1, r2, c2 = node.args[0].value
        size = r2 - r1 + 1
        if c1 != c2 or size == want or want < 2 or (r2 - first[0]) not in (0, -1):
            continue

        def make(t, pos, r1=r1, r2=r2):
            dr = pos[0] - first[0]
            lo, hi = r2 + dr - want + 1, r2 + dr
            if lo not in ctx.data:
                return None
            pat = re.compile(rf"(\$?){column_letters(c1)}(\$?){r1 + dr}:(\$?){column_letters(c1)}(\$?){r2 + dr}(?![0-9])")
            if len(pat.findall(t)) != 1:
                return None
            return pat.sub(lambda q: f"{q.group(1)}{column_letters(c1)}{q.group(2)}{lo}:{q.group(3)}"
                           f"{column_letters(c1)}{q.group(4)}{hi}", t)

        changes = _changes(ctx, run, make)
        bad = changes and _wrong(ctx, changes)
        if not bad:
            continue
        pos, now, right = bad
        msg = T(f"{ctx.label(run)} nennt {want}, das Fenster nimmt aber {size} Zeilen "
                f"({address((r1, c1))}:{address((r2, c1))} in {address(first)}). Richtig wäre in {address(pos)} "
                f"{_fmt(round(right, 4))} statt {_fmt(round(now, 4))}.",
                f"{ctx.label(run)} says {want}, but the window takes {size} rows "
                f"({address((r1, c1))}:{address((r2, c1))} in {address(first)}). In {address(pos)}, it should be "
                f"{_fmt(round(right, 4))} instead of {_fmt(round(now, 4))}.")
        found.append(_finding("span", ctx, run, msg, "fehler", changes))
    return found[:MAX_EACH]


# --- 13: Marge durch die Kosten (das wäre ein Aufschlag) ---

_MARGIN = _rx(r"marge|margin|rendite|gewinnquote|umsatzrendite")
_MARKUP = _rx(r"aufschlag|markup|zuschlag|kalkulation|vom einkauf|auf (?:die )?kosten|von (?:den )?kosten")
_SALES = ("umsatz", "erlös", "erlöse", "ertrag", "erträge", "verkauf", "verkaufspreis", "vk", "revenue", "sales",
          "einnahmen", "preis")


def _markup(ctx: _Ctx) -> list[Finding]:
    found = []
    for run in ctx.runs:
        first = run[0]
        name = ctx.name(first)
        shown_pct = str(ctx.sheet.cells[first].value).strip().endswith("%")
        if first[0] not in ctx.data or not _search(_MARGIN, name) or _search(_MARKUP, name) or \
                not (_search(_PCT, name) or shown_pct):
            continue
        if _search(_rx(r"\b(?:vom|von|des|of|am)\b"), name):
            continue  # Die Basis steht im Namen: Das prüft quoten._named_base.
        inner, _ = _times(ctx.sheet.nodes[first], 100.0)
        if inner.kind != "op" or inner.value != "/" or inner.args[1].kind != "ref" or inner.args[1].value[0] != first[0]:
            continue
        bottom = inner.args[1].value
        if not _has(ctx.name(bottom), _COSTS):
            continue
        sales = [c for c, h in ctx.plan.headers.items() if c != first[1] and _has(h, _SALES) and not _has(h, _COSTS)]
        if len(sales) != 1:
            continue

        def make(t, pos, bottom=bottom, col=sales[0]):
            return _swap_ref(t, (pos[0], bottom[1]), f"{{d1}}{column_letters(col)}{{d2}}{pos[0]}")

        changes = _changes(ctx, run, make)
        bad = changes and _wrong(ctx, changes)
        if not bad:
            continue
        pos, now, right = bad
        msg = T(f"{ctx.label(run)} teilt durch {column_letters(bottom[1])} „{ctx.name(bottom)}“. Eine Marge rechnet man "
                f"auf den Umsatz ({column_letters(sales[0])} „{ctx.head(sales[0])}“); durch die Kosten geteilt ist es "
                f"ein Aufschlag. Auf den Umsatz wäre es in {address(pos)} {_shown(ctx, pos, right)} statt "
                f"{_shown(ctx, pos, now)}. Ist ein Aufschlag gemeint, schreib „Aufschlag“ in den Namen.",
                f"{ctx.label(run)} divides by {column_letters(bottom[1])} “{ctx.name(bottom)}”. A margin is calculated "
                f"on revenue ({column_letters(sales[0])} “{ctx.head(sales[0])}”); divided by costs, it is a markup. "
                f"On revenue, {address(pos)} would be {_shown(ctx, pos, right)} instead of "
                f"{_shown(ctx, pos, now)}. If a markup is meant, write “markup” in the name.")
        found.append(_finding("markup", ctx, run, msg, "pruefen", changes))
    return found[:MAX_EACH]


# --- 24: „Bester“ mit MIN, „Schwächster“ mit MAX; 25: Kennzahl über die falsche Spalte ---

_HIGH = _rx(r"\b(?:beste[rsn]?|bester|top|max(?:imum|imal)?|most|highest|largest)\b|"
            r"(?:höchst|größt|grösst|stärkst|meist(?!ens)|umsatzstärkst)\w*")
_LOW = _rx(r"\b(?:min(?:imum|imal)?|lowest|smallest|least|worst)\b|"
           r"(?:schwächst|schlechtest|niedrigst|tiefst|kleinst|wenigst)\w*")
_BAD_IS_HIGH = ("kosten", "fehler", "ausfall", "dauer", "wartezeit", "verlust", "ausgaben", "aufwand", "beschwerden",
                "rang", "platz", "costs", "errors")


def _extreme(ctx: _Ctx) -> list[Finding]:
    found = []
    for run in ctx.runs:
        first = run[0]
        text = ctx.sheet.cells[first].formula or ""
        has_min, has_max = bool(re.search(r"\bMIN\(", text)), bool(re.search(r"\bMAX\(", text))
        if has_min == has_max or len(run) != 1:
            continue
        name = ctx.name(first)
        high, low = bool(_search(_HIGH, name)), bool(_search(_LOW, name))
        if high == low:
            continue
        cols = {a.value[1] for c in formel.calls(ctx.sheet.nodes[first]) if c.value in ("MIN", "MAX")
                for a in c.args if a.kind == "range"}
        if any(_has(ctx.head(c), _BAD_IS_HIGH) for c in cols) or _has(name, _BAD_IS_HIGH):
            continue
        if (high and has_max) or (low and has_min):
            continue
        new = re.sub(r"\bMIN\(", "MAX(", text) if has_min else re.sub(r"\bMAX\(", "MIN(", text)
        right, now = ctx.run(new), ctx.value(first)
        if right is None or right == now:
            continue
        word = _search(_HIGH if high else _LOW, name).group(0)
        msg = T(f"{ctx.label(run)} sucht mit {'MIN' if has_min else 'MAX'}, der Name sagt aber „{word}“. Richtig wäre "
                f"{_fmt(right)} statt {_fmt(now)}.",
                f"{ctx.label(run)} uses {'MIN' if has_min else 'MAX'}, but the name says “{word}”. It should be "
                f"{_fmt(right)} instead of {_fmt(now)}.")
        found.append(_finding("extreme", ctx, run, msg, "fehler", {first: new}))
    return found[:MAX_EACH]


_GENERIC = {_stem(w) for w in ("total", "summe", "gesamt", "gesamtsumme", "mittelwert", "durchschnitt", "anzahl",
                                "wert", "werte", "betrag", "zahl", "sum", "average", "mean", "count", "value")}
_STATS = {"STDEV", "STDEV.S", "STDEV.P", "STDEVA", "VAR", "VAR.S", "VAR.P", "MEDIAN", "MIN", "MAX", "AVERAGE", "SUM",
          "COUNT", "MODE", "MODE.SNGL", "GEOMEAN", "SKEW", "KURT", "QUARTILE", "PERCENTILE", "LARGE", "SMALL"}


def _stat_column(ctx: _Ctx) -> list[Finding]:
    found = []
    headers = {c: _key(h) for c, h in ctx.plan.headers.items()}
    for run in ctx.runs:
        first = run[0]
        if len(run) != 1 or (first[0] in ctx.data and first[1] in ctx.plan.headers):
            continue
        node = ctx.sheet.nodes[first]
        calls = [c for c in formel.calls(node) if c.value in _STATS]
        ranges = [a.value for c in calls for a in c.args if a.kind == "range" and a.value[1] == a.value[3]]
        if len(ranges) != 1 or ranges[0][1] not in headers:
            continue
        label = ctx.name(first)
        col = ranges[0][1]
        if col == first[1]:
            continue  # Die Summenzeile unter der eigenen Spalte: Ihr Name ist der der Zeile („Total“).
        named = [c for c, words in headers.items() if _covers(words, label) and words - _GENERIC]
        # Die längste Übereinstimmung zählt: „Umsatz-Index“ vor „Umsatz“, wenn beide Wörter dastehen.
        if named:
            most = max(len(headers[c]) for c in named)
            named = [c for c in named if len(headers[c]) == most]
        if len(named) != 1 or named[0] == col or _covers(headers[col], label):
            continue
        other = named[0]
        r1, _, r2, _ = ranges[0]

        def make(t, pos, col=col, other=other):
            pat = re.compile(rf"(\$?){column_letters(col)}(\$?){r1}:(\$?){column_letters(col)}(\$?){r2}(?![0-9])")
            if len(pat.findall(t)) != 1:
                return None
            return pat.sub(lambda q: f"{q.group(1)}{column_letters(other)}{q.group(2)}{r1}:{q.group(3)}"
                           f"{column_letters(other)}{q.group(4)}{r2}", t)

        changes = _changes(ctx, run, make)
        if not changes:
            continue
        right, now = ctx.run(changes[first]), ctx.value(first)
        if right is None or right == now:
            continue
        msg = T(f"{ctx.label(run)} rechnet über {column_letters(col)} „{ctx.head(col)}“, der Name meint aber "
                f"{column_letters(other)} „{ctx.head(other)}“. Richtig wäre {_fmt(right)} statt {_fmt(now)}.",
                f"{ctx.label(run)} calculates over {column_letters(col)} “{ctx.head(col)}”, but the name means "
                f"{column_letters(other)} “{ctx.head(other)}”. It should be {_fmt(right)} instead of {_fmt(now)}.")
        found.append(_finding("statcol", ctx, run, msg, "fehler", changes))
    return found[:MAX_EACH]


# --- 29: jede Zeile nimmt die Zeile darüber; 30: zwei Spalten vertauscht ---

_CHANGE = _rx(r"vor(?:monat|jahr|periode|woche|tag|quartal)|letzte|previous|prior|last|veränderung|änderung|"
                     r"wachstum|zuwachs|delta|differenz|diff|kumul|laufend|running|saldo|bestand|übertrag|vortrag|"
                     r"growth|change|vs|ggü|gegenüber|anfang|start|eröffnung|opening")


def _shifted(ctx: _Ctx) -> list[Finding]:
    found = []
    for run in ctx.runs:
        first = run[0]
        if len(run) < 3 or first[0] not in ctx.data or _search(_CHANGE, ctx.name(first)):
            continue
        refs = [n.value for n in _leaves(ctx.sheet.nodes[first])]
        if len(refs) < 2 or any(r[1] == first[1] for r in refs):
            continue
        offsets = {r[0] - first[0] for r in refs}
        if offsets != {0, -1} or sum(r[0] == first[0] - 1 for r in refs) != 1:
            continue
        late = next(r for r in refs if r[0] == first[0] - 1)
        if _search(_CHANGE, ctx.name((first[0], late[1]))):
            continue

        def make(t, pos, late=late):
            dr = pos[0] - first[0]
            return _swap_ref(t, (late[0] + dr, late[1]), f"{{d1}}{column_letters(late[1])}{{d2}}{pos[0]}")

        changes = _changes(ctx, run, make)
        bad = changes and _wrong(ctx, changes)
        if not bad:
            continue
        pos, now, right = bad
        header_hit = late[0] not in ctx.data
        msg = T(f"{ctx.label(run)} nimmt {column_letters(late[1])} „{ctx.head(late[1])}“ aus der Zeile darüber "
                f"({address(late)} in {address(first)}), die übrigen Bezüge aber aus der eigenen Zeile."
                + (f" In {address(first)} ist das sogar der Spaltenkopf." if header_hit else
                   " Ist das gewollt (Vormonat), schreib es in den Namen.")
                + f" Aus der eigenen Zeile wäre {address(pos)} {_fmt(round(right, 4))} statt {_fmt(round(now, 4))}.",
                f"{ctx.label(run)} takes {column_letters(late[1])} “{ctx.head(late[1])}” from the row above "
                f"({address(late)} in {address(first)}), but the other references from its own row."
                + (f" In {address(first)}, that is even the column header." if header_hit else
                   " If that is intended (previous month), say so in the name.")
                + f" From its own row, {address(pos)} would be {_fmt(round(right, 4))} instead of "
                f"{_fmt(round(now, 4))}.")
        found.append(_finding("shifted", ctx, run, msg, "fehler" if header_hit else "pruefen", changes))
    return found[:MAX_EACH]


def _leaves(node):
    if node.kind == "ref":
        yield node
    for a in node.args or ():
        if isinstance(a, formel.Node):
            yield from _leaves(a)


_REVENUE = ("umsatz", "erlös", "erlöse", "ertrag", "erträge", "einnahmen", "revenue", "sales", "verkauf")


def _swapped(ctx: _Ctx) -> list[Finding]:
    plan = ctx.plan
    rev = [c for c, h in plan.headers.items() if _has(h, _REVENUE) and not _key(h) - {_stem(w) for w in _REVENUE}]
    cost = [c for c, h in plan.headers.items() if _has(h, _COSTS) and not _key(h) - {_stem(w) for w in _COSTS}]
    if len(rev) != 1 or len(cost) != 1:
        return []
    r, k = rev[0], cost[0]
    from .massstab import _unit

    ur, uk = _unit(plan.headers[r]), _unit(plan.headers[k])
    if ur and uk and ur[:2] != uk[:2]:
        return []  # T€ gegen €: Die Zahlen lassen sich so nicht vergleichen (das sagt massstab.py).
    pairs = [(ctx.value((row, r)), ctx.value((row, k))) for row in ctx.rows]
    pairs = [(a, b) for a, b in pairs if isinstance(a, float) and isinstance(b, float)]
    if len(pairs) < 3 or not all(0 < a < b for a, b in pairs):
        return []
    hr, hk = (plan.header_row, r), (plan.header_row, k)
    cr, ck = ctx.sheet.cells.get(hr), ctx.sheet.cells.get(hk)
    fix = None
    if cr is not None and ck is not None and not cr.formula and not ck.formula:
        fix = {"range": f"{address(hr)}{T(' und ', ' and ')}{address(hk)}", "cells": {address(hr): ck.value, address(hk): cr.value},
               "was": {address(hr): {"text": cr.value}, address(hk): {"text": ck.value}},
               "german": T(f"„{ck.value}“ und „{cr.value}“ tauschen", f"Swap “{ck.value}” and “{cr.value}”"),
               "text": True}
    msg = T(f"In allen {len(pairs)} Zeilen ist {column_letters(k)} „{plan.headers[k]}“ größer als "
            f"{column_letters(r)} „{plan.headers[r]}“: jedes Ergebnis wäre ein Verlust. Sind die Spalten vertauscht?",
            f"In all {len(pairs)} rows, {column_letters(k)} “{plan.headers[k]}” is greater than "
            f"{column_letters(r)} “{plan.headers[r]}”: every result would be a loss. Are the columns swapped?")
    tip = T("Sind es vertauschte Köpfe, tauscht Einsetzen sie. Stimmt es so, setz „Passt so“.",
            "If the headers are swapped, Apply swaps them back. If it is right as it is, choose “Looks right”.")
    return [Finding(f"swapped:{address(hr)}", msg, "pruefen", tip, fix=fix, cell=address(hr))]


# --- 20: „bis heute“ oder „YTD“ mit einem festen Datum ---

_TODAY = _rx(r"bis heute|ytd|year.to.date|seit jahresbeginn|bisher|aktuell|laufendes jahr")


def _ytd(ctx: _Ctx) -> list[Finding]:
    found = []
    for run in ctx.runs:
        first = run[0]
        text = ctx.sheet.cells[first].formula or ""
        if len(run) != 1 or not _search(_TODAY, ctx.name(first)) or "TODAY()" in text.upper() or "NOW()" in text.upper():
            continue
        if not re.search(r"DATE\(\d{4},|\"[<>]=?\d", text):
            continue
        msg = T(f"{ctx.label(run)} heißt „bis heute“, rechnet aber mit einem festen Datum in der Formel. Morgen stimmt "
                "es nicht mehr.",
                f"{ctx.label(run)} means “to date”, but calculates with a fixed date in the formula. Tomorrow it will "
                "no longer be right.")
        tip = T("Nimm HEUTE() statt des festen Datums, etwa \"<=\"&HEUTE().",
                "Use TODAY() instead of the fixed date, e.g. \"<=\"&TODAY().")
        found.append(Finding(f"ytd:{address(first)}", msg, "pruefen", tip))
    return found[:MAX_EACH]



# --- z-Score, der durch den Mittelwert statt durch die Streuung teilt ---

_ZSCORE = _rx(r"z score|z wert|zscore|standard score|standardscore|standardisiert|z statistik|teststatistik")
_SPREAD = _rx(r"standardabweichung|std abw|stdabw|standard deviation|std dev|stdev|streuung|sigma|σ|\bsd\b|\bs\.d\.")
_EXPECTED = _rx(r"erwart|expected|referenz|poisson|soll|sollwert")
_RATIO = re.compile(r"^(=?)\s*\(\s*(.+?)\s*[-−]\s*(\$?[A-Z]{1,3}\$?\d+)\s*\)\s*/\s*(\$?[A-Z]{1,3}\$?\d+)\s*$")


def _bare(ref: str) -> str:
    return ref.replace("$", "").upper()


def _spread_cell(ctx: _Ctx, near: Pos) -> Pos | None:
    """Eine Standardabweichung auf dem Blatt, nach ihrem Namen: die nächste zu near."""
    best = None
    for pos, cell in ctx.sheet.cells.items():
        value = ctx.value(pos) if cell.formula else ctx.sheet.inputs.get(pos)
        if not isinstance(value, float) or isinstance(value, bool) or value <= 0:
            continue
        if not _search(_SPREAD, _left_label(ctx.sheet.cells, *pos) or ""):
            continue
        far = abs(pos[0] - near[0]) + 3 * abs(pos[1] - near[1])
        if best is None or far < best[0]:
            best = (far, pos)
    return best[1] if best else None


def _zscore(ctx: _Ctx) -> list[Finding]:
    """„z-Score“ =(B61-B60)/B60 teilt die Abweichung durch den Erwartungswert selbst, das ist eine relative
    Abweichung. Ein z-Score teilt durch die Streuung: die Standardabweichung, bei erwarteten Fällen (Poisson) die
    Wurzel aus dem Erwartungswert. Fehler mit Korrektur, wenn die Mappe sagt, welche: eine Zelle „Standardabweichung“
    oder ein erwarteter Wert; sonst „prüfen“."""
    found = []
    for run in ctx.runs:
        first = run[0]
        name = ctx.name(first)
        if not name or not (_search(_ZSCORE, name) or _plain(name) == "z"):
            continue
        text = ctx.sheet.cells[first].formula or ""
        m = _RATIO.match(text)
        if not m or _bare(m.group(3)) != _bare(m.group(4)):
            continue
        eq, x, mean = m.group(1), m.group(2), m.group(3)
        mean_pos = formel.position(_bare(mean))
        mean_name = _left_label(ctx.sheet.cells, *mean_pos) or ctx.name(mean_pos) or ""
        mean_value = ctx.value(mean_pos)
        now = ctx.value(first)
        if not isinstance(now, float) or not isinstance(mean_value, float):
            continue
        spread = _spread_cell(ctx, first)
        said = f"{address(mean_pos)}" + (T(f" („{mean_name}“)", f" (“{mean_name}”)") if mean_name else "")
        if spread is not None:
            sd = address(spread)
            new = f"{eq}({x}-{mean})/{sd}"
            why = T(f"Ein z-Score teilt die Abweichung durch die Streuung, hier die Standardabweichung in {sd} "
                    f"(„{_left_label(ctx.sheet.cells, *spread)}“).",
                    f"A z-score divides the deviation by the spread, here the standard deviation in {sd} "
                    f"(“{_left_label(ctx.sheet.cells, *spread)}”).")
        elif mean_value > 0 and (_search(_EXPECTED, mean_name) or any(
                _search(_EXPECTED, _left_label(ctx.sheet.cells, r, 1) or "") for r in range(first[0] - 6, first[0]))):
            new = f"{eq}({x}-{mean})/SQRT({mean})"
            why = T("Ein z-Score teilt die Abweichung durch die Streuung. Bei erwarteten Fällen (Poisson) ist das die "
                    "Wurzel aus dem Erwartungswert.",
                    "A z-score divides the deviation by the spread. For expected counts (Poisson), that is the "
                    "square root of the expected value.")
        else:
            msg = T(f"{address(first)} heißt „{name}“ und teilt die Abweichung durch {said} selbst. Das ist eine "
                    "relative Abweichung, kein z-Score: Der teilt durch die Standardabweichung.",
                    f"{address(first)} is called “{name}” and divides the deviation by {said} itself. That is a "
                    "relative deviation, not a z-score: a z-score divides by the standard deviation.")
            found.append(Finding(f"zscore:{address(first)}", msg, "pruefen",
                                 T("Teile durch die Standardabweichung (oder bei Fallzahlen durch die Wurzel aus dem "
                                   "Erwartungswert) oder nenn die Zelle „Abweichung in %“.",
                                   "Divide by the standard deviation (or, for case counts, by the square root of the "
                                   "expected value), or name the cell “deviation in %”.")))
            continue
        right = ctx.run(new)
        if not isinstance(right, float) or _close(right, now):
            continue
        msg = T(f"{address(first)} heißt „{name}“, teilt die Abweichung aber durch {said} selbst. {why} "
                f"Richtig wäre {_fmt(round(right, 4))} statt {_fmt(round(now, 4))}.",
                f"{address(first)} is called “{name}”, but divides the deviation by {said} itself. {why} "
                f"It should be {_fmt(round(right, 4))} instead of {_fmt(round(now, 4))}.")
        found.append(_finding("zscore", ctx, [first], msg, "fehler", {first: new}))
    return found[:MAX_EACH]


CHECKS = (_sign_order, _plus_minus, _gross, _percent, _time_unit, _period_unit, _index_base, _window_length, _markup,
          _extreme, _stat_column, _shifted, _swapped, _ytd, _zscore)


def checks(sheet: Sheet, rows: list[int]) -> list[Finding]:
    if sheet.plan.header_row is None:
        return []
    ctx = _Ctx(sheet, rows)
    found: list[Finding] = []
    seen: set[str] = set()
    for check in CHECKS:
        try:
            more = check(ctx)
        except Exception:  # noqa: BLE001 - eine Prüfung darf die anderen nicht mitreißen
            log.exception("Prüfung %s", check.__name__)
            continue
        for item in more:
            where = item.key.partition(":")[2]
            if where not in seen:
                seen.add(where)
                found.append(item)
    return found
