"""Fehler aus Controlling und Buchhaltung, wie sie das Finance-Plug-in beschreibt (Abweichungsanalyse, Abstimmung,
Abschluss, Anlagen, Investition): Preis- und Mengeneffekt, die nicht aufgehen, eine Brücke ohne Treiber, Veränderung
gegen eine negative Basis, Vorzeichen gegen „abzüglich“, Soll gegen Haben, AfA ohne Restwert, Buchwert unter null,
Jahreszins mit Monaten, MAPE ohne Betrag, NBW über Jahr 0, Anteil an einer Zeile ohne $.

Jede Prüfung liest nur, was die Formeln und ihre Namen sagen, und schlägt eine Korrektur nur vor, wenn sie auf dem
Blatt nachgerechnet etwas anderes ergibt."""
from __future__ import annotations

import logging
import re

from . import einsetzen, formel
from .analyse import Finding, Sheet, _fmt, address
from .kopf import _Ctx, _close, _plain, _rx, _search, _swap_ref
from .sprache import T
from .workbook import _left_label
from .xlsx import column_letters

log = logging.getLogger(__name__)
Pos = tuple[int, int]
MAX_EACH = 4


def _walk(node):
    yield node
    for arg in node.args or ():
        if isinstance(arg, formel.Node):
            yield from _walk(arg)


def _formulas(ctx: _Ctx):
    for pos, node in sorted(ctx.sheet.nodes.items()):
        if node is None or pos in ctx.sheet.unknown:
            continue
        cell = ctx.sheet.cells.get(pos)
        if cell is not None and cell.formula:
            yield pos, node, cell.formula


def _number(v) -> bool:
    return isinstance(v, float) and not isinstance(v, bool)


def _col_head(ctx: _Ctx, col: int, below: int) -> str:
    """Der Kopf einer Spalte: aus der Tabelle, sonst der nächste Text darüber."""
    if col in ctx.plan.headers and ctx.plan.header_row is not None and below > ctx.plan.header_row:
        return ctx.plan.headers[col]
    for r in range(below - 1, 0, -1):
        cell = ctx.sheet.cells.get((r, col))
        if cell is not None and not cell.formula and isinstance(cell.value, str) and re.search(r"[^\W\d_]", cell.value):
            return cell.value.strip()
    return ""


def _row_name(ctx: _Ctx, pos: Pos) -> str:
    """Der Name einer Zeile: der Text links daneben. Eine Einheitenspalte („TEUR“, „%“) dazwischen zählt nicht."""
    for c in range(pos[1] - 1, 0, -1):
        cell = ctx.sheet.cells.get((pos[0], c))
        if cell is not None and not cell.formula and isinstance(cell.value, str) and re.search(r"[^\W\d_]", cell.value) \
                and not _unit_cell(ctx, (pos[0], c)):
            return cell.value.strip()
    return ""


# Einheiten, wie sie Finanzmodelle in eine eigene Spalte zwischen Zeilenname und Zahlen schreiben.
_UNIT_HEAD = _rx(r"^(einheit|einh|unit|units|dimension|wahrung|currency|masseinheit|unite|unita|valuta)$")
_UNIT_VALUE = _rx(r"^(t|k|tsd\.?|mio\.?|mrd\.?|bn|m)?\s*(eur|chf|usd|gbp|€|\$)$|^%|^(x|faktor|jahre?|monate?|tage?|"
                  r"stuck|stk\.?|anzahl|mio\.? stuck|years?|months?|days?|pieces|pcs|fte|vza|ma)$")


def _row_unit(ctx: _Ctx, pos: Pos) -> str:
    """Die Einheit der Zeile aus ihrer Einheitenspalte („Mio. EUR“, „%“), sonst ""."""
    for c in range(pos[1] - 1, 0, -1):
        cell = ctx.sheet.cells.get((pos[0], c))
        if cell is None or cell.formula or not isinstance(cell.value, str) \
                or isinstance(ctx.sheet.inputs.get((pos[0], c)), (int, float)):
            continue
        if _unit_cell(ctx, (pos[0], c)):
            return cell.value.strip()
        return ""
    return ""


def _unit_cell(ctx: _Ctx, pos: Pos) -> bool:
    cell = ctx.sheet.cells.get(pos)
    text = str(cell.value).strip() if cell is not None else ""
    if len(text) > 24:
        return False
    if _search(_UNIT_VALUE, text):
        return True
    first = _unit_heads(ctx).get(pos[1])
    return first is not None and first < pos[0]


def _unit_heads(ctx: _Ctx) -> dict[int, int]:
    """Je Spalte die oberste Zeile mit einem Einheitenkopf („Einheit“). Einmal je Blatt, nicht je Zelle: große
    Mappen fragen das für tausende Zellen."""
    heads = getattr(ctx, "_unit_heads", None)
    if heads is None:
        heads = {}
        for (r, c), cell in ctx.sheet.cells.items():
            if not cell.formula and isinstance(cell.value, str) and len(cell.value) < 40 and (c not in heads or r < heads[c]) \
                    and _search(_UNIT_HEAD, cell.value.strip()):
                heads[c] = r
        ctx._unit_heads = heads
    return heads


def _name(ctx: _Ctx, pos: Pos) -> str:
    """Wie eine Zelle heißt: der Text links daneben, sonst der Spaltenkopf."""
    label = _left_label(ctx.sheet.cells, *pos)
    if label and _unit_cell(ctx, (pos[0], pos[1] - 1)):
        label = _row_name(ctx, pos)
    return label or _col_head(ctx, pos[1], pos[0])


def _fix(ctx: _Ctx, changes: dict[Pos, str]) -> dict | None:
    return ctx.fix(changes)


def _tip(ctx: _Ctx, changes: dict[Pos, str]) -> str:
    cells = sorted(changes, key=lambda p: (p[1], p[0]))
    first = cells[0]
    formula = f"={ctx.german(changes[first])}"
    tip = T(f"Schreib in {address(first)} {formula}", f"In {address(first)}, enter {formula}")
    if len(cells) > 1:
        last, rest = address(cells[-1]), ", ".join(address(p) for p in cells[1:4])
        tip += T(f" und zieh es herunter bis {last}", f" and fill it down to {last}") \
            if all(p[1] == first[1] for p in cells) else \
            T(f" und zieh es nach rechts bis {last}", f" and fill it right to {last}") \
            if all(p[0] == first[0] for p in cells) else \
            T(f" (und so weiter in {rest})", f" (and so on in {rest})")
    return tip + "."


def _shown(ctx: _Ctx, pos: Pos, value) -> str:
    cell = ctx.sheet.cells.get(pos)
    if _number(value) and cell is not None and str(cell.value).strip().endswith("%"):
        return f"{_fmt(round(value * 100, 2))} %"
    return _fmt(round(value, 2)) if _number(value) else _fmt(value)


def _finding(kind: str, ctx: _Ctx, pos: Pos, text: str, level: str, changes: dict[Pos, str] | None) -> Finding | None:
    fix = _fix(ctx, changes) if changes else None
    if changes and fix is None:
        return None
    return Finding(f"{kind}:{address(pos)}", text, level, _tip(ctx, changes) if changes else "", fix=fix)


def _differs(ctx: _Ctx, changes: dict[Pos, str]) -> tuple[Pos, float, float] | None:
    for pos, text in sorted(changes.items()):
        now, right = ctx.value(pos), ctx.run(text)
        if _number(now) and _number(right) and not _close(now, right):
            return pos, now, right
    return None


def _column(ctx: _Ctx, col: int, top: int, bottom: int) -> list[Pos]:
    return [(r, col) for r in range(top, bottom + 1) if (r, col) in ctx.sheet.nodes and ctx.sheet.nodes[(r, col)] is not None]


def _refs(node) -> list[Pos]:
    return [n.value for n in _walk(node) if n.kind == "ref"]


# --- Preis- und Mengeneffekt gehen nicht auf ---

_VOLUME = _rx(r"mengeneffekt|volumeneffekt|mengenabweichung|volume effect|volume variance|effet volume|effetto volume")
_PRICE = _rx(r"preiseffekt|preisabweichung|price effect|price variance|rate effect|effet prix|effetto prezzo")
_MIX = _rx(r"mixeffekt|mix effect|mix variance|mixabweichung")
_VARIANCE = _rx(r"abweichung|differenz|delta|variance|varianz|veranderung|ecart|scostamento")
_PLAN = _rx(r"\bplan|budget|\bsoll\b|forecast|vorjahr|prior|\bvj\b|bdg|\bbud\b")
_ACTUAL = _rx(r"\bist\b|actual|\bact\b|aktuell|effektiv|real|reel|consuntivo")
_PRICEWORD = _rx(r"preis|price|\basp\b|tarif|prix|prezzo|(?<!um)satz\b|\brate\b")
_QTYWORD = _rx(r"menge|volumen|stuck|anzahl|volume|units|\bqty\b|quantit|absatz|falle|cases")


def _effects(ctx: _Ctx) -> list[Finding]:
    heads = ctx.plan.headers
    if ctx.plan.header_row is None or not ctx.rows:
        return []
    volume = [c for c, h in heads.items() if _search(_VOLUME, h)]
    price = [c for c, h in heads.items() if _search(_PRICE, h)]
    mix = [c for c, h in heads.items() if _search(_MIX, h)]
    if len(volume) != 1 or len(price) != 1 or len(mix) > 1:
        return []
    variance = [c for c, h in heads.items() if _search(_VARIANCE, h) and c not in volume + price + mix
                and not str(ctx.sheet.cells.get((ctx.rows[0], c)).value if (ctx.rows[0], c) in ctx.sheet.cells else "").endswith("%")]
    if len(variance) != 1:
        return []
    v, p, d = volume[0], price[0], variance[0]
    rows = [r for r in ctx.rows if all(_number(ctx.value((r, c))) for c in (v, p, d))]
    if len(rows) < 2:
        return []

    def gap(r: int, values: dict[int, float] | None = None) -> float:
        values = values or {}
        total = sum(values.get(c, ctx.value((r, c))) for c in [v, p] + mix)
        return total - ctx.value((r, d))

    off = [r for r in rows if abs(gap(r)) > 1e-6 * max(1.0, abs(ctx.value((r, d))))]
    if not off:
        return []
    # Welche Spalte ist Plan, welche Ist: je Preis und Menge ein Paar.
    pairs = {}
    for word, key in ((_PRICEWORD, "preis"), (_QTYWORD, "menge")):
        plan = [c for c, h in heads.items() if _search(word, h) and _search(_PLAN, h) and not _search(_VARIANCE, h)]
        act = [c for c, h in heads.items() if _search(word, h) and _search(_ACTUAL, h) and not _search(_VARIANCE, h)]
        if len(plan) == 1 and len(act) == 1:
            pairs[key] = (plan[0], act[0])
    # Die üblichen Formeln: Mengeneffekt mit dem Plan-Preis, Preiseffekt mit der Ist-Menge. Steht in einer der beiden
    # Spalten der andere Wert, ist das die Ursache.
    for col, key, want_plan, what in ((v, "preis", True, T("Plan-Preis", "budget price")),
                                      (p, "menge", False, T("Ist-Menge", "actual quantity"))):
        if key not in pairs:
            continue
        plan_c, act_c = pairs[key]
        right_c, wrong_c = (plan_c, act_c) if want_plan else (act_c, plan_c)
        cells = _column(ctx, col, rows[0], rows[-1])
        if len(cells) < len(rows) or not all((r, wrong_c) in _refs(ctx.sheet.nodes[(r, col)]) and
                                             (r, right_c) not in _refs(ctx.sheet.nodes[(r, col)]) for r, _ in cells):
            continue
        changes = {}
        for pos in cells:
            new = _swap_ref(ctx.sheet.cells[pos].formula, (pos[0], wrong_c), f"{{d1}}{column_letters(right_c)}{{d2}}{pos[0]}")
            if new is None:
                break
            changes[pos] = new
        else:
            fixed = {r: ctx.run(changes[(r, col)]) for r in rows if (r, col) in changes}
            if not all(_number(x) for x in fixed.values()):
                continue
            if any(abs(gap(r, {col: fixed[r]})) > 1e-6 * max(1.0, abs(ctx.value((r, d)))) for r in rows):
                continue
            first = (off[0], col)
            other = T("Preiseffekt", "the price effect") if col == v else T("Mengeneffekt", "the volume effect")
            usual = T("Mengeneffekt mit dem Plan-Preis, der Preiseffekt mit der Ist-Menge", "volume effect at the budget "
                      "price and the price effect at the actual quantity") if col == v else \
                T("Preiseffekt mit der Ist-Menge, der Mengeneffekt mit dem Plan-Preis", "price effect at the actual "
                  "quantity and the volume effect at the budget price")
            span, partner, short = (f"{address(cells[0])}:{address(cells[-1])}", heads[p if col == v else v],
                                    _fmt(round(-gap(off[0]), 2)))
            right, now = _shown(ctx, first, fixed[off[0]]), _shown(ctx, first, ctx.value(first))
            text = T(f"{span} ({heads[col]}) und {partner} ergeben "
                     f"zusammen nicht die {heads[d]}: In Zeile {off[0]} fehlen {short}. "
                     f"{heads[col]} rechnet mit {heads[wrong_c]} statt mit {heads[right_c]}. Üblich ist der "
                     f"{usual}; "
                     f"dann gehen beide genau auf. Mit dem {what} wäre {address(first)} "
                     f"{right} statt {now}, und {other} "
                     f"bleibt, wie er ist.",
                     f"{span} ({heads[col]}) and {partner} together don't add up to {heads[d]}: row {off[0]} is short "
                     f"by {short}. {heads[col]} uses {heads[wrong_c]} instead of {heads[right_c]}. The usual way is the "
                     f"{usual}; then both add up exactly. With the {what}, {address(first)} would be {right} instead "
                     f"of {now}; {other} stays as it is.")
            item = _finding("effekte", ctx, cells[0], text, "fehler", changes)
            return [item] if item else []
    first = (off[0], v)
    much = _fmt(round(gap(off[0]), 2))
    text = T(f"{heads[v]} und {heads[p]}{' und ' + heads[mix[0]] if mix else ''} ergeben in {len(off)} von {len(rows)} "
             f"Zeilen nicht die {heads[d]} (Zeile {off[0]}: {much} zu viel). Üblich: Mengeneffekt "
             f"= (Ist-Menge − Plan-Menge) × Plan-Preis, Preiseffekt = (Ist-Preis − Plan-Preis) × Ist-Menge.",
             f"{heads[v]}{', ' + heads[p] + ' and ' + heads[mix[0]] if mix else ' and ' + heads[p]} don't add up to "
             f"{heads[d]} in {len(off)} of {len(rows)} rows (row {off[0]}: {much} too much). Usual: volume effect "
             f"= (actual quantity − budget quantity) × budget price, price effect = (actual price − budget price) × "
             f"actual quantity.")
    return [Finding(f"effekte:{address(first)}", text, "pruefen", "")]


# --- Brücke: ein Treiber fehlt in der Summe ---

def _bridge(ctx: _Ctx) -> list[Finding]:
    found = []
    for pos, node, text in _formulas(ctx):
        terms = _plus_terms(node)
        if not terms or len(terms) < 3 or any(t[1] != pos[1] for t in terms):
            continue
        rows = sorted(t[0] for t in terms)
        if rows != list(range(rows[0], rows[-1] + 1)) or rows[-1] >= pos[0] - 1:
            continue
        skipped = [(r, pos[1]) for r in range(rows[-1] + 1, pos[0])]
        if not all(_input_number(ctx, p) and _row_name(ctx, p) for p in skipped):
            continue
        if any(ctx.sheet.dependents.get(p) for p in skipped):
            continue  # Liest eine andere Formel die Zeile, ist sie nicht vergessen.
        new = text + "".join(f"+{address(p)}" for p in skipped)
        changes = {pos: new}
        bad = _differs(ctx, changes)
        if not bad:
            continue
        names = ", ".join(f"{address(p)} ({_row_name(ctx, p)})" for p in skipped)
        label, top, bottom = _row_name(ctx, pos) or _name(ctx, pos), address((rows[0], pos[1])), address((rows[-1], pos[1]))
        right, now = _shown(ctx, pos, bad[2]), _shown(ctx, pos, bad[1])
        msg = T(f"{address(pos)} ({label}) zählt {top} bis "
                f"{bottom} zusammen, aber nicht {names} direkt darüber. Keine andere Formel liest "
                f"{'diese Zeile' if len(skipped) == 1 else 'diese Zeilen'}: Sie fehlt in der Summe. Mit ihr wäre "
                f"{address(pos)} {right} statt {now}.",
                f"{address(pos)} ({label}) adds up {top} to {bottom}, but not {names} directly above. No other formula "
                f"reads {'this row' if len(skipped) == 1 else 'these rows'}: "
                f"{'it is' if len(skipped) == 1 else 'they are'} missing from the total. With "
                f"{'it' if len(skipped) == 1 else 'them'}, {address(pos)} would be {right} instead of {now}.")
        item = _finding("bruecke", ctx, pos, msg, "fehler", changes)
        if item:
            found.append(item)
    return found[:MAX_EACH]


def _plus_terms(node) -> list[Pos] | None:
    """A+B+C+…, nur Bezüge und Pluszeichen."""
    if node.kind == "ref":
        return [node.value]
    if node.kind == "op" and node.value == "+" and len(node.args) == 2:
        left, right = _plus_terms(node.args[0]), _plus_terms(node.args[1])
        if left is not None and right is not None:
            return left + right
    return None


def _input_number(ctx: _Ctx, pos: Pos) -> bool:
    cell = ctx.sheet.cells.get(pos)
    return cell is not None and not cell.formula and _number(ctx.value(pos))


# --- Veränderung gegen eine negative Basis ---

_SIGNED = _rx(r"veranderung|abweichung|wachstum|zuwachs|change|growth|delta|differenz|\bvs\b|ggu|gegenuber|"
              r"variation|crescita|evolution|steigerung")


def _negative_base(ctx: _Ctx) -> list[Finding]:
    found = []
    for run in ctx.runs:
        if len(run) < 3:
            continue
        bases = []
        for pos in run:
            base = _base_of(ctx.sheet.nodes[pos])
            if base is None:
                break
            bases.append(base)
        else:
            if len({(b[0] - p[0], b[1]) for b, p in zip(bases, run)}) != 1:
                continue
            values = [ctx.value(b) for b in bases]
            if not all(_number(x) for x in values) or not any(x < 0 for x in values) or not any(x > 0 for x in values):
                continue
            name = _col_head(ctx, run[0][1], run[0][0])
            if name and not _search(_SIGNED, name) and not str(ctx.sheet.cells[run[0]].value).strip().endswith("%"):
                continue
            changes = {}
            for pos, base in zip(run, bases):
                new = _abs_base(ctx.sheet.cells[pos].formula, ctx.sheet.nodes[pos], base)
                if new is None:
                    break
                changes[pos] = new
            else:
                bad = _differs(ctx, changes)
                if not bad:
                    continue
                pos, now, right = bad
                base = bases[run.index(pos)]
                span, label = f"{address(run[0])}:{address(run[-1])}", name or _name(ctx, run[0])
                by, low = _col_head(ctx, base[1], base[0]) or address(base), _fmt(ctx.value(base))
                high = _fmt(_other_value(ctx, ctx.sheet.nodes[pos], base))
                shows, should = _shown(ctx, pos, now), _shown(ctx, pos, right)
                msg = T(f"{span} ({label}) teilt durch "
                        f"{by}, und das ist in {address(base)} negativ "
                        f"({low}). Durch eine negative Zahl geteilt kehrt sich das Vorzeichen um: "
                        f"Von {low} auf {high} zeigt "
                        f"{address(pos)} {shows}, richtig ist {should}. Mit dem Betrag "
                        f"der Basis (ABS) stimmt das Vorzeichen in jeder Zeile.",
                        f"{span} ({label}) divides by {by}, which is negative in {address(base)} ({low}). Dividing by a "
                        f"negative number flips the sign: {address(pos)} shows {shows} for the change from {low} to "
                        f"{high}; the correct value is {should}. With the absolute value of the base (ABS), the sign is right in "
                        f"every row.")
                item = _finding("negbasis", ctx, pos, msg, "fehler", changes)
                if item:
                    found.append(item)
    return found[:MAX_EACH]


def _base_of(node) -> Pos | None:
    """(X−Y)/Y oder X/Y−1: der Bezug Y, durch den geteilt wird."""
    if node.kind == "op" and node.value == "/" and len(node.args) == 2 and node.args[1].kind == "ref":
        top = node.args[0]
        if top.kind == "op" and top.value == "-" and len(top.args) == 2 and top.args[1].kind == "ref" and \
                top.args[0].kind == "ref" and top.args[1].value == node.args[1].value:
            return node.args[1].value
    if node.kind == "op" and node.value == "-" and len(node.args) == 2 and node.args[1].kind == "num" and \
            node.args[1].value == 1.0:
        div = node.args[0]
        if div.kind == "op" and div.value == "/" and all(a.kind == "ref" for a in div.args):
            return div.args[1].value
    return None


def _other_value(ctx: _Ctx, node, base: Pos):
    for ref in _refs(node):
        if ref != base:
            return ctx.value(ref)
    return None


def _abs_base(formula: str, node, base: Pos) -> str | None:
    b = address(base)
    if node.kind == "op" and node.value == "-":
        other = next((r for r in _refs(node) if r != base), None)
        return f"({address(other)}-{b})/ABS({b})" if other else None
    # Den Nenner (den letzten Bezug auf die Basis) in ABS setzen.
    match = list(re.finditer(rf"(?<![A-Za-z0-9_.$!:])\$?{column_letters(base[1])}\$?{base[0]}(?![0-9A-Za-z(!:])", formula))
    if len(match) != 2:
        return None
    m = match[-1]
    if not formula[:m.start()].rstrip().endswith("/"):
        return None
    return formula[:m.start()] + f"ABS({m.group(0)})" + formula[m.end():]


# --- Vorzeichen gegen „abzüglich“ / „zuzüglich“ ---

_LESS = _rx(r"^(?:abzuglich|abzgl|abz\b|\./\.|minus\b|less\b|deduct|abzug|abgezogen|moins|meno)")
_MORE = _rx(r"^(?:zuzuglich|zzgl|zuz\b|plus\b|add\b|addiert|hinzu|plus|piu)")


def _signs(ctx: _Ctx) -> list[Finding]:
    found = []
    for pos, node, text in _formulas(ctx):
        terms = _signed_terms(node, 1)
        if not terms or len(terms) < 3 or len({t[1] for t, _ in terms}) != 1:
            continue
        wrong = []
        for ref, sign in terms:
            name = _plain(_row_name(ctx, ref), keep_spaces=True).strip()
            value = ctx.value(ref)
            if not name or not _number(value) or value <= 0:
                continue
            if (sign > 0 and _LESS.search(name)) or (sign < 0 and _MORE.search(name)):
                wrong.append((ref, sign))
        if not wrong or len(wrong) == len(terms):
            continue
        new = text
        for ref, sign in wrong:
            new = _flip(new, ref, sign)
            if new is None:
                break
        if new is None:
            continue
        changes = {pos: new}
        bad = _differs(ctx, changes)
        if not bad:
            continue
        label, right, now = _row_name(ctx, pos) or _name(ctx, pos), _shown(ctx, pos, bad[2]), _shown(ctx, pos, bad[1])
        msg = T(f"{address(pos)} ({label}) hat "
                + ", ".join(f"{address(r)} „{_row_name(ctx, r)}“ {'addiert' if s > 0 else 'abgezogen'}" for r, s in wrong)
                + f", die Zeile sagt aber das Gegenteil. Mit dem richtigen Vorzeichen wäre {address(pos)} {right} statt {now}.",
                f"{address(pos)} ({label}) "
                + ", ".join(f"{'adds' if s > 0 else 'subtracts'} {address(r)} “{_row_name(ctx, r)}”" for r, s in wrong)
                + f", but the row label says the opposite. With the correct sign, {address(pos)} would be {right} "
                f"instead of {now}.")
        item = _finding("vorzeichen", ctx, pos, msg, "fehler", changes)
        if item:
            found.append(item)
    return found[:MAX_EACH]


def _signed_terms(node, sign: int) -> list[tuple[Pos, int]] | None:
    if node.kind == "ref":
        return [(node.value, sign)]
    if node.kind == "op" and node.value in ("+", "-") and len(node.args) == 2:
        left = _signed_terms(node.args[0], sign)
        right = _signed_terms(node.args[1], sign if node.value == "+" else -sign)
        if left is not None and right is not None:
            return left + right
    return None


def _flip(text: str, ref: Pos, sign: int) -> str | None:
    rx = re.compile(rf"([+\-])(\s*\$?{column_letters(ref[1])}\$?{ref[0]})(?![0-9A-Za-z(!:])")
    matches = rx.findall(text)
    if len(matches) != 1 or (matches[0][0] == "+") != (sign > 0):
        return None
    return rx.sub(lambda m: ("-" if m.group(1) == "+" else "+") + m.group(2), text)


# --- Soll gegen Haben, Aktiva gegen Passiva ---

_SIDES = ((_rx(r"\bsoll\b"), _rx(r"\bhaben\b")), (_rx(r"\bdebit|\bdebet"), _rx(r"\bcredit|\bkredit")),
          (_rx(r"\baktiva\b|\baktivseite\b"), _rx(r"\bpassiva\b|\bpassivseite\b")),
          (_rx(r"\bassets\b"), _rx(r"\bliabilities\b")), (_rx(r"\bactif\b"), _rx(r"\bpassif\b")),
          (_rx(r"\battivo\b|\battivita\b"), _rx(r"\bpassivo\b|\bpassivita\b")),
          (_rx(r"mittelverwendung"), _rx(r"mittelherkunft")))
_TOTALISH = {"SUM", "SUBTOTAL"}


def _balance(ctx: _Ctx) -> list[Finding]:
    totals = []
    for pos, node, text in _formulas(ctx):
        if node.kind != "call" or node.value not in _TOTALISH or not _number(ctx.value(pos)):
            continue
        name = " ".join(x for x in (_row_name(ctx, pos), _col_head(ctx, pos[1], pos[0])) if x)
        if len(name.split()) > 8:
            continue
        totals.append((pos, name))
    found = []
    for left, right in _SIDES:
        a = [(p, n) for p, n in totals if _search(left, n) and not _search(right, n)]
        b = [(p, n) for p, n in totals if _search(right, n) and not _search(left, n)]
        if len(a) != 1 or len(b) != 1:
            continue
        (pa, na), (pb, nb) = a[0], b[0]
        va, vb = ctx.value(pa), ctx.value(pb)
        if abs(va - vb) <= 0.005 + 1e-9 * max(abs(va), abs(vb)):
            continue
        diff = vb - va
        side_a = _side(left, na)
        side_b = _side(right, nb)
        gap = _fmt(round(abs(diff), 2))
        msg = T(f"{side_a} ({address(pa)}) und {side_b} ({address(pb)}) müssen gleich sein, ergeben aber "
                f"{_fmt(va)} und {_fmt(vb)}: {gap} Differenz.",
                f"{side_a} ({address(pa)}) and {side_b} ({address(pb)}) must be equal, but come to {_fmt(va)} and "
                f"{_fmt(vb)}: a difference of {gap}.")
        swapped, two = _transposed(ctx, pa, pb, abs(diff))
        if swapped:
            msg += T(f" Die Differenz ist durch 9 teilbar, typisch für einen Zahlendreher: "
                     f"{swapped} {'hat' if not two else 'haben'} dieselben Ziffern in anderer Reihenfolge.",
                     f" The difference is divisible by 9, typical of transposed digits: {swapped} "
                     f"{'has' if not two else 'have'} the same digits in a different order.")
        found.append(Finding(f"ausgleich:{address(pb)}", msg, "pruefen", ""))
    return found[:MAX_EACH]


def _side(rx: re.Pattern, name: str) -> str:
    plain = _plain(name, keep_spaces=True)
    m = rx.search(plain)
    return name.split()[0] if m is None else next((w for w in name.split() if rx.search(_plain(w, True))), name)


def _parts(ctx: _Ctx, pos: Pos) -> list[Pos]:
    node = ctx.sheet.nodes.get(pos)
    out = []
    for n in _walk(node):
        if n.kind == "range":
            r1, c1, r2, c2 = n.value
            out += [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
        elif n.kind == "ref":
            out.append(n.value)
    return [p for p in out if _input_number(ctx, p)]


def _transposed(ctx: _Ctx, pa: Pos, pb: Pos, diff: float) -> tuple[str, bool]:
    """Zwei Zellen, je eine je Seite, mit denselben Ziffern vertauscht und genau der Differenz auseinander; dazu, ob
    es zwei Zellen sind."""
    if diff < 9 or abs(diff / 9 - round(diff / 9)) > 1e-9:
        return "", False
    sides = [_parts(ctx, pa), _parts(ctx, pb)]
    for x in sides[0]:
        for y in sides[1]:
            vx, vy = ctx.value(x), ctx.value(y)
            if _close(abs(vx - vy), diff) and _digits(vx) == _digits(vy) and vx != vy:
                return T(f"{address(x)} ({_fmt(vx)}) und {address(y)} ({_fmt(vy)})",
                         f"{address(x)} ({_fmt(vx)}) and {address(y)} ({_fmt(vy)})"), True
    for side in sides:
        for x in side:
            vx = ctx.value(x)
            for swap in _swaps(vx):
                if _close(abs(vx - swap), diff):
                    return T(f"{address(x)} ({_fmt(vx)}, vielleicht {_fmt(swap)})",
                             f"{address(x)} ({_fmt(vx)}, perhaps {_fmt(swap)})"), False
    return "", False


def _digits(v: float) -> str:
    return "".join(sorted(f"{abs(v):.2f}".replace(".", "").lstrip("0")))


def _swaps(v: float) -> list[float]:
    if not float(v).is_integer():
        return []
    s = str(int(abs(v)))
    out = []
    for i in range(len(s) - 1):
        if s[i] != s[i + 1]:
            t = s[:i] + s[i + 1] + s[i] + s[i + 2:]
            if not t.startswith("0"):
                out.append(float(t) * (1 if v >= 0 else -1))
    return out


# --- AfA ohne Restwert ---

_AFA = _rx(r"\bafa\b|abschreibung|depreciation|amortisation|amortization|ammortamento")
_SALVAGE = _rx(r"restwert|schrottwert|residualwert|residual|salvage|liquidationswert|valeur residuelle|valore residuo")
_COST = _rx(r"anschaffung|herstellkosten|kaufpreis|\bkosten\b|cost|wert\b|betrag|investition")
_LIFE = _rx(r"nutzungsdauer|laufzeit|\bnd\b|jahre|useful life|\blife\b|durata|duree")


def _salvage(ctx: _Ctx) -> list[Finding]:
    heads = ctx.plan.headers
    salvage = [c for c, h in heads.items() if _search(_SALVAGE, h)]
    if len(salvage) != 1 or not ctx.rows:
        return []
    s = salvage[0]
    found = []
    for run in ctx.runs:
        col = run[0][1]
        if run[0][0] not in ctx.data or not _search(_AFA, heads.get(col, "")) or _search(_SALVAGE, heads.get(col, "")):
            continue
        node = ctx.sheet.nodes[run[0]]
        r = run[0][0]
        if any(ref[1] == s for ref in _refs(node)):
            continue
        new_of = None
        if node.kind == "op" and node.value == "/" and node.args[0].kind == "ref" and node.args[1].kind == "ref":
            cost, life = node.args[0].value, node.args[1].value
            if cost[0] != r or not _search(_COST, heads.get(cost[1], "")) or not _search(_LIFE, heads.get(life[1], "")):
                continue
            def new_of(pos, cost=cost, life=life):
                dr = pos[0] - r
                return (f"({column_letters(cost[1])}{cost[0] + dr}-{column_letters(s)}{pos[0]})/"
                        f"{_ref_text(ctx.sheet.cells[pos].formula, (life[0] + dr, life[1]))}")
        elif node.kind == "call" and node.value == "SLN" and len(node.args) == 3 and node.args[1].kind == "num" \
                and node.args[1].value == 0:
            def new_of(pos):
                text = ctx.sheet.cells[pos].formula
                m = re.search(r"(SLN\([^,;]+[,;]\s*)0(\s*[,;])", text, re.IGNORECASE)
                return None if m is None else text[:m.start()] + m.group(1) + f"{column_letters(s)}{pos[0]}" + \
                    m.group(2) + text[m.end():]
        else:
            continue
        values = [ctx.value((p[0], s)) for p in run]
        if not all(_number(x) for x in values) or not any(x > 0 for x in values):
            continue
        changes = {}
        for pos in run:
            new = new_of(pos)
            if new is None:
                break
            changes[pos] = new
        else:
            bad = _differs(ctx, changes)
            if not bad:
                continue
            pos, now, right = bad
            span, there, rest = f"{address(run[0])}:{address(run[-1])}", address((pos[0], s)), _fmt(ctx.value((pos[0], s)))
            should, shows = _shown(ctx, pos, right), _shown(ctx, pos, now)
            msg = T(f"{span} ({heads[col]}) schreibt die ganzen Anschaffungskosten ab, "
                    f"obwohl daneben ein {heads[s]} steht ({there}: {rest}). "
                    f"Abgeschrieben wird nur, was über den {heads[s]} hinaus verbraucht wird: (Anschaffung − {heads[s]}) / "
                    f"Nutzungsdauer. So wäre {address(pos)} {should} statt {shows}.",
                    f"{span} ({heads[col]}) depreciates the full acquisition cost, although there is a {heads[s]} next "
                    f"to it ({there}: {rest}). Only what is used up beyond the {heads[s]} is depreciated: (acquisition "
                    f"cost − {heads[s]}) / useful life. Then {address(pos)} would be {should} instead of {shows}.")
            item = _finding("restwert", ctx, run[0], msg, "fehler", changes)
            if item:
                found.append(item)
    return found[:MAX_EACH]


def _ref_text(formula: str, ref: Pos) -> str:
    """Wie die Formel den Bezug schreibt (mit oder ohne $)."""
    m = re.search(rf"(?<![A-Za-z0-9_.$!:])(\$?){column_letters(ref[1])}(\$?){ref[0]}(?![0-9A-Za-z(!:])", formula)
    return m.group(0) if m else address(ref)


# --- Buchwert unter null: Die AfA läuft über die Nutzungsdauer ---

_BOOK = _rx(r"buchwert|restbuchwert|book value|carrying|\bnbv\b|valeur nette|valore netto")


def _book_value(ctx: _Ctx) -> list[Finding]:
    found = []
    by_col: dict[int, list[Pos]] = {}
    for pos, node, text in _formulas(ctx):
        by_col.setdefault(pos[1], []).append(pos)
    for col, cells in by_col.items():
        head = _col_head(ctx, col, cells[0][0])
        if not _search(_BOOK, head):
            continue
        negative = [p for p in cells if _number(ctx.value(p)) and ctx.value(p) < -0.005]
        if not negative:
            continue
        first = negative[0]
        before = (first[0] - 1, col)
        # Die AfA dieser Zeile: eine Spalte links mit AfA im Kopf.
        afa_cols = [c for c in by_col if c != col and _search(_AFA, _col_head(ctx, c, first[0]))
                    and not _search(_rx(r"kumul|summe|aufgelaufen|accumulated|cumul"), _col_head(ctx, c, first[0]))]
        msg = T(f"{address(first)} ({head}) wird negativ ({_fmt(ctx.value(first))}). Ein Buchwert unter null gibt es "
                f"nicht: Die Abschreibung läuft weiter, obwohl die Anlage schon ganz abgeschrieben ist",
                f"{address(first)} ({head}) turns negative ({_fmt(ctx.value(first))}). There is no such thing as a book "
                f"value below zero: depreciation keeps running although the asset is already fully depreciated")
        if len(afa_cols) == 1 and _number(ctx.value(before)):
            a = afa_cols[0]
            run = [(r, a) for r in range(first[0], max(p[0] for p in negative) + 1) if (r, a) in ctx.sheet.nodes]
            ends = [p for p in cells if p[0] < first[0] and _number(ctx.value(p)) and abs(ctx.value(p)) <= 0.005]
            stop = (ends[0][0] + 1, a) if ends else run[0] if run else None
            changes = {}
            for pos in run:
                text = ctx.sheet.cells[pos].formula
                changes[pos] = f"MIN({text},MAX(0,{column_letters(col)}{pos[0] - 1}))"
            if changes and _differs(ctx, changes) and stop is not None:
                when = _period_label(ctx, stop) or _row_name(ctx, stop)
                msg += T(f" (ab {address(stop)}{', ' + when if when else ''}). Mit MIN schreibt "
                         f"{_col_head(ctx, a, stop[0]) or 'die AfA'} höchstens den Buchwert der Zeile darüber ab, dann "
                         f"bleibt er bei 0.",
                         f" (from {address(stop)}{', ' + when if when else ''}). With MIN, "
                         f"{_col_head(ctx, a, stop[0]) or 'the depreciation'} writes off at most the book value of the row "
                         f"above, so it stays at 0.")
                item = _finding("buchwert", ctx, stop if stop in changes else run[0], msg, "fehler", changes)
                if item:
                    found.append(item)
                    continue
        found.append(Finding(f"buchwert:{address(first)}", msg + ".", "pruefen", ""))
    return found[:MAX_EACH]


# --- Jahreszins mit Monaten ---

_YEARLY = _rx(r"p\.?\s*a\.?|pro jahr|im jahr|jahrlich|jahreszins|per annum|annual|yearly|per year|\bannuel|annuo|\bjahr\b")
_MONTHS = _rx(r"monat|month|\bmois\b|\bmesi\b|\bmte\b")
_MONTHLY_RESULT = _rx(r"monatsrate|monatlich|pro monat|je monat|monthly|per month|mensuel|mensile")
_YEARS = _rx(r"jahre|years|\bjahr\b|\bans\b|\banni\b")
_ANNUITY = {"PMT": (0, 1), "IPMT": (0, 2), "PPMT": (0, 2), "PV": (0, 1), "FV": (0, 1), "NPER": (0, None),
            "CUMIPMT": (0, 1), "CUMPRINC": (0, 1)}


def _rate_period(ctx: _Ctx) -> list[Finding]:
    found = []
    for pos, node, text in _formulas(ctx):
        for call in formel.calls(node):
            if call.value not in _ANNUITY or len(call.args) < 2:
                continue
            ri, ni = _ANNUITY[call.value]
            rate = call.args[ri]
            if rate.kind != "ref":
                continue  # B3/12 ist schon umgerechnet
            rate_name = _name(ctx, rate.value)
            if not _search(_YEARLY, rate_name) or _search(_MONTHS, rate_name):
                continue
            nper = call.args[ni] if ni is not None else None
            nper_name = _name(ctx, nper.value) if nper is not None and nper.kind == "ref" else ""
            monthly_n = bool(nper_name) and _search(_MONTHS, nper_name)
            result = _name(ctx, pos)
            monthly_out = _search(_MONTHLY_RESULT, result)
            if not monthly_n and not (monthly_out and nper_name and _search(_YEARS, nper_name)):
                continue
            new = _swap_ref(text, rate.value, f"{{d1}}{column_letters(rate.value[1])}{{d2}}{rate.value[0]}/12")
            if new is None:
                continue
            if not monthly_n:
                new = _swap_ref(new, nper.value, f"{{d1}}{column_letters(nper.value[1])}{{d2}}{nper.value[0]}*12")
                if new is None:
                    continue
            changes = {pos: new}
            bad = _differs(ctx, changes)
            if not bad:
                continue
            why = (T(f"{address(nper.value)} ({nper_name}) zählt Monate",
                     f"but {address(nper.value)} ({nper_name}) counts months") if monthly_n else
                   T(f"{address(pos)} heißt „{result}“, die Laufzeit {address(nper.value)} ({nper_name}) zählt aber Jahre",
                     f"but {address(pos)} is called “{result}” while the term {address(nper.value)} ({nper_name}) "
                     f"counts years"))
            fixwhat = T("den Jahreszins durch 12", "divide the annual rate by 12") if monthly_n else \
                T("den Jahreszins durch 12 und die Jahre mal 12", "divide the annual rate by 12 and multiply the years by 12")
            right, now = _shown(ctx, pos, bad[2]), _shown(ctx, pos, bad[1])
            msg = T(f"{address(pos)} ({result}) rechnet {call.value} mit dem Jahreszins {address(rate.value)} "
                    f"({rate_name}), {why}. Zins und Laufzeit müssen dieselbe Periode haben: {fixwhat}. "
                    f"So wäre {address(pos)} {right} statt {now}.",
                    f"{address(pos)} ({result}) calculates {call.value} with the annual interest rate "
                    f"{address(rate.value)} ({rate_name}), {why}. Rate and term must use the same period: {fixwhat}. "
                    f"Then {address(pos)} would be {right} instead of {now}.")
            item = _finding("zinsperiode", ctx, pos, msg, "fehler", changes)
            if item:
                found.append(item)
            break
    return found[:MAX_EACH]


# --- MAPE ohne Betrag ---

_ABSOLUTE = _rx(r"\bmape\b|\bmae\b|\bmad\b|absolut|betrag|\babs\b|mittlere abweichung|absolute|wmape|smape|"
                r"durchschnittlicher fehler|mittlerer fehler|prognosefehler|forecast error|ecart absolu")


def _mape(ctx: _Ctx) -> list[Finding]:
    found = []
    for pos, node, text in _formulas(ctx):
        if node.kind != "call" or node.value not in ("AVERAGE", "AVERAGEA") or len(node.args) != 1 or \
                node.args[0].kind != "range":
            continue
        name = _row_name(ctx, pos) or _name(ctx, pos)
        if not _search(_ABSOLUTE, name) or _search(_rx(r"bias|verzerrung|mpe\b|signed|mit vorzeichen"), name):
            continue
        r1, c1, r2, c2 = node.args[0].value
        values = [ctx.value((r, c)) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
        values = [v for v in values if _number(v)]
        if len(values) < 3 or not any(v < 0 for v in values) or not any(v > 0 for v in values):
            continue
        area = text[text.index("(") + 1:text.rindex(")")].strip()
        new = f"SUMPRODUCT(ABS({area}))/COUNT({area})"
        changes = {pos: new}
        bad = _differs(ctx, changes)
        if not bad:
            continue
        negative, right, now = sum(v < 0 for v in values), _shown(ctx, pos, bad[2]), _shown(ctx, pos, bad[1])
        msg = T(f"{address(pos)} ({name}) mittelt die Fehler mit Vorzeichen: Zu hohe und zu tiefe Prognosen heben sich "
                f"auf ({negative} von {len(values)} sind negativ). Ein mittlerer absoluter Fehler nimmt "
                f"je Wert den Betrag: {right} statt {now}.",
                f"{address(pos)} ({name}) averages the errors with their signs: forecasts that are too high and too low "
                f"cancel each other out ({negative} of {len(values)} are negative). A mean absolute error takes the "
                f"absolute value of each one: {right} instead of {now}.")
        item = _finding("mape", ctx, pos, msg, "fehler", changes)
        if item:
            found.append(item)
    return found[:MAX_EACH]


# --- NBW über Jahr 0, untereinander ---

_ZERO = _rx(r"^(?:jahr|year|periode|period|t|annee|anno)?\s*0$|^heute$|^today$|^t\s*=\s*0$")


def _npv_zero(ctx: _Ctx) -> list[Finding]:
    found = []
    for pos, node, text in _formulas(ctx):
        calls = [c for c in formel.calls(node) if c.value == "NPV" and len(c.args) == 2 and c.args[1].kind == "range"]
        if len(calls) != 1:
            continue
        r1, c1, r2, c2 = calls[0].args[1].value
        if c1 != c2 or r2 - r1 < 2:
            continue  # nebeneinander prüft zeilenmodell.py
        first = (r1, c1)
        label = _period_label(ctx, first)
        if label is None or not _ZERO.search(_plain(label, keep_spaces=True).strip()):
            continue
        singles, _, _ = formel.references(node, ctx.sheet.name)
        if first in singles:
            continue
        span = _call_span(text, "NPV")
        if span is None:
            continue
        call_text = text[span[0]:span[1]]
        rng = re.compile(rf"(\$?){column_letters(c1)}(\$?){r1}:")
        if len(rng.findall(call_text)) != 1:
            continue
        call_new = rng.sub(lambda m: f"{m.group(1)}{column_letters(c1)}{m.group(2)}{r1 + 1}:", call_text)
        whole = span == (0, len(text))
        new = f"{address(first)}+{call_new}" if whole else text[:span[0]] + f"({address(first)}+{call_new})" + text[span[1]:]
        changes = {pos: new}
        bad = _differs(ctx, changes)
        if not bad:
            continue
        name, right, now = _row_name(ctx, pos) or _name(ctx, pos), _shown(ctx, pos, bad[2]), _shown(ctx, pos, bad[1])
        msg = T(f"{address(pos)} ({name}) zinst mit NBW auch {address(first)} "
                f"({label}) ab. NBW teilt schon den ersten Wert durch (1 + Zins), Jahr 0 fällt aber heute an und gehört "
                f"vor das NBW. So wäre {address(pos)} {right} statt {now}.",
                f"{address(pos)} ({name}) also discounts {address(first)} ({label}) with NPV. NPV already divides the "
                f"first value by (1 + rate), but year 0 falls due today and belongs in front of the NPV. Then "
                f"{address(pos)} would be {right} instead of {now}.")
        item = _finding("nbw0", ctx, pos, msg, "fehler", changes)
        if item:
            found.append(item)
    return found[:MAX_EACH]


def _period_label(ctx: _Ctx, pos: Pos) -> str | None:
    """Die Periode einer Zeile: die Zahl oder der Text links (Jahr 0, t=0, heute)."""
    for c in range(pos[1] - 1, 0, -1):
        cell = ctx.sheet.cells.get((pos[0], c))
        if cell is None or cell.formula:
            continue
        value = cell.value
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            head = _col_head(ctx, c, pos[0])
            return f"{head} {_fmt(value)}" if head else _fmt(value)
        if isinstance(value, str) and value.strip():
            text = value.strip()
            if re.fullmatch(r"-?\d+(?:[.,]0+)?", text):
                head = _col_head(ctx, c, pos[0])
                return f"{head} {int(float(text.replace(',', '.')))}" if head else text
            return text
    return None


def _call_span(formula: str, name: str) -> tuple[int, int] | None:
    plain = re.sub(r'"(?:[^"]|"")*"', lambda m: " " * len(m.group(0)), formula)
    match = re.search(rf"(?<![A-Za-z0-9_.]){name}\(", plain, re.IGNORECASE)
    if not match:
        return None
    depth = 0
    for i in range(match.end() - 1, len(plain)):
        if plain[i] == "(":
            depth += 1
        elif plain[i] == ")":
            depth -= 1
            if depth == 0:
                return match.start(), i + 1
    return None


# --- Anteil an einer Zeile („% vom Umsatz“) ohne $ ---

_BASE = re.compile(r"(?:%|prozent|anteil|quote|share|percent|pourcentage|percentuale)\s*(?:vom|von|am|an|der|des|of|du|de|"
                   r"del|della|in\s+%\s+(?:vom|von|des))\s+(.+)$", re.IGNORECASE)


def _share_base(ctx: _Ctx) -> list[Finding]:
    found = []
    by_col: dict[int, list[Pos]] = {}
    for pos, node, text in _formulas(ctx):
        if pos[0] in ctx.data:
            by_col.setdefault(pos[1], []).append(pos)
    for col, cells in by_col.items():
        head = ctx.plan.headers.get(col, "")
        m = _BASE.search(head.strip())
        if not m or len(cells) < 3:
            continue
        word = _plain(m.group(1))
        divs = []
        for pos in cells:
            node = ctx.sheet.nodes[pos]
            if not (node.kind == "op" and node.value == "/" and all(a.kind == "ref" for a in node.args)):
                break
            top, bottom = node.args[0].value, node.args[1].value
            if top[0] != pos[0] or bottom[1] != top[1]:
                break
            divs.append((pos, top, bottom))
        else:
            vcol = divs[0][1][1]
            if any(t[1] != vcol for _, t, _ in divs):
                continue
            base = [r for r in sorted(ctx.data | {ctx.plan.header_row + 1}) if word and
                    _plain(_row_name(ctx, (r, vcol))) == word]
            if len(base) != 1:
                continue
            b = base[0]
            wrong = [(p, bottom) for p, _, bottom in divs if bottom[0] != b]
            if not wrong:
                continue
            changes = {p: f"{column_letters(vcol)}{p[0]}/{column_letters(vcol)}${b}" for p, _, _ in divs}
            bad = _differs(ctx, changes)
            if not bad:
                continue
            pos, now, right = bad
            bottom = dict((p, bt) for p, bt in wrong)[pos] if pos in dict(wrong) else wrong[0][1]
            label, said, fixed = _row_name(ctx, bottom) or address(bottom), m.group(1).strip(), f"{column_letters(vcol)}${b}"
            should, shows = _shown(ctx, pos, right), _shown(ctx, pos, now)
            msg = T(f"{address(wrong[0][0])} ({head}) teilt durch {address(bottom)} "
                    f"({label}), der Name sagt aber „{said}“, also "
                    f"{address((b, vcol))}. Ohne $ im Nenner rutscht der Bezug beim Herunterziehen mit. Mit "
                    f"{fixed} wäre {address(pos)} {should} statt {shows}.",
                    f"{address(wrong[0][0])} ({head}) divides by {address(bottom)} ({label}), but the name says "
                    f"“{said}”, so {address((b, vcol))}. Without $ in the denominator, the reference shifts when you "
                    f"fill down. With {fixed}, {address(pos)} would be {should} instead of {shows}.")
            item = _finding("basis", ctx, wrong[0][0], msg, "fehler", changes)
            if item:
                found.append(item)
    return found[:MAX_EACH]



# --- Steuern addiert: Jahresüberschuss größer als das Ergebnis vor Steuern ---

_TAX = _rx(r"steuer|\btax|\btaxes\b|impot|imposte|\bertragsteuer")
_AFTER_TAX = _rx(r"jahresuberschuss|jahresfehlbetrag|jahresergebnis|nach steuern|net income|net profit|nettogewinn|"
                 r"reingewinn|periodenergebnis|konzernergebnis|after tax|benefice net|resultat net|utile netto|"
                 r"risultato netto|\bnet earnings")
_BEFORE_TAX = _rx(r"vor steuern|\bebt\b|pre.?tax|before tax|avant impot|ante imposte")


def _tax_added(ctx: _Ctx) -> list[Finding]:
    found = []
    done: set[int] = set()
    for pos, node, text in _formulas(ctx):
        terms = _signed_terms(node, 1)
        if not terms or len(terms) < 2 or pos[0] in done:
            continue
        name = _row_name(ctx, pos)
        taxes = [(ref, sign) for ref, sign in terms if sign > 0 and _search(_TAX, _row_name(ctx, ref))
                 and not _search(_AFTER_TAX, _row_name(ctx, ref)) and not _search(_BEFORE_TAX, _row_name(ctx, ref))]
        before = [ref for ref, sign in terms if sign > 0 and _search(_BEFORE_TAX, _row_name(ctx, ref))]
        if len(taxes) != 1 or not (_search(_AFTER_TAX, name) or before):
            continue
        tax = taxes[0][0]
        value, base = ctx.value(tax), ctx.value(before[0]) if before else None
        if not _number(value) or value <= 0 or (base is not None and (not _number(base) or base <= 0)):
            continue
        # Die Steuern stehen positiv und werden addiert. Stehen die anderen Kosten im Modell negativ (Zinsaufwand
        # (504)), gehört das Minus in die Steuerzeile; sonst zieht die Summe sie ab.
        others = [ref for ref, sign in terms if ref != tax and sign > 0 and ref not in before]
        negative_costs = _costs_negative(ctx, before[0] if before else None)
        row_t = [(tax[0], c) for c in range(1, tax[1] + 30) if (tax[0], c) in ctx.sheet.nodes
                 and ctx.sheet.nodes[(tax[0], c)] is not None and ctx.sheet.cells[(tax[0], c)].formula]
        changes: dict[Pos, str] = {}
        if negative_costs and row_t and not others:
            for p in row_t:
                changes[p] = _negated(ctx.sheet.cells[p].formula)
            where = T(f"Die übrigen Kosten stehen im Modell negativ (etwa {negative_costs}), die Steuern in "
                      f"{address(row_t[0])}:{address(row_t[-1])} aber positiv.",
                      f"The other costs are negative in the model (e.g. {negative_costs}), but the taxes in "
                      f"{address(row_t[0])}:{address(row_t[-1])} are positive.")
            target = row_t[0]
        else:
            new = _flip(text, tax, 1)
            if new is None:
                continue
            changes[pos] = new
            where = T("Positive Steuern werden abgezogen, nicht addiert.", "Positive taxes are subtracted, not added.")
            target = pos
        bad = _differs(ctx, changes)
        if not bad:
            continue
        after = ctx.value(pos)
        label, earlier = name or _name(ctx, pos), f" {address(before[0])} ({_fmt(base)})" if before else ""
        right = _fmt(round(_rerun(ctx, changes, pos), 2))
        msg = T(f"{address(pos)} ({label}) addiert die Steuern {address(tax)} "
                f"({_fmt(value)}) und ist damit größer als das Ergebnis vor Steuern"
                f"{earlier}: {_fmt(after)}. {where} Mit dem "
                f"richtigen Vorzeichen wäre {address(pos)} {right}.",
                f"{address(pos)} ({label}) adds the taxes {address(tax)} ({_fmt(value)}) and is therefore greater than "
                f"the earnings before taxes{earlier}: {_fmt(after)}. {where} With the correct sign, {address(pos)} "
                f"would be {right}.")
        item = _finding("steuern", ctx, target, msg, "fehler", changes)
        if item:
            found.append(item)
            done.add(pos[0])
    return found[:MAX_EACH]


def _costs_negative(ctx: _Ctx, before: Pos | None) -> str:
    """Ein Aufwand im Modell, der negativ dasteht (Zinsaufwand (504)): „D31 (Zinsaufwand) -504“, sonst ""."""
    if before is None:
        return ""
    node = ctx.sheet.nodes.get(before)
    terms = _signed_terms(node, 1) if node is not None else None
    for ref, sign in terms or []:
        v = ctx.value(ref)
        if sign > 0 and _number(v) and v < 0 and _search(_COSTS, _row_name(ctx, ref)):
            return f"{address(ref)} ({_row_name(ctx, ref)}) {_fmt(round(v, 2))}"
    return ""


_COSTS = _rx(r"aufwand|kosten|zins|abschreibung|afa|expense|cost|interest|depreciation|charges|costi|oneri")


def _negated(formula: str) -> str:
    node = formel.parse("=" + formula)
    if node.kind in ("call", "ref") or (node.kind == "op" and node.value in ("*", "/")):
        return "-" + formula
    return f"-({formula})"


def _rerun(ctx: _Ctx, changes: dict[Pos, str], pos: Pos) -> float:
    """Der Wert von pos, wenn changes eingesetzt sind (eine Ebene: pos selbst oder die geänderten Zellen darin)."""
    if pos in changes:
        return ctx.run(changes[pos])
    text = ctx.sheet.cells[pos].formula
    for p, new in changes.items():
        value = ctx.run(new)
        text = re.sub(rf"(?<![A-Za-z0-9_.$!:])\$?{column_letters(p[1])}\$?{p[0]}(?![0-9A-Za-z(!:])",
                      f"({_fmt_num(value)})", text)
    return ctx.run(text)


def _fmt_num(v: float) -> str:
    return repr(float(v))


# --- Gordon-Formel ohne (1 + g) ---

_GROWTH = _rx(r"wachstum|growth|\bg\b|croissance|crescita|ewige rente")
_DISCOUNT = _rx(r"wacc|diskont|kapitalkosten|kalkulationszins|zins|discount|cost of capital|\brate\b|rendite|taux|tasso")
_NEXT = _rx(r"folgejahr|n\s*\+\s*1|t\s*\+\s*1|nachstes jahr|next year|normalis|ewig|terminal|fortfuhrung|steady")
_TERMINAL = _rx(r"terminal|endwert|restwert|fortfuhrungswert|ewige rente|gordon|continuing value|valeur terminale|"
                r"valore terminale|\btv\b")


def _gordon(ctx: _Ctx) -> list[Finding]:
    """Der Terminal Value nach Gordon: Cashflow des Jahres danach durch (Zins − Wachstum). Gefunden wird jede
    Abweichung von dieser Form, nicht ein bestimmter Tippfehler: der Cashflow des letzten Jahres ohne × (1 + g), ein
    Nenner (Zins + Wachstum) oder (Wachstum − Zins), und beides zusammen."""
    found = []
    for pos, node, text in _formulas(ctx):
        if node.kind != "op" or node.value != "/" or len(node.args) != 2:
            continue
        top, bottom = node.args
        if not (bottom.kind == "op" and bottom.value in "+-" and len(bottom.args) == 2
                and all(a.kind == "ref" for a in bottom.args)):
            continue
        first, second = (a.value for a in bottom.args)
        roles = {p: ("g" if _search(_GROWTH, _name(ctx, p)) else "r" if _search(_DISCOUNT, _name(ctx, p)) else "")
                 for p in (first, second)}
        if sorted(roles.values()) != ["g", "r"]:
            continue
        rate = first if roles[first] == "r" else second
        growth = second if rate == first else first
        name = _row_name(ctx, pos) or _name(ctx, pos)
        if not _search(_TERMINAL, name):
            continue
        r_t, g_t = _ref_text(text, rate), _ref_text(text, growth)
        wrong: list[str] = []
        new = text
        # Der Nenner: Zins minus Wachstum, in dieser Reihenfolge.
        if not (bottom.value == "-" and rate == first):
            shown = f"({address(first)} {bottom.value} {address(second)})"
            inner = rf"\(\s*{re.escape(_ref_text(text, first))}\s*{re.escape(bottom.value)}\s*{re.escape(_ref_text(text, second))}\s*\)"
            if len(re.findall(inner, new)) != 1:
                continue
            new = re.sub(inner, lambda m: f"({r_t}-{g_t})", new)
            wrong.append(T(f"Der Nenner ist {shown}, richtig ist (Zins − Wachstum) = ({address(rate)} − {address(growth)})",
                           f"The denominator is {shown}; correct is (discount rate − growth) = ({address(rate)} − "
                           f"{address(growth)})"))
        # Der Zähler: der Cashflow des Jahres nach dem letzten, also das letzte Jahr × (1 + g).
        if top.kind == "ref" and growth not in _refs(top):
            flow = top.value
            flow_name = _row_name(ctx, flow) or _name(ctx, flow)
            head = _col_head(ctx, flow[1], flow[0])
            if not (_search(_NEXT, flow_name) or _search(_NEXT, head)) and _last_of_series(ctx, flow):
                swapped = _swap_ref(new, flow, f"{{d1}}{column_letters(flow[1])}{{d2}}{flow[0]}*(1+{g_t})")
                if swapped is None:
                    continue
                new = swapped
                wrong.append(T(f"Der Zähler ist der Cashflow {address(flow)} ({flow_name}{' ' + head if head else ''}) des "
                               f"letzten Jahres; die Formel braucht den des Jahres danach: {address(flow)} × (1 + {address(growth)})",
                               f"The numerator is the cash flow {address(flow)} ({flow_name}{' ' + head if head else ''}) "
                               f"of the last year; the formula needs the one of the following year: {address(flow)} × "
                               f"(1 + {address(growth)})"))
        if not wrong:
            continue
        changes = {pos: new}
        bad = _differs(ctx, changes)
        if not bad:
            continue
        right, now = _shown(ctx, pos, bad[2]), _shown(ctx, pos, bad[1])
        msg = T(f"{address(pos)} ({name}) rechnet den Terminal Value nach Gordon: Cashflow des Folgejahrs ÷ (Zins − "
                f"Wachstum). {'. '.join(wrong)}. So wäre {address(pos)} {right} statt "
                f"{now}.",
                f"{address(pos)} ({name}) calculates the terminal value with the Gordon formula: next year's cash flow ÷ "
                f"(discount rate − growth). {'. '.join(wrong)}. Then {address(pos)} would be {right} instead of {now}.")
        item = _finding("gordon", ctx, pos, msg, "fehler", changes)
        if item:
            found.append(item)
    return found[:MAX_EACH]


_FACTOR = _rx(r"diskontierungsfaktor|abzinsungsfaktor|barwertfaktor|diskontfaktor|discount factor|\bdf\b|"
              r"facteur d.actualisation|fattore di sconto")


def _tv_discount(ctx: _Ctx) -> list[Finding]:
    """Der Barwert des Terminal Value: Der Terminal Value gilt am Ende des letzten Planjahrs, aus dessen Cashflow er
    rechnet. Abgezinst wird er darum mit dem Faktor genau dieses Jahres, nicht eines früheren oder späteren."""
    found = []
    for pos, node, text in _formulas(ctx):
        if node.kind != "op" or node.value != "*" or len(node.args) != 2 or not all(a.kind == "ref" for a in node.args):
            continue
        tv = next((a.value for a in node.args if _search(_TERMINAL, _row_name(ctx, a.value) or _name(ctx, a.value))
                   and ctx.sheet.nodes.get(a.value) is not None), None)
        factor = next((a.value for a in node.args if a.value != tv and _search(_FACTOR, _row_name(ctx, a.value) or "")), None)
        if tv is None or factor is None:
            continue
        lasts = {p[1] for p in _refs(ctx.sheet.nodes[tv]) if _last_of_series(ctx, p) and p[0] != tv[0]
                 and _number(ctx.value((factor[0], p[1])))}
        if len(lasts) != 1:
            continue
        year = next(iter(lasts))
        right = (factor[0], year)
        if year == factor[1] or not _number(ctx.value(right)):
            continue
        new = _swap_ref(text, factor, f"{{d1}}{column_letters(year)}{{d2}}{factor[0]}")
        if new is None:
            continue
        changes = {pos: new}
        bad = _differs(ctx, changes)
        if not bad:
            continue
        head, wrong_head = _col_head(ctx, year, factor[0]), _col_head(ctx, factor[1], factor[0])
        name = _row_name(ctx, pos) or _name(ctx, pos)
        should, now = _shown(ctx, pos, bad[2]), _shown(ctx, pos, bad[1])
        msg = T(f"{address(pos)} ({name}) zinst den Terminal Value {address(tv)} mit {address(factor)}"
                f"{f' ({wrong_head})' if wrong_head else ''} ab. Der Terminal Value rechnet aber aus dem letzten Jahr"
                f"{f' {head}' if head else ''} und gilt an dessen Ende, sein Faktor steht in {address(right)}. So wäre "
                f"{address(pos)} {should} statt {now}.",
                f"{address(pos)} ({name}) discounts the terminal value {address(tv)} with {address(factor)}"
                f"{f' ({wrong_head})' if wrong_head else ''}. But the terminal value is based on the last year"
                f"{f' {head}' if head else ''} and applies at its end; its factor is in {address(right)}. Then "
                f"{address(pos)} would be {should} instead of {now}.")
        item = _finding("tvfaktor", ctx, pos, msg, "fehler", changes)
        if item:
            found.append(item)
    return found[:MAX_EACH]


_OPENING = _rx(r"anfangsbestand|\banfang\b|\bbeginn\b|eroffnung|opening|beginning|\bbop\b|debut|iniziale|inizio")
_CLOSING = _rx(r"endbestand|\bende\b|schlussbestand|closing|ending|\beop\b|\bfin\b|finale")
_STOCK_WORDS = r"(?:anfangsbestand|endbestand|schlussbestand|anfang|beginn|eroffnung|ende|opening|closing|beginning|ending|bop|eop)"


def _stock(name: str) -> str:
    """„Bankdarlehen Ende“ → „bankdarlehen“: der Bestand ohne das Wort für Anfang oder Ende."""
    return re.sub(rf"\b{_STOCK_WORDS}\b", " ", _plain(name, keep_spaces=True)).strip()


def _row_runs(changes: dict[Pos, str]) -> list[list[Pos]]:
    runs: list[list[Pos]] = []
    for pos in sorted(changes):
        if runs and runs[-1][-1][0] == pos[0] and runs[-1][-1][1] == pos[1] - 1:
            runs[-1].append(pos)
        else:
            runs.append([pos])
    return runs


def _opening(ctx: _Ctx) -> list[Finding]:
    """Der Name sagt „auf Anfangsbestand“ (Zins auf den Anfangsbestand), die Formel liest aber den Endbestand
    desselben Bestands, und die Zeile mit dem Anfangsbestand gibt es. Der Endbestand hängt oft vom Ergebnis ab, so
    rechnet der Zins mit einer Zahl, die er selbst verändert."""
    changes: dict[Pos, str] = {}
    why: dict[Pos, tuple[Pos, Pos]] = {}
    for pos, node, text in _formulas(ctx):
        name = _row_name(ctx, pos)
        if not name or not _search(_OPENING, name) or _search(_CLOSING, name):
            continue
        new = text
        for ref in dict.fromkeys(_refs(node)):
            other = _row_name(ctx, ref)
            # Der Endbestand der Vorperiode ist der Anfangsbestand: nur derselbe Zeitraum (dieselbe Spalte) ist falsch.
            if ref[1] != pos[1] or not other or not _search(_CLOSING, other) or _search(_OPENING, other):
                continue
            base = _stock(other)
            start = next((r for r in range(ref[0] - 15, ref[0] + 16) if r != ref[0]
                          and _search(_OPENING, _row_name(ctx, (r, ref[1])) or "")
                          and _stock(_row_name(ctx, (r, ref[1]))) == base and _number(ctx.value((r, ref[1])))), None)
            if start is None or not base:
                continue
            swapped = _swap_ref(new, ref, f"{{d1}}{column_letters(ref[1])}{{d2}}{start}")
            if swapped is None:
                break
            new = swapped
            why[pos] = (ref, (start, ref[1]))
        # MITTELWERT(C53;C59) wird MITTELWERT(C53;C53): das ist einfach C53.
        new = re.sub(r"(?:AVERAGE|MIN|MAX)\((\$?[A-Z]{1,3}\$?\d+),\s*\1\)", r"\1", new)
        if new != text:
            changes[pos] = new
    found = []
    for run in _row_runs(changes):
        part = {p: changes[p] for p in run}
        bad = _differs(ctx, part)
        if not bad:
            continue
        first = run[0]
        end, start = why[first]
        where = address(first) if len(run) == 1 else f"{address(first)}:{address(run[-1])}"
        right, now = _shown(ctx, bad[0], bad[2]), _shown(ctx, bad[0], bad[1])
        msg = T(f"{where} ({_row_name(ctx, first)}) liest {address(end)} ({_row_name(ctx, end)}), den Bestand am Ende "
                f"der Periode. Der Name sagt aber Anfangsbestand, und der steht in {address(start)} "
                f"({_row_name(ctx, start)}). So wäre {address(bad[0])} {right} statt "
                f"{now}.",
                f"{where} ({_row_name(ctx, first)}) reads {address(end)} ({_row_name(ctx, end)}), the balance at the end "
                f"of the period. But the name says opening balance, and that is in {address(start)} "
                f"({_row_name(ctx, start)}). Then {address(bad[0])} would be {right} instead of {now}.")
        item = _finding("anfang", ctx, first, msg, "fehler", part)
        if item:
            found.append(item)
    return found[:MAX_EACH]


_QUALIFIER_SKIP = {"zinssatz", "satz", "zins", "rate", "in", "auf", "fur", "der", "die", "das", "p", "a", "von", "je", "pro"}


def _words_of(name: str) -> set[str]:
    return {w for w in re.findall(r"[a-z0-9ø]+", _plain(name, keep_spaces=True)) if len(w) > 2}


def _qualifier(ctx: _Ctx) -> list[Finding]:
    """Zwei Annahmen mit demselben Namen bis auf ein Wort (Zinssatz Guthaben, Zinssatz Kredit). Nennt die Formel das
    andere Wort (Zinsaufwand Kredit) und nicht das eigene, liest sie die falsche."""
    inputs = {p: _row_name(ctx, p) for p, v in ctx.sheet.inputs.items()
              if p not in ctx.sheet.nodes and _number(v) and not isinstance(v, bool)}
    groups: dict[tuple, list[Pos]] = {}
    for p, n in inputs.items():
        words = _words_of(n)
        if len(words) < 2:
            continue
        for w in words:
            groups.setdefault((p[1], frozenset(words - {w})), []).append(p)
    pairs: dict[Pos, list[tuple[Pos, str, str]]] = {}
    for (_, rest), members in groups.items():
        if len(members) < 2 or not rest:
            continue
        for a in members:
            own = (_words_of(inputs[a]) - rest)
            for b in members:
                if a != b and abs(a[0] - b[0]) <= 12:
                    pairs.setdefault(a, []).append((b, next(iter(own)), next(iter(_words_of(inputs[b]) - rest))))
    found = []
    covered: set[Pos] = set()
    for pos, node, text in _formulas(ctx):
        name = _row_name(ctx, pos)
        if not name:
            continue
        words = _words_of(name)
        for ref in dict.fromkeys(_refs(node)):
            for other, mine, theirs in pairs.get(ref, ()):
                if mine in _QUALIFIER_SKIP or theirs in _QUALIFIER_SKIP:
                    continue
                if theirs in words and mine not in words:
                    new = _swap_ref(text, ref, f"{{d1}}{column_letters(other[1])}{{d2}}{other[0]}")
                    if new is None:
                        continue
                    if pos in covered:
                        break
                    changes = {pos: new}
                    for c in range(pos[1] + 1, pos[1] + 60):
                        nxt = ctx.sheet.cells.get((pos[0], c))
                        more = _swap_ref(nxt.formula, ref, f"{{d1}}{column_letters(other[1])}{{d2}}{other[0]}") \
                            if nxt is not None and nxt.formula else None
                        if more is None:
                            break
                        changes[(pos[0], c)] = more
                    bad = _differs(ctx, changes)
                    if not bad:
                        continue
                    last = max(changes)
                    span = address(pos) if len(changes) == 1 else f"{address(pos)}:{address(last)}"
                    right, now = _shown(ctx, bad[0], bad[2]), _shown(ctx, bad[0], bad[1])
                    msg = T(f"{span} ({name}) rechnet mit {address(ref)} ({inputs[ref]}). Der Name sagt aber "
                            f"„{theirs.capitalize()}“, dafür gibt es {address(other)} ({inputs[other]}). So wäre "
                            f"{address(bad[0])} {right} statt {now}.",
                            f"{span} ({name}) uses {address(ref)} ({inputs[ref]}). But the name says "
                            f"“{theirs.capitalize()}”, and there is {address(other)} ({inputs[other]}) for that. Then "
                            f"{address(bad[0])} would be {right} instead of {now}.")
                    item = _finding("annahme", ctx, pos, msg, "fehler", changes)
                    if item:
                        covered.update(changes)
                        found.append(item)
                    break
    return found[:MAX_EACH]


def _last_of_series(ctx: _Ctx, pos: Pos) -> bool:
    """Ist pos das letzte Jahr einer Reihe: links (oder darüber) Zahlen, rechts (oder darunter) keine?"""
    for before, after in (((pos[0], pos[1] - 1), (pos[0], pos[1] + 1)), ((pos[0] - 1, pos[1]), (pos[0] + 1, pos[1]))):
        if _number(ctx.value(before)) and not _number(ctx.value(after)):
            return True
    return False


# --- Eine Auszahlung mit dem falschen Vorzeichen: Investition, Tilgung, Kosten positiv addiert ---

_OUTFLOW = _rx(r"investition|\binvest|capex|auszahlung|tilgung|steuer|zinsaufwand|zinsen|dividende|ausschuttung|kosten|"
               r"aufwand|ausgaben|\bmaterial|personal|\bmiete|repayment|amortization of debt|purchase|\btax|interest|"
               r"dividend|\bcost|expense|depense|investissement|impot|costi|spese|investiment")
# Was zufließt oder zurückgerechnet wird: Erlös, Aufnahme eines Darlehens, AfA als Rückrechnung im Cashflow.
_INFLOW = _rx(r"erlos|ertrag|umsatz|einzahlung|zufluss|aufnahme|zuschuss|erstattung|ruckrechnung|ruckfluss|abschreibung|"
              r"\bafa\b|revenue|income|proceeds|refund|add.?back|depreciation|vor steuern|nach steuern|\bebt|ebit|"
              r"uberschuss|ergebnis|cashflow|cash flow|summe|gesamt|total|saldo|netto|brutto|marge|kumuliert")


def _sum_terms(ctx: _Ctx, node) -> list[tuple[Pos, int]] | None:
    """Die Summanden einer reinen Summe mit Vorzeichen: =B43+B44-B45 oder =SUMME(B22:B25) (eine Spalte oder Zeile)."""
    terms = _signed_terms(node, 1)
    if terms is not None:
        return terms
    if node.kind == "call" and node.value == "SUM" and len(node.args) == 1 and node.args[0].kind == "range":
        r1, c1, r2, c2 = node.args[0].value
        if (r1 == r2 or c1 == c2) and (r2 - r1) + (c2 - c1) < 60:
            return [((r, c), 1) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1) if (r, c) in ctx.sheet.cells]
    return None


def _outflows(ctx: _Ctx) -> list[Finding]:
    """Ein Modell schreibt Auszahlungen einheitlich: negativ und addiert, oder positiv und abgezogen. Steht in einer
    Summe eine Auszahlung positiv und wird addiert, während eine andere Auszahlung derselben Summe negativ dasteht,
    hat die eine das falsche Vorzeichen (Investition 600000 im Netto-Cashflow neben Tilgung -60000)."""
    found = []
    done: set[Pos] = set()
    for pos, node, text in _formulas(ctx):
        terms = _sum_terms(ctx, node)
        if not terms or len(terms) < 2:
            continue
        added = []
        for ref, sign in terms:
            name = _row_name(ctx, ref) if ref[0] != pos[0] else _col_head(ctx, ref[1], ref[0])
            if sign > 0 and _search(_OUTFLOW, name) and not _search(_INFLOW, name):
                value = ctx.value(ref)
                if _number(value) and value != 0:
                    added.append((ref, name, value))
        negative = [a for a in added if a[2] < 0]
        positive = [a for a in added if a[2] > 0]
        if len(positive) != 1 or not negative or positive[0][0] in done:
            continue
        ref, name, value = positive[0]
        cell = ctx.sheet.cells.get(ref)
        changes: dict[Pos, str] = {}
        if cell is not None and cell.formula:
            # Die Zeile der Auszahlung rechnet: Ihr gehört das Minus, wie den anderen Auszahlungen.
            row = [p for p in _same_line(ctx, ref) if _number(ctx.value(p)) and ctx.value(p) > 0]
            for p in row:
                changes[p] = _negated(ctx.sheet.cells[p].formula)
            target = row[0]
            span = f"{address(row[0])}{':' + address(row[-1]) if len(row) > 1 else ''}"
            how = T(f"Gib {span} das Minus wie den übrigen "
                    f"Auszahlungen", f"Give {span} a minus sign like the other outflows")
        else:
            new = _flip(text, ref, 1)
            if new is None:
                continue
            changes[pos] = new
            target = pos
            how = T(f"Zieh {address(ref)} in {address(pos)} ab", f"Subtract {address(ref)} in {address(pos)}")
        after = _rerun(ctx, changes, pos)
        now = ctx.value(pos)
        if not _number(now) or after is None or _close(now, after):
            continue
        others = T(", ", "; ").join(f"{address(p)} ({n}) {_fmt(round(v, 2))}" for p, n, v in negative[:2])
        label, amount, right, shows = (_row_name(ctx, pos) or _name(ctx, pos), _fmt(round(value, 2)),
                                       _fmt(round(after, 2)), _fmt(round(now, 2)))
        msg = T(f"{address(pos)} ({label}) addiert {address(ref)} ({name}) mit "
                f"{amount}, also als Zufluss. Die übrigen Auszahlungen derselben Rechnung stehen negativ "
                f"da ({others}). {how}: Dann wäre {address(pos)} {right} statt {shows}.",
                f"{address(pos)} ({label}) adds {address(ref)} ({name}) as an inflow: {amount}. The other "
                f"outflows in the same calculation are negative ({others}). {how}: then {address(pos)} would be "
                f"{right} instead of {shows}.")
        item = _finding("auszahlung", ctx, target, msg, "fehler", changes)
        if item:
            found.append(item)
            done.add(ref)
    return found[:MAX_EACH]


def _same_line(ctx: _Ctx, pos: Pos) -> list[Pos]:
    """pos und die Zellen links und rechts daneben mit derselben Formel (die Jahre einer Zeile)."""
    from .workbook import pattern

    shape = pattern(ctx.sheet.cells[pos].formula, pos, "Zeile")
    out = [pos]
    for step in (-1, 1):
        c = pos[1] + step
        while (cell := ctx.sheet.cells.get((pos[0], c))) is not None and cell.formula \
                and pattern(cell.formula, (pos[0], c), "Zeile") == shape:
            out.append((pos[0], c))
            c += step
    return sorted(out, key=lambda p: p[1])


_VALUE_IN = _rx(r"barwert|terminal|restwert|einzahlung|zufluss|present value|\bpv\b|cash inflow|valeur actuelle|"
                r"valore attuale")
_REDUCES = _rx(r"schmalerung|minderung|rabatt|skonto|retoure|gutschrift|abzug|korrektur|abschlag|wertberichtigung")
_DIFFERENCE = _rx(r"abweichung|differenz|delta|variance|varianz|veranderung|ecart|scostamento|vergleich|vs\.?\b")


def _inflows(ctx: _Ctx) -> list[Finding]:
    """Das Gegenstück zu _outflows: Ein Kapitalwert addiert die Barwerte. Wird einer davon abgezogen, während ein
    anderer Barwert derselben Summe addiert wird (=B37-B38-B18: Barwert Terminalwert abgezogen), ist es ein Fehler."""
    found = []
    for pos, node, text in _formulas(ctx):
        terms = _sum_terms(ctx, node)
        if not terms or len(terms) < 2:
            continue
        name = _row_name(ctx, pos) or _name(ctx, pos)
        if _search(_DIFFERENCE, name):
            continue
        values = []
        for ref, sign in terms:
            label = _row_name(ctx, ref) if ref[0] != pos[0] else _col_head(ctx, ref[1], ref[0])
            value = ctx.value(ref)
            if _search(_VALUE_IN, label) and not _search(_REDUCES, label) and _number(value) and value > 0:
                values.append((ref, sign, label, value))
        plus = [v for v in values if v[1] > 0]
        minus = [v for v in values if v[1] < 0]
        if len(minus) != 1 or not plus:
            continue
        ref, _, label, value = minus[0]
        new = _flip(text, ref, -1)
        if new is None:
            continue
        changes = {pos: new}
        bad = _differs(ctx, changes)
        if not bad:
            continue
        others = ", ".join(f"{address(p)} ({n})" for p, _, n, _ in plus[:2])
        amount, right, now = _fmt(round(value, 2)), _shown(ctx, pos, bad[2]), _shown(ctx, pos, bad[1])
        msg = T(f"{address(pos)} ({name}) zieht {address(ref)} ({label}, {amount}) ab, addiert aber "
                f"{others}. Ein Barwert zählt dazu wie die anderen: {address(pos)} wäre {right} statt "
                f"{now}.",
                f"{address(pos)} ({name}) subtracts {address(ref)} ({label}, {amount}) but adds {others}. A present "
                f"value is added like the others: {address(pos)} would be {right} instead of {now}.")
        item = _finding("barwert", ctx, pos, msg, "fehler", changes)
        if item:
            found.append(item)
    return found[:MAX_EACH]


# --- Ein Satz als ganze Zahl: WACC 9 statt 9 % ---

_RATE_NAME = _rx(r"(?<!um)satz|zins|wacc|rendite|quote|wachstum|marge|inflation|diskont|kapitalkosten|\brate\b|yield|growth|"
                 r"margin|discount|taux|tasso|crescita|steigerung|anteil|pramie|premium|\bp\.?\s?a\.?(\s|$)|%")
_AMOUNT_NAME = _rx(r"€|\beur\b|chf|usd|\$|\bjahre\b|\bmonate\b|\btage\b|stunden|anzahl|stuck|basispunkt|\bbp\b|punkte|years|months|"
                r"days|count|\(x\)|faktor|tsd|mio")


def _percent_shown(cell) -> bool:
    return cell is not None and not cell.formula and str(cell.value).strip().endswith("%")


def _whole_rates(ctx: _Ctx) -> list[Finding]:
    """Ein Satz, der als ganze Zahl dasteht (9), wo die Sätze daneben Prozent sind (8 %, 5 %): Excel rechnet mit
    900 %. Nur, wenn keine Formel ihn durch 100 teilt."""
    from .analyse import changed_by, said

    found = []
    for pos, cell in sorted(ctx.sheet.cells.items()):
        if cell.formula or pos not in ctx.sheet.inputs:
            continue
        value = ctx.sheet.inputs.get(pos)
        if not _number(value) or not 1 < value <= 100 or _percent_shown(cell):
            continue
        near = [(r, pos[1]) for r in range(pos[0] - 12, pos[0] + 13) if r != pos[0]]
        siblings = [p for p in near if _percent_shown(ctx.sheet.cells.get(p)) and _number(ctx.sheet.inputs.get(p))
                    and abs(ctx.sheet.inputs[p]) < 1]
        if len(siblings) < 2:
            continue
        name = _row_name(ctx, pos)
        if not name or not _search(_RATE_NAME, name) or _search(_AMOUNT_NAME, name):
            continue
        unit = _row_unit(ctx, pos)
        if unit and not unit.startswith("%"):
            continue
        readers = sorted(ctx.sheet.dependents.get(pos, ()))
        if any(re.search(rf"\$?{column_letters(pos[1])}\$?{pos[0]}\s*(/\s*100|\*\s*(1%|0[.,]01))", ctx.sheet.cells[r].formula or "")
               for r in readers if r in ctx.sheet.cells):
            continue
        right = float(f"{value / 100:.12g}")
        fix = einsetzen.values(ctx.sheet, {pos: right}, siblings[0], f"{_fmt(value)} %")
        if fix is None:
            continue
        changed = changed_by(ctx.sheet, ctx.sheet.dependents.get(pos, ()), {pos: right}, near=True)
        effect = ""
        if changed:
            shown = [T(f"{ctx.sheet.label(p)} {said(after)} statt {said(before)}",
                       f"{ctx.sheet.label(p)} would be {said(after)} instead of {said(before)}")
                     for p, before, after in changed[:2]]
            effect = T(f" Mit {_fmt(value)} % ergäbe {' und '.join(shown)}.", f" With {_fmt(value)} %, {' and '.join(shown)}.")
        like = ", ".join(f"{address(p)} {_fmt(round(ctx.sheet.inputs[p] * 100, 4))} %" for p in siblings[:2])
        msg = T(f"{address(pos)} ({name}) steht als {_fmt(value)} da, die Sätze daneben als Prozent ({like}). Excel "
                f"rechnet damit mit {_fmt(value * 100)} %, gemeint sind wohl {_fmt(value)} %.{effect}",
                f"{address(pos)} ({name}) is entered as {_fmt(value)} while the rates next to it are percentages ({like}). "
                f"So Excel calculates with {_fmt(value * 100)} %; {_fmt(value)} % is probably meant.{effect}")
        level = "fehler" if readers else "pruefen"
        found.append(Finding(f"ganzzahl:{address(pos)}", msg, level,
                             T(f"Schreib in {address(pos)} {_fmt(value)} % (mit Prozentzeichen).",
                               f"In {address(pos)}, enter {_fmt(value)} % (with the percent sign)."), fix=fix))
    return found[:MAX_EACH]


# --- Monate im ersten Jahr ab einem Startmonat ---

_START_WORD = _rx(r"inbetriebnahme|in betrieb|anschaffung|zugang|erwerb|\bkauf|aktivierung|beginn|\bstart|eintritt|"
                  r"einstellung|eroffnung|acquisition|purchase|commission|\bbegin|\bhire|mise en service|acquisto|inizio")
_ELAPSED = _rx(r"\bvor\b|verstrichen|vergangen|bereits|bisher|elapsed|before|prior|avant|prima")
_DEPRECIATION = _rx(r"\bafa\b|abschreib|depreciat|amortis|ammortament|pro rata")
_MONTH_NAMES = ("Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober",
                "November", "Dezember")
_MONTH_NAMES_EN = ("January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
                   "November", "December")


def _start_month(ctx: _Ctx) -> list[Finding]:
    """=12-B11 mit B11 „Inbetriebnahme (Monat)“ = 7: Von Juli bis Dezember sind es 6 Monate, nicht 5. Der Startmonat
    zählt mit (pro rata temporis, § 7 Abs. 1 EStG: ab dem Monat der Anschaffung), also 13 minus Monat."""
    found = []
    for pos, node, text in _formulas(ctx):
        parents = {}
        for n in _walk(node):
            for a in n.args or ():
                if isinstance(a, formel.Node):
                    parents[id(a)] = n
        for n in _walk(node):
            if not (n.kind == "op" and n.value == "-" and len(n.args) == 2 and n.args[0].kind == "num"
                    and n.args[0].value == 12 and n.args[1].kind == "ref"):
                continue
            up = parents.get(id(n))
            if up is not None and up.kind == "op" and up.value in "+-" and any(
                    a is not n and a.kind == "num" for a in up.args):
                continue  # =12-B11+1 zählt den Startmonat schon mit
            ref = n.args[1].value
            cell = ctx.sheet.cells.get(ref)
            month = ctx.value(ref)
            if cell is None or cell.formula or not _number(month) or month != int(month) or not 1 <= month <= 12:
                continue
            name = _name(ctx, ref)
            unit = ctx.sheet.cells.get((ref[0], ref[1] + 1))
            unit_text = str(unit.value) if unit is not None and not unit.formula and unit.value else ""
            if not _search(_START_WORD, name) or _search(_ELAPSED, name) or \
                    not (_search(_MONTHS, name) or _search(_MONTHS, unit_text)):
                continue
            new = re.sub(r"(?<![\d.,A-Z$])12(\s*-\s*\$?" + column_letters(ref[1]) + r"\$?" + str(ref[0]) + r")(?!\d)",
                         r"13\1", text, count=1)
            if new == text:
                continue
            changes = {pos: new}
            bad = _differs(ctx, changes)
            if not bad:
                continue
            m = int(month)
            around = " ".join((_name(ctx, pos), _col_head(ctx, pos[1], pos[0]), name))
            tax_rule = _search(_DEPRECIATION, around)
            right, now = _shown(ctx, pos, bad[2]), _shown(ctx, pos, bad[1])
            msg = T(f"{address(pos)} ({_name(ctx, pos)}) rechnet 12 − {address(ref)} ({name}, {m}) = {12 - m} Monate. "
                    f"Der Startmonat zählt mit: {_MONTH_NAMES[m - 1]} bis Dezember sind {13 - m} Monate"
                    + (" (pro rata temporis, nach § 7 EStG ab dem Monat der Anschaffung)" if tax_rule
                       else "") + f". Richtig ist 13 − {address(ref)}. So wäre "
                    f"{address(pos)} {right} statt {now}.",
                    f"{address(pos)} ({_name(ctx, pos)}) calculates 12 − {address(ref)} ({name}, {m}) = {12 - m} "
                    f"month{'' if 12 - m == 1 else 's'}. The start month counts too: {_MONTH_NAMES_EN[m - 1]} to "
                    f"December is {13 - m} month{'' if 13 - m == 1 else 's'}"
                    + (" (pro rata temporis; under § 7 EStG from the month of acquisition)" if tax_rule else "")
                    + f". Correct is 13 − {address(ref)}. Then {address(pos)} would be {right} instead of {now}.")
            item = _finding("startmonat", ctx, pos, msg, "fehler", changes)
            if item:
                found.append(item)
            break
    return found[:MAX_EACH]


CHECKS = (_effects, _bridge, _negative_base, _signs, _balance, _salvage, _book_value, _rate_period, _mape, _npv_zero,
          _share_base, _tax_added, _gordon, _tv_discount, _opening, _qualifier, _outflows, _inflows, _whole_rates, _start_month)
# Was eine dieser Prüfungen zu einer Zelle sagt, ersetzt den allgemeineren Hinweis zur selben Zelle.
KEYS = ("effekte", "bruecke", "negbasis", "vorzeichen", "ausgleich", "restwert", "buchwert", "zinsperiode", "mape",
        "nbw0", "basis", "steuern", "gordon", "tvfaktor", "anfang", "annahme", "auszahlung", "barwert", "ganzzahl", "startmonat")


def checks(sheet: Sheet, rows: list[int]) -> list[Finding]:
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
