"""Herleitung einer Zahl: woraus eine Zelle entsteht, Stufe für Stufe bis zu den Eingaben, und wohin sie fließt.

Für das Seitenfenster in Excel. Man markiert eine Zelle und sieht ihre Formel (auch in Worten), den Wert aus
Excel und ob er nachgerechnet stimmt. Darunter stehen die Zellen, aus denen er kommt, auch auf anderen Blättern,
bei SVERWEIS, XVERWEIS und INDEX auch die Zelle, die gefunden wurde. Große Bereiche werden zusammengefasst.
Alles ohne KI, mit dem eigenen Formelrechner (formel.py) und den Werten aus Excel.
"""

from __future__ import annotations

from itertools import islice

from . import auswertung, formel
from .analyse import Sheet as Calc
from .analyse import _close, _fmt, _german, address
from .formel import Err, Node
from .sprache import T
from .workbook import layout
from .xlsx import Sheet, _date_text, number_text

Pos = tuple[int, int]

MAX_DEPTH = 6
MAX_NODES = 80
# Bereiche bis zu so vielen Zellen zeigt die Herleitung einzeln, größere nur zusammengefasst.
SMALL_RANGE = 8
MAX_USES = 8


class Book:
    """Die Blätter der Mappe, je Blatt einmal mit geparsten Formeln (analyse.Sheet), erst wenn gebraucht."""

    def __init__(self, cells_of: dict[str, dict]) -> None:
        self.cells_of = cells_of
        self._names = {name.casefold(): name for name in cells_of}
        self._calcs: dict[str, Calc] = {}
        self._evaluators: dict[str, formel.Evaluator] = {}
        self._suspect: dict[str, set[Pos]] = {}

    def name(self, sheet: str) -> str | None:
        return self._names.get(sheet.casefold())

    def calc(self, sheet: str) -> Calc:
        name = self.name(sheet)
        if name not in self._calcs:
            cells = self.cells_of[name]
            self._calcs[name] = Calc(cells, layout(Sheet(name, cells)), name, self.cells_of)
        return self._calcs[name]

    def suspect(self, sheet: str) -> set[Pos]:
        """Formeln, die nachgerechnet nicht stimmen, und alle, die auf diesem Blatt davon abhängen."""
        name = self.name(sheet)
        if name not in self._suspect:
            calc = self.calc(name)
            todo = [p for p, node in calc.nodes.items() if node is not None and p not in calc.unknown
                    and _check(calc.shown.get(p), self.run(name, node)[0])[0] == "anders"]
            found = set(todo)
            while todo:
                for after in calc.dependents.get(todo.pop(), ()):
                    if after not in found:
                        found.add(after)
                        todo.append(after)
            self._suspect[name] = found
        return self._suspect[name]

    def run(self, sheet: str, node: Node) -> tuple[object, list]:
        """Die Formel mit den Werten aus Excel nachrechnen, und welche Zellen Nachschlagen geliefert hat."""
        name = self.name(sheet)
        if name not in self._evaluators:
            self._evaluators[name] = self.calc(name).evaluator()[0]
        evaluator = self._evaluators[name]
        evaluator.hits = []
        try:
            return evaluator.run(node), evaluator.hits
        except formel.Unknown:
            return None, []
        finally:
            evaluator.hits = None


def explain(cells_of: dict[str, dict], sheet: str, cell: str) -> dict:
    """Die Herleitung von cell („C4“) auf dem Blatt sheet, als Baum für das Seitenfenster."""
    book = Book(cells_of)
    name = book.name(sheet)
    try:
        pos = formel.position(cell.replace("$", "").upper())
    except (AttributeError, ValueError):
        return {"state": "none"}
    if name is None:
        return {"state": "none"}
    calc = book.calc(name)
    if pos not in calc.cells:
        return {"state": "empty", "sheet": name, "cell": address(pos)}
    budget = [MAX_NODES]
    tree = _node(book, name, pos, 0, set(), budget)
    if tree.get("kind") == "formel":
        # Welche Zeilen der Tabelle die Formel nimmt, in einem Satz.
        tree["rows"] = auswertung.in_words(calc, pos)
    uses = _uses(book, name, pos)
    tree["uses"], tree["uses_more"] = uses[:MAX_USES], len(uses) > MAX_USES
    return {"state": "ok", "sheet": name, "cell": address(pos), "tree": tree, "cut": budget[0] <= 0,
            "problems": _problems(tree, name)}


def _problems(tree: dict, sheet: str) -> list[str]:
    """Die Zellen im Baum, die nachgerechnet etwas anderes ergeben („Tarife!B4“ auf anderen Blättern)."""
    found: list[str] = []

    def walk(node: dict) -> None:
        if node.get("check") == "anders":
            where = node["cell"] if node["sheet"] == sheet else f"{node['sheet']}!{node['cell']}"
            if where not in found:
                found.append(where)
        for child in node.get("children", []):
            walk(child)

    walk(tree)
    return found


def _node(book: Book, sheet: str, pos: Pos, depth: int, seen: set, budget: list[int], note: str = "") -> dict:
    budget[0] -= 1
    calc = book.calc(sheet)
    cell = calc.cells.get(pos)
    item = {"sheet": sheet, "cell": address(pos), "label": calc.label(pos), "value": cell.value if cell else "", "note": note}
    if cell is None:
        item.update(kind="leer", value=T("leer", "empty"))
        return item
    node = calc.nodes.get(pos)
    if cell.formula is None:
        item["kind"] = "eingabe"
        return item
    item.update(kind="formel", formula=f"={cell.typed or cell.formula}", words=calc.readable(pos))
    if node is None or pos in calc.unknown:
        item["check"] = "unbekannt"
        return item
    ours, found = book.run(sheet, node)
    item["check"], item["ours"] = _check(calc.shown.get(pos), ours)
    if item["check"] == "anders" and isinstance(ours, float):
        item["ours"] = _like(cell.value, ours)
    fixed = sorted(set(formel.literals(node)) - {0.0, 1.0})
    if fixed:
        item["fixed"] = [number_text(v) for v in fixed[:5]]
    key = (sheet.casefold(), pos)
    if key in seen:
        item["again"] = True
        return item
    seen.add(key)
    if depth >= MAX_DEPTH:
        item["more"] = len(_parts(node, sheet))
        return item
    children = []
    hits = {(book.name(s) if s else sheet, p): fn for fn, s, p in found}
    for part in _parts(node, sheet):
        if budget[0] <= 0:
            item["more"] = item.get("more", 0) + 1
            continue
        other, value = part
        other_name = book.name(other)
        if other_name is None:
            children.append({"sheet": other, "cell": _where(value), "label": f"{other}!{_where(value)}", "kind": "fremd",
                             "value": T("andere Mappe oder fehlendes Blatt", "other workbook or missing sheet")})
            continue
        if isinstance(value[0], int) and len(value) == 2:
            note = _found_by(hits.pop((other_name, value))) if (other_name, value) in hits else ""
            children.append(_node(book, other_name, value, depth + 1, seen, budget, note))
        else:
            children.append(_area(book, other_name, value, depth, seen, budget, hits))
    for (other, p), fn in hits.items():
        if budget[0] > 0 and other is not None:
            children.append(_node(book, other, p, depth + 1, seen, budget, _found_by(fn)))
    item["children"] = children
    return item


def _area(book: Book, sheet: str, area: tuple, depth: int, seen: set, budget: list[int], hits: dict) -> dict:
    """Ein Bereich: eine kurze Zeile oder Spalte einzeln, sonst zusammengefasst. Aus einer Tabelle (SVERWEIS)
    nur die gefundene Zelle, aus großen Bereichen nur Formeln, die nicht stimmen."""
    budget[0] -= 1
    calc = book.calc(sheet)
    rows, cols = calc.extent(sheet)
    r1, c1, r2, c2 = area
    r2, c2 = min(r2, max(rows, r1)), min(c2, max(cols, c1))
    inside = sorted(p for p in calc.cells if r1 <= p[0] <= r2 and c1 <= p[1] <= c2)
    where = _where((r1, c1, r2, c2))
    line = r1 == r2 or c1 == c2
    header = calc.plan.headers.get(c1) if c1 == c2 and calc.plan.in_table(r1, c1) else None
    label = f"{where} ({header})" if header else where
    if line:
        numbers = [v for p in inside if isinstance(v := calc.inputs.get(p), float) and not isinstance(v, bool)]
        texts = sum(1 for p in inside if isinstance(calc.inputs.get(p), str))
        summary = [T(f"{len(numbers)} {'Zahl' if len(numbers) == 1 else 'Zahlen'}",
                     f"{len(numbers)} {'number' if len(numbers) == 1 else 'numbers'}")]
        if numbers:
            summary.append(T(f"Summe {_fmt(sum(numbers))}", f"total {_fmt(sum(numbers))}"))
        if texts:
            summary.append(T(f"{texts} mit Text", f"{texts} with text"))
        empty = (r2 - r1 + 1) * (c2 - c1 + 1) - len(inside)
        if empty and (r2 - r1 + 1) * (c2 - c1 + 1) <= 10_000:
            summary.append(T(f"{empty} leer", f"{empty} empty"))
    else:
        summary = [T(f"{r2 - r1 + 1} Zeilen, {c2 - c1 + 1} Spalten",
                     f"{r2 - r1 + 1} {'row' if r2 == r1 else 'rows'}, {c2 - c1 + 1} {'column' if c2 == c1 else 'columns'}")]
    item = {"sheet": sheet, "cell": where, "label": label, "kind": "bereich", "value": ", ".join(summary), "note": ""}
    formulas = [p for p in inside if calc.cells[p].formula]
    found = [p for p in inside if (sheet, p) in hits]
    if line and len(inside) <= SMALL_RANGE:
        shown = inside
    else:
        # Was nicht stimmt oder auf etwas aufbaut, das nicht stimmt: dort lohnt es sich hineinzuschauen.
        suspect = book.suspect(sheet) if formulas else set()
        shown = found + list(islice((p for p in formulas if p not in found and p in suspect), 4))
        if formulas:
            item["value"] += T(f", {len(formulas)} davon Formeln", f", {len(formulas)} of them formulas")
    children = []
    for p in shown:
        if budget[0] <= 0:
            item["more"] = item.get("more", 0) + 1
            continue
        note = _found_by(hits.pop((sheet, p))) if (sheet, p) in hits else ""
        children.append(_node(book, sheet, p, depth + 1, seen, budget, note))
    item["children"] = children
    return item


def _found_by(function: str) -> str:
    """„gefunden von SVERWEIS“: die Zelle, die eine Nachschlagefunktion geliefert hat."""
    return T(f"gefunden von {_german(function)}", f"found by {function}")


def _like(shown: str, value: float) -> str:
    """Unser Wert so geschrieben wie der in Excel: „30 %“ neben „20 %“, ein Datum neben einem Datum."""
    if shown.endswith("%"):
        return f"{number_text(round(value * 100, 10))} %"
    if formel.serial_of(shown) is not None:
        return _date_text(value, False)
    return _fmt(value)


def _check(shown, ours) -> tuple[str, str]:
    """„ok“, „anders“ (mit unserem Wert) oder „unbekannt“, wie Nachrechnen es zählt: Zahl gegen Zahl."""
    if isinstance(shown, float) and isinstance(ours, float) and not isinstance(ours, bool) and not isinstance(shown, bool):
        return ("ok", "") if _close(ours, shown) else ("anders", _fmt(ours))
    if isinstance(ours, Err) and isinstance(shown, Err):
        return ("ok", "") if ours.code == shown.code else ("anders", ours.code)
    if isinstance(ours, str) and isinstance(shown, str):
        return ("ok", "") if ours == shown else ("anders", ours)
    if ours is None and shown is None:
        return "ok", ""
    if type(ours) is type(shown) and ours == shown:
        return "ok", ""
    return "unbekannt", ""


def _parts(node: Node, sheet: str) -> list[tuple[str, tuple]]:
    """Die Zellen und Bereiche einer Formel in der Reihenfolge, in der sie stehen, jede einmal."""
    found: list[tuple[str, tuple]] = []

    def walk(n: Node) -> None:
        if n.kind in ("ref", "range"):
            part = (sheet, n.value)
        elif n.kind in ("xref", "xrange"):
            part = (n.value[0], n.value[1])
        else:
            part = None
        if part is not None and (part[0].casefold(), part[1]) not in {(s.casefold(), v) for s, v in found}:
            found.append(part)
        for child in n.args:
            walk(child)

    walk(node)
    return found


def _where(value: tuple) -> str:
    if len(value) == 2:
        return address(value)
    r1, c1, r2, c2 = value
    if r2 >= formel.MAX_ROW and r1 == 1:
        return f"{address((1, c1))[:-1]}:{address((1, c2))[:-1]}"
    return f"{address((r1, c1))}:{address((r2, c2))}" if (r1, c1) != (r2, c2) else address((r1, c1))


def _uses(book: Book, sheet: str, pos: Pos) -> list[dict]:
    """Wohin die Zahl fließt: Formeln, die diese Zelle lesen, auf diesem und den anderen Blättern."""
    found = []
    mine = sheet.casefold()
    for name in book.cells_of:
        cells = book.cells_of[name]
        if name.casefold() != mine and not any(c.formula and "!" in c.formula for c in cells.values()):
            continue
        calc = book.calc(name)
        for p, node in sorted(calc.nodes.items()):
            if node is None or (name.casefold() == mine and p == pos):
                continue
            if any(s.casefold() == mine and _covers(v, pos) for s, v in _parts(node, name)):
                found.append({"sheet": name, "cell": address(p), "label": calc.label(p) if name.casefold() == mine else f"{name}!{calc.label(p)}"})
                if len(found) > MAX_USES:
                    return found
    return found


def _covers(value: tuple, pos: Pos) -> bool:
    if len(value) == 2:
        return value == pos
    r1, c1, r2, c2 = value
    return r1 <= pos[0] <= r2 and c1 <= pos[1] <= c2
