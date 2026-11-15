"""Eine Folge beim Ursprung: Steht 1,74 statt 174 in Größe (cm), ist der BMI derselben Zeile unmöglich. Das ist ein
Fehler, nicht zwei. Der Hinweis zum BMI fällt für diese Zeile weg, und der Hinweis zur Größe sagt, was der BMI damit
zeigt (er wird zum Fehler, denn der BMI zeigt, dass die Größe nicht stimmen kann).

Ebenso ein Ausreißer, den „unmöglich“ schon nennt: ein Hinweis je Zelle. Nennt „unmöglich“ eine Zahl, die mal 100
passt (scale), bleibt der genauere Hinweis mit der Zahl zum Einsetzen, als Fehler."""

from __future__ import annotations

import re

from .analyse import Finding, Sheet, address, said
from .formel import position
from .sprache import T

Pos = tuple[int, int]
MAX_LISTED = 3
# So weit geht die Suche nach den Eingaben einer Formel zurück.
MAX_VISITS = 200
_CELL = re.compile(r"\b[A-Z]{1,3}[1-9]\d*\b")


def fold(sheet: Sheet, findings: list[Finding]) -> list[Finding]:
    from .auswertung import impossible_hint

    details = sheet.memo.get("auswertung.bad", {})
    by_key = {f.key: f for f in findings}
    # Eingetippte Zahlen, zu denen es schon einen Hinweis gibt: genauer (scale) vor unmöglich vor Ausreißer.
    origins: dict[Pos, Finding] = {}
    for kind in ("scale", "impossible", "outlier"):
        for item in findings:
            if not item.key.startswith(kind + ":"):
                continue
            for pos in _cells_of(item, details):
                if pos not in sheet.nodes:
                    origins.setdefault(pos, item)
    drop: set[str] = set()
    rebuilt: dict[str, Finding | None] = {}
    follows: dict[str, list[Pos]] = {}
    for key, (col, header, values, bad, why) in list(details.items()):
        item = by_key.get(key)
        if item is None:
            continue
        kept = []
        for r in bad:
            pos = (r, col)
            origin = origins.get(pos)
            if origin is not None and origin is not item:
                # Dieselbe Zahl hat einen genaueren Hinweis (mal 100): Der bekommt die Stufe von „unmöglich“.
                _sure(sheet, origin)
                continue
            causes = [origins[p] for p in _inputs(sheet, pos) if p in origins and origins[p] is not item]
            if pos in sheet.nodes and causes:
                for key_of in dict.fromkeys(c.key for c in causes):
                    follows.setdefault(key_of, []).append(pos)
                    _sure(sheet, by_key[key_of])
                continue
            kept.append(r)
        if kept != bad:
            rebuilt[key] = impossible_hint(sheet, col, header, values, kept, why) if kept else None
    # Ein Ausreißer, den „unmöglich“ schon nennt.
    for pos, origin in origins.items():
        if origin.key.startswith(("impossible:", "scale:")):
            other = by_key.get(f"outlier:{address(pos)}")
            if other is not None and other is not origin:
                drop.add(other.key)
    for key, cells in follows.items():
        item = by_key[key]
        # Was eine Korrektur dort ändert, sagt der Hinweis schon („Darum ergibt G10 …“).
        cells = [p for p in cells if address(p) not in item.text]
        if cells:
            shown = [f"{sheet.label(p)} {said(sheet.inputs.get(p))}" for p in cells[:MAX_LISTED]]
            more = T(f" und {len(cells) - MAX_LISTED} weitere", f" and {len(cells) - MAX_LISTED} more") if len(cells) > MAX_LISTED else ""
            text = shown[0] if len(shown) == 1 else ", ".join(shown[:-1]) + T(f" und {shown[-1]}", f" and {shown[-1]}")
            item.text += T(f" Darum ist auch {text}{more} unmöglich.", f" That is also why {text}{more} is impossible.")
    out = []
    for item in findings:
        if item.key in drop:
            continue
        if item.key in rebuilt:
            if rebuilt[item.key] is not None:
                out.append(rebuilt[item.key])
            continue
        out.append(item)
    return out


def _sure(sheet: Sheet, item: Finding) -> None:
    """Ein Fehler: Was die Zahl in einer Formel anrichtet, ist unmöglich. Mal 100 ist dann die Korrektur, nicht nur
    eine Möglichkeit."""
    item.level = "fehler"
    unsure, sure = sheet.memo.get("analyse.scaled", {}).get(item.key, ("", ""))
    if unsure and unsure in item.text:
        item.text = item.text.replace(unsure, sure)


def _cells_of(item: Finding, details: dict) -> list[Pos]:
    if item.key in details:
        col, _, _, bad, _ = details[item.key]
        return [(r, col) for r in bad]
    return [position(item.key.partition(":")[2])]


def _inputs(sheet: Sheet, start: Pos) -> list[Pos]:
    """Die eingetippten Zellen, mit denen eine Formel rechnet, auch über Formeln dazwischen (höchstens MAX_VISITS)."""
    if start not in sheet.nodes:
        return []
    found, seen, todo = [], {start}, [start]
    while todo and len(seen) < MAX_VISITS:
        pos = todo.pop()
        before = sheet.precedents.get(pos)
        if before is None or not isinstance(before, set):
            continue  # ein großer Bereich (SUMME über die Spalte) macht aus einem Wert keine Zeilen-Folge
        for p in before:
            if p in seen:
                continue
            seen.add(p)
            if p in sheet.nodes:
                todo.append(p)
            else:
                found.append(p)
    return found
