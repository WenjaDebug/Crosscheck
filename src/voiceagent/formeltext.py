"""Eine Formel, die als Text in der Zelle steht: War die Zelle als Text formatiert (oder beginnt die Eingabe mit einem
Hochkomma), bleibt =SUMME(B2:B4) stehen, wie getippt. Excel zeigt die Formel statt des Ergebnisses, SUMME und
Verweise lesen dort keine Zahl, und das fällt erst auf, wenn man genau hinsieht.

Gemeldet wird nur, was Excel als Formel lesen würde: bekannte Funktionen (deutsch oder englisch), Zellbezüge oder
Namen der Mappe. Eine Spalte „Formel“ oder „Erklärung“, die Formeln zum Nachlesen aufschreibt, bleibt still, ebenso
der Text neben einer echten Formel, die genau so rechnet. Einsetzen stellt die Zelle auf „Standard“ und schreibt die
Formel als Formel; Rückgängig macht wieder Text daraus."""

from __future__ import annotations

import re

from . import einsetzen, formel
from .analyse import Finding, Sheet, _checked, address, changed_by, more_formulas, said
from .sprache import T
from .workbook import FUNCTIONS

Pos = tuple[int, int]
MAX_FOUND = 20
MAX_LISTED = 3
# Spalten, die Formeln zum Nachlesen aufschreiben.
_NOTES = re.compile(r"formel|formula|berechnung|rechenweg|erklärung|erlaeuterung|erläuterung|beschreibung|hinweis|"
                    r"notiz|kommentar|bemerkung|anleitung|beispiel|syntax|note|comment|description|example", re.I)
_GERMAN_NAMES = set(FUNCTIONS.values())


def checks(sheet: Sheet) -> list[Finding]:
    found = []
    for pos, cell in sorted(sheet.cells.items(), key=lambda item: (item[0][1], item[0][0])):
        if cell.formula is not None or not cell.value.lstrip().startswith("="):
            continue
        typed = cell.value.strip()
        formula = _formula(sheet, typed[1:])
        if formula is None or _documented(sheet, pos, formula):
            continue
        item = _hint(sheet, pos, typed, formula)
        if item is not None:
            found.append(item)
        if len(found) >= MAX_FOUND:
            break
    return found


def _formula(sheet: Sheet, body: str) -> str | None:
    """Die Formel auf Englisch, wie Excel sie liest, oder None, wenn der Text keine wäre (== Titel ==, =Netto)."""
    body = body.strip()
    if not body:
        return None
    english = einsetzen.english(body)
    for text in dict.fromkeys((body, english)):
        try:
            node = formel.parse(text)
        except formel.Unknown:
            continue
        if _readable(sheet, node):
            return text
    return None


def _readable(sheet: Sheet, node) -> bool:
    """Bekannte Funktionen, Bezüge oder Namen der Mappe, und mindestens ein Bezug oder eine Funktion."""
    names = {n.casefold() for n in getattr(sheet, "names", {}) or {}}
    seen = False
    todo = [node]
    while todo:
        n = todo.pop()
        if n.kind == "unknown":
            if str(n.value).casefold() not in names:
                return False
            seen = True
        elif n.kind == "call":
            if n.value not in FUNCTIONS:
                return False
            seen = True
        elif n.kind in ("ref", "range", "xref", "xrange", "whole"):
            seen = True
        todo.extend(n.args)
    return seen


def _documented(sheet: Sheet, pos: Pos, formula: str) -> bool:
    """Formeln zum Nachlesen: in einer Spalte „Formel“, oder neben einer echten Formel, die genau so rechnet."""
    plan = sheet.plan
    head = plan.headers.get(pos[1], "") if plan.header_row is not None and pos[0] > plan.header_row else ""
    if head and _NOTES.search(head):
        return True
    if _NOTES.search(str(sheet.shown.get((pos[0], pos[1] - 1), ""))):
        return True
    # Die echte Formel steht in derselben Zeile, der Text erklärt sie.
    wanted = _same(formula)
    row = sheet.memo.setdefault("formeltext.rows", {})
    if not row:
        for (r, _), cell in sheet.cells.items():
            if cell.formula is not None:
                row.setdefault(r, set()).add(_same(cell.formula))
    return wanted in row.get(pos[0], ())


def _same(formula: str) -> str:
    return re.sub(r"\s+|\$", "", formula).upper()


def _hint(sheet: Sheet, pos: Pos, typed: str, formula: str) -> Finding | None:
    from .bereiche import _german_formula

    readers = sheet.dependents.get(pos, ())
    column = [p for p in sheet.cells if p[1] == pos[1] and p != pos and sheet.cells[p].formula is not None]
    level = "fehler" if readers or column else "pruefen"
    checked = _checked(sheet, formula)
    effect = ""
    if checked is not None:
        calc, _ = sheet.evaluator()
        try:
            value = calc.run(formel.parse(formula))
        except (formel.Unknown, ValueError, TypeError, ZeroDivisionError, RecursionError, IndexError, KeyError):
            value = None
        if isinstance(value, float):
            changed = changed_by(sheet, readers, {pos: value}, near=True)
            shown = [T(f"{sheet.label(p)} {said(b)} statt {said(a)}", f"{sheet.label(p)} gives {said(b)} instead of {said(a)}")
                     for p, b, a in changed[:MAX_LISTED]]
            if shown:
                more = more_formulas(len(changed) - MAX_LISTED) if len(changed) > MAX_LISTED else ""
                listed = shown[0] if len(shown) == 1 else ", ".join(shown[:-1]) + T(f" und {shown[-1]}", f" and {shown[-1]}")
                effect = (T(f" Darum ergibt {', '.join(shown)}{more}.", f" As a result, {', '.join(shown)}{more}.")
                          if more else T(f" Darum ergibt {listed}.", f" As a result, {listed}."))
    german = "=" + _german_formula(formula)
    text = T(f"{sheet.label(pos)} zeigt {typed} als Text, statt zu rechnen: Die Zelle ist als Text formatiert, oder vor "
             f"der Formel steht ein Hochkomma. Formeln, die sie lesen, sehen dort keine Zahl.{effect}",
             f"{sheet.label(pos)} shows {typed} as text instead of calculating it: the cell is formatted as text, or "
             f"there is an apostrophe before the formula. Formulas that read it see no number there.{effect}")
    tip = (T(f"Korrigiert in {address(pos)}: {checked}.", f"Corrected in {address(pos)}: {checked}.") if checked else
           T("Stell die Zelle auf das Zahlenformat „Standard“ und tippe die Formel neu ein.",
             "Set the cell to the “General” number format and type the formula again."))
    return Finding(f"formeltext:{address(pos)}", text, level, tip, fix=_fix(sheet, pos, typed, formula, german))


def _fix(sheet: Sheet, pos: Pos, typed: str, formula: str, german: str) -> dict:
    """Einsetzen: Zahlenformat „Standard“, dann die Formel. Vorher steht der Text in der Zelle; das Seitenfenster
    liest ihn als „Formel“ ohne „=“ (Office.js gibt bei Text den Text zurück)."""
    return {"range": address(pos), "cells": {address(pos): "=" + formula}, "was": {address(pos): typed[1:].strip()},
            "german": german, "general": [address(pos)]}
