"""Excel-Prüfer: die offene Mappe direkt aus Excel lesen und nachrechnen, ohne KI.

Zwei Stufen: die schnellen Prüfungen aus task.py und die tiefe Prüfung der ganzen Rechenkette
aus analyse.py (Formelmuster, Kreisbezüge, Doppelzählung, Nachrechnen, Sensitivität, Daten).

Was sich nachrechnen lässt, rechnet Python sicher (task.py): Fehlerwerte, eine SUMME, die eine Zeile
auslässt oder anders rechnet, gerundete Werte, die zusammen nicht mehr aufgehen, eine feste Zahl statt
einer Formel, abgetippte Zahlen statt Zellbezügen und dieselbe feste Zahl in vielen Formeln.
Nichts davon verlässt den Rechner.
"""

from __future__ import annotations

import contextvars
import hashlib
import json
import re
import threading
import time
from dataclasses import asdict
from pathlib import Path

from . import (addin, ansicht, bericht, bezuege, einlesen, einsetzen, formel, herleitung, inworten, pivotdiagramm,
               sichtbar, spaltenregeln, sprache, vergleich, verbunden, verknuepfungen)
from .bereiche import _german_formula
from .sprache import T
from .analyse import Report, Stopped, analyse, failure, step
from .analyse import _Guard as Guard
from .analyse import _log
from .analyse import stoppable as analyse_stoppable
from .analyse import Sheet as Analysed
from .task import Fact, facts, material
from .workbook import WorkbookTracker, layout, others as other_tables
from .xlsx import Locked, Sheet, area_of, cut_of, read_xlsx

# Ab diesem Gewicht (task.Fact.weight) ist es ein Hinweis, darunter eine Übersicht.
HINT = 30
# So lange prüft die Mappe in Worten noch nicht geprüfte Blätter, um ihre Hinweise zu zählen.
DESCRIBE_BUDGET = 3.0


class _Said(dict):
    """Feste Sätze in beiden Sprachen (deutsch, englisch): [art] und get() geben den Satz in der Sprache der Anfrage."""

    def __getitem__(self, key):
        return T(*super().__getitem__(key))

    def get(self, key, default=None):
        return self[key] if key in self else default


STATES = _Said({
    "off": ("Den Excel-Prüfer gibt es nur am Mac und unter Windows.",
            "The Excel checker is only available on Mac and Windows."),
    "closed": ("Öffne eine Mappe in Excel, dann prüfe ich sie.", "Open a workbook in Excel, then I check it."),
    "connecting": ("Ich verbinde mich mit Excel …", "Connecting to Excel …"),
    "empty": ("Excel ist offen, aber ohne Mappe.", "Excel is open, but without a workbook."),
    "denied": (
        "Ich darf Excel nicht lesen. Erlaube es in den Systemeinstellungen unter Datenschutz und Sicherheit, "
        "Automation: Bei Terminal den Schalter Microsoft Excel einschalten.",
        "I am not allowed to read Excel. Allow it in System Settings under Privacy & Security, "
        "Automation: for Terminal, turn on the Microsoft Excel switch.",
    ),
    "busy": ("Excel ist gerade beschäftigt (vielleicht tippst du noch in einer Zelle).",
             "Excel is busy right now (maybe you are still typing in a cell)."),
    "error": ("Excel lässt sich gerade nicht lesen.", "Excel cannot be read right now."),
})


# Was man meistens tun kann, je Art des Hinweises (task.Fact.key vor dem Doppelpunkt).
TIPS = _Said({
    "error": ("Schau dir die Formel an: Oft fehlt ein Wert, oder es wird durch 0 geteilt.",
              "Look at the formula: often a value is missing, or it divides by 0."),
    "gap": ("Zieh den Bereich der SUMME bis zur letzten Zeile der Tabelle.",
            "Extend the range of the SUM to the last row of the table."),
    "sum": ("Prüfe, ob Zahlen als Text gespeichert sind (grünes Dreieck oben links in der Zelle).",
            "Check whether numbers are stored as text (green triangle in the top left corner of the cell)."),
    "round": ("Rundest du später, prüf auch die Summe der gerundeten Werte.",
              "If you round later, also check the sum of the rounded values."),
    "typed": ("Prüfe, ob dort eine Formel fehlt.", "Check whether a formula is missing there."),
    "constant": ("Schreib die Zahl einmal in eine eigene Zelle und verweise mit $ darauf, dann änderst du sie an einer Stelle.",
                 "Put the number once in a cell of its own and refer to it with $; then you change it in one place."),
    "copied": ("Verweise auf die Zelle, statt die Zahl abzutippen. Dann rechnet es mit, wenn sich der Wert ändert.",
               "Refer to the cell instead of typing the number. Then it follows along when the value changes."),
})


def finding(fact: Fact) -> dict:
    level = "fehler" if fact.weight >= 80 else "pruefen" if fact.weight >= HINT else "info"
    return {"key": fact.key, "text": fact.text, "level": level, "tip": TIPS.get(fact.key.split(":")[0], "") if fact.weight >= HINT else ""}


_ADDRESS = re.compile(r"\b[A-Z]{1,3}[1-9]\d{0,6}(?::[A-Z]{1,3}[1-9]\d{0,6})?\b")


def _cell_of(key: str, text: str) -> str:
    """Die Zelle (oder der Bereich), um die es geht, damit das Add-in sie in Excel markieren kann."""
    tail = key.split(":", 1)[-1]
    if _ADDRESS.fullmatch(tail):
        return tail
    match = _ADDRESS.search(text)
    return match.group(0) if match else ""


def _unreadable(err: ValueError) -> str:
    return str(err) if isinstance(err, Locked) else T(f"{err} Nimm eine .xlsx-Datei, die Excel gespeichert hat.",
                                                       f"{err} Use an .xlsx file saved by Excel.")


def result(book: str, sheet: str, found: list[Fact], report: Report) -> dict:
    """Hinweise, Abhängigkeiten und Übersicht für ein Blatt, wie Seite und Add-in sie zeigen."""
    ordered = sorted(found, key=lambda f: -f.weight)
    # Fehlerwerte erklärt fehlerwerte.py mit ihrer Ursache (in report.findings).
    # Eine Summe, die Zeilen auslässt: Sagt die tiefe Prüfung genau, wo der Bereich falsch ist, und bringt sie die
    # nachgerechnete Korrektur für diese Zelle mit, reicht ihr Hinweis.
    repaired = {item.key.split(":", 1)[1] for item in report.findings
                if item.fix and item.level == "fehler" and item.key.split(":", 1)[1] in (item.fix.get("cells") or {})}
    # Eine einzelne feste Zahl in einer rechnenden Spalte sagt zeilenmodell.typed_down mit Korrektur.
    typed = [item.key.split(":", 1)[1] for item in report.findings if item.key.startswith("typedcell:")]
    hints = [finding(f) for f in ordered if f.weight >= HINT and not f.key.startswith("error:")
             and not (f.key.startswith("gap:") and f.key.split(":", 1)[1] in repaired)
             and not (f.key.startswith("typed:") and any(re.search(rf"(?<![A-Z]){a}(?!\d)", f.text) for a in typed))]
    known = {h["key"].split(":", 1)[-1] for h in hints}
    # Die Summenlücke aus zeilenmodell._column_gaps nur, wenn keine genauere Prüfung dieselbe Zelle korrigiert.
    exact = {item.key.split(":", 1)[1] for item in report.findings if not item.key.startswith("gap:")
             and item.fix and item.level == "fehler" and item.key.split(":", 1)[1] in (item.fix.get("cells") or {})}
    for item in report.findings:
        if item.key.startswith("gap:") and item.key.split(":", 1)[1] in exact:
            continue
        # Was task.py schon für diese Zelle meldet, nicht noch einmal.
        if item.key.startswith("recalc:") and item.key.split(":", 1)[1] in known:
            continue
        hints.append({"key": item.key, "text": item.text, "level": item.level, "tip": item.tip,
                      **({"fix": item.fix} if item.fix else {}), **({"action": item.action} if item.action else {}),
                      **({"cell": item.cell} if item.cell is not None else {}),
                      **({"choices": item.choices} if item.choices else {})})
    order = {"fehler": 0, "pruefen": 1, "info": 2}
    hints.sort(key=lambda h: order[h["level"]])
    for hint in hints:
        if "cell" not in hint:
            hint["cell"] = _cell_of(hint["key"], hint["text"])
        # Für das Seitenfenster: Titel, wo, der Satz und was er gegenüberstellt (der Text bleibt, wie er ist).
        hint.update(ansicht.view(hint))
    if report.partial:
        # Zu Formeln, die mit nicht gelesenen Zeilen rechnen, weiß der Prüfer nichts Sicheres.
        hints = [h for h in hints if not _touches(h["cell"], report.partial)]
    done = {
        "state": "ok",
        "book": book,
        "sheet": sheet,
        "text": tally(hints),
        "hints": hints,
        "drivers": [asdict(d) for d in report.drivers],
        "stats": report.stats,
        "overview": [finding(f) for f in ordered if f.weight < HINT],
    }
    if report.cuts:
        done["cut"] = cut_note(sheet, report)
    if report.failed:
        done["failed"] = failed_note(report.failed)
    return done


MAX_ID_CELLS = 200


def hint_id(hint: dict, cells: dict) -> str:
    """Woran das Fenster einen Hinweis wiedererkennt, den jemand mit „Passt so“ ausgeblendet hat: Art, Satz und was
    in seinen Zellen steht. Ändert sich dort etwas (eine andere Formel, ein anderer Wert), kommt der Hinweis wieder."""
    seen = []
    for part in (hint.get("cell") or "").split(","):
        try:
            r1, c1, r2, c2 = area_of(part.strip())
        except (ValueError, KeyError, TypeError):
            continue
        for r in range(r1, min(r2, r1 + MAX_ID_CELLS) + 1):
            for c in range(c1, min(c2, c1 + MAX_ID_CELLS) + 1):
                cell = cells.get((r, c))
                seen.append("" if cell is None else f"{cell.formula or ''}\t{cell.value}")
        if len(seen) >= MAX_ID_CELLS:
            break
    said = "\n".join([hint["key"], hint["text"], *seen])
    return hashlib.sha1(said.encode("utf-8")).hexdigest()[:12]


def passed(done: dict | None, ids) -> dict | None:
    """Das Ergebnis eines Blatts ohne die Hinweise, die mit „Passt so“ ausgeblendet sind; passt: wie viele."""
    if done is None or not ids or not done.get("hints"):
        return done
    kept = [h for h in done["hints"] if h.get("id") not in ids]
    hidden = len(done["hints"]) - len(kept)
    if not hidden:
        return done
    return {**done, "hints": kept, "text": tally(kept) if kept else T("Mir fällt nichts weiter auf.", "Nothing else stands out to me."), "passt": hidden}


def tally(hints: list[dict]) -> str:
    """„3 Fehler und 1 Stelle zum Prüfen.“: was auf dem Blatt ist, nach Stufe, damit ein Fehler nicht wie eine
    bloße Frage klingt."""
    if not hints:
        return T("Mir fällt nichts auf.", "Nothing stands out to me.")
    count = [h["level"] for h in hints].count
    errors, review, notes = count("fehler"), count("pruefen"), count("info")
    parts = [text for n, text in (
        (errors, T(f"{errors} Fehler", f"{errors} {'error' if errors == 1 else 'errors'}")),
        (review, T(f"{review} {'Stelle' if review == 1 else 'Stellen'} zum Prüfen",
                   f"{review} {'spot' if review == 1 else 'spots'} to review")),
        (notes, T(f"{notes} {'Hinweis' if notes == 1 else 'Hinweise'}", f"{notes} {'note' if notes == 1 else 'notes'}"))) if n]
    return (T(" und ", " and ").join([", ".join(parts[:-1]), parts[-1]]) if len(parts) > 1 else parts[0]) + "."


def _touches(cells: str, partial: set) -> bool:
    """Liegt eine der Zellen („C5,C14:C16“) unter partial? Große Bereiche nur an ihren Ecken und Formeln darin."""
    for part in (cells or "").split(","):
        area = area_of(part.strip())
        if area is None:
            continue
        r1, c1, r2, c2 = area
        if (r2 - r1 + 1) * (c2 - c1 + 1) <= 10_000:
            if any((r, c) in partial for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)):
                return True
        elif any(r1 <= r <= r2 and c1 <= c <= c2 for r, c in partial):
            return True
    return False


def cut_note(sheet: str, report: Report) -> str:
    """Was der Prüfer bei einem großen Blatt nicht gelesen hat und was daraus folgt, in einem Satz für das Fenster."""
    parts = []
    own = report.cuts.get(None)
    if own:
        parts.append(T(f"Das Blatt ist sehr groß: Ich prüfe die Zeilen bis {_number(own)}.",
                       f"The sheet is very large: I check the rows up to {_number(own)}."))
    elif own is not None:
        parts.append(T("Die Mappe ist sehr groß: Von diesem Blatt lese ich nichts mehr.",
                       "The workbook is very large: I read nothing more from this sheet."))
    others = sorted((name, row) for name, row in report.cuts.items() if name is not None)
    for name, row in others[:3]:
        parts.append(T(f"„{name}“ ist sehr groß: Davon lese ich nur die Zeilen bis {_number(row)}.",
                       f"“{name}” is very large: I read only its rows up to {_number(row)}.") if row else
                     T(f"Von „{name}“ lese ich nichts mehr, die Mappe ist sehr groß.",
                       f"I read nothing more from “{name}”; the workbook is very large."))
    if report.partial:
        count = len(report.partial)
        parts.append(T(f"{'Eine Formel hier rechnet' if count == 1 else f'{_number(count)} Formeln hier rechnen'} mit Zeilen "
                       "weiter unten, die rechne ich nicht nach.",
                       f"{'One formula here calculates' if count == 1 else f'{_number(count)} formulas here calculate'} with "
                       "rows further down; I do not recalculate those."))
    return " ".join(parts)


def _number(value: int) -> str:
    return f"{value:,}".replace(",", "\u202f") if value >= 10_000 else str(value)


# So lange darf das Add-in je Prüfung auf die übrigen Blätter warten. Was dann fehlt, kommt beim nächsten Mal.
OTHERS_BUDGET = 3.0
# Größere übrige Blätter prüft er nicht nebenbei, sondern im Hintergrund, wenn so lange Ruhe ist (Sekunden).
OTHERS_CELLS = 30_000
IDLE_WAIT = 1.5


def sheet_result(book: str, name: str, cells_of: dict[str, dict], filters: list | None = None, parts: Sheet | None = None,
                 tables: dict | None = None, live: bool = False, charts: bool = True, rules: list[dict] | None = None) -> dict:
    """Hinweise für ein Blatt der Mappe, mit den Blättern, auf die seine Formeln verweisen.

    parts: das Blatt mit seinen Pivot-Tabellen und Diagrammen (deren Daten auch auf anderen Blättern stehen können),
    tables: die Excel-Tabellen der Mappe, live: aus dem Seitenfenster (dann mit Knöpfen). charts=False: die Zahlen in
    Pivot-Tabellen zählen nicht als Eingaben, sie selbst werden aber nicht geprüft. rules: die eigenen Regeln je
    Spalte (spaltenregeln.py)."""
    with formel.remembered():
        try:
            return _sheet_result(book, name, cells_of, filters, parts, tables, live, charts, rules)
        except Stopped:
            raise
        except Exception as err:  # noqa: BLE001 - ein Blatt, das nicht geht, nimmt die übrigen nicht mit
            _log.error("Blatt nicht geprüft: %s", failure(err))
            return {"state": "ok", "book": book, "sheet": name, "hints": [], "drivers": [], "stats": {}, "overview": [],
                    "text": T("Dieses Blatt konnte ich nicht prüfen.", "I could not check this sheet."),
                    "failed": T("An diesem Blatt ist die Prüfung gescheitert, die Mappe ist unverändert. ",
                                "The check failed on this sheet; the workbook is unchanged. ") + report_hint()}


def report_hint() -> str:
    """Wie man den Grund weitergeben kann, ohne die Daten der Mappe."""
    return T("Den Grund (ohne deine Daten) enthält der Fehlerbericht: Startseite von Gegenprobe → Mehr → Zustand.",
             "The error report has the reason (without your data): Gegenprobe start page → More → Status.")


# Die Teile der Prüfung (analyse._Guard), wie failed_note sie auf Englisch nennt.
_STEPS = {
    "Als 0 gezeigt": "shown as 0", "Auswertung der Daten": "data analysis", "Blöcke einer Spalte": "blocks in a column",
    "Daten": "data", "Doppelte Schlüssel": "duplicate keys", "Einheiten je Zeile": "units per row",
    "Einsetzen": "fixes", "Fehlerwerte": "error values", "Finanzen": "finance", "Finanzplan": "financial plan",
    "Folgen beim Ursprung": "effects at their origin", "Formelmuster": "formula patterns",
    "Formeln als Text": "formulas as text", "Geschäftsjahr und Monatsende": "fiscal year and month end",
    "Kreisbezüge": "circular references", "Listen": "lists", "Modell": "model", "Nachrechnen": "recalculation",
    "Pivot-Tabellen und Diagramme": "pivot tables and charts", "Rechenkette": "calculation chain",
    "SVERWEIS und Nachschlagen": "VLOOKUP and lookups", "Sensitivität": "sensitivity",
    "Sensitivitätstabellen": "sensitivity tables", "Spalte zum Namen": "column for the name", "Statistik": "statistics",
    "Statistiktests": "statistical tests", "Summen und ihre Teile": "totals and their parts",
    "Verdeckte Zahlen": "concealed numbers", "Verknüpfungen": "external links", "Verrutschte Sätze": "shifted rates",
    "Zeiträume": "time periods", "Zwischensummen": "subtotals", "ausgeblendete Zeilen": "hidden rows",
    "eigene Regeln": "your own rules", "leere Zellen in Formeln": "empty cells in formulas",
    "schnelle Prüfungen": "quick checks", "schwer prüfbare Formeln": "formulas that are hard to check",
    "verbundene Zellen": "merged cells", "weitere Tabellen": "further tables",
}


def failed_note(failed: list[str]) -> str:
    """Welche Teile der Prüfung an diesem Blatt gescheitert sind, in einem Satz für das Fenster."""
    if not failed:
        return ""
    names = [T(f, _STEPS.get(f, f)) for f in failed]
    parts = names[0] if len(names) == 1 else ", ".join(names[:-1]) + T(f" und {names[-1]}", f" and {names[-1]}")
    return T(f"{'Ein Teil der Prüfung' if len(failed) == 1 else 'Einige Teile der Prüfung'} ({parts}) "
             f"{'ging' if len(failed) == 1 else 'gingen'} bei diesem Blatt nicht. Die Hinweise oben gelten, es kann aber "
             f"einer fehlen. {report_hint()}",
             f"{'One part of the check' if len(failed) == 1 else 'Some parts of the check'} ({parts}) did not work on "
             f"this sheet. The findings above hold, but one may be missing. {report_hint()}")


def _sheet_result(book: str, name: str, cells_of: dict[str, dict], filters: list | None, parts: Sheet | None,
                  tables: dict | None, live: bool, charts: bool, rules: list[dict] | None) -> dict:
    cells = cells_of[name]
    uses = _uses(cells, cells_of, name)
    whole = Sheet(name, cells, pivots=list(parts.pivots) if parts is not None else [])
    plan = layout(whole)
    others = {n: cells_of[n] for n in uses}
    report = analyse(cells, plan, exact=True, name=name, book=others, filtered=filters, rules=rules)
    guard = Guard(report)
    found = guard("schnelle Prüfungen", lambda: facts(cells, plan), [])
    step()
    # Mehrere Tabellen untereinander (ein Modell mit Übergangsmatrix, Verteilung und Verlust je Jahr): Jede weitere
    # wird mit ihrer eigenen Kopfzeile geprüft. Von ihr zählt nur, was in ihr liegt und noch nicht gesagt ist.
    extras = guard("weitere Tabellen", lambda: other_tables(whole, plan), [])
    for extra in extras:
        more = analyse(cells, extra, exact=True, name=name, book=others, filtered=filters, rules=rules)
        report.failed += [f for f in more.failed if f not in report.failed]
        inside = [f for f in more.findings if _in_table(f.cell if f.cell is not None else _cell_of(f.key, f.text), extra)]
        # Was in dieser Tabelle liegt, sagt sie mit ihren eigenen Namen („B32 (Neu-Ausfallanteil Jahr 4)“).
        mine = {f.key for f in inside}
        report.findings = [f for f in report.findings if f.key not in mine] + inside
        said = {f.key for f in report.findings}
        more_facts = [f for f in guard("schnelle Prüfungen", lambda extra=extra: facts(cells, extra), [])
                      if f.key not in said and _in_table(_cell_of(f.key, f.text), extra)]
        mine = {f.key for f in more_facts}
        found = [f for f in found if f.key not in mine] + more_facts
        step()
    if parts is not None and parts.merged:
        report.findings += guard("verbundene Zellen", lambda: verbunden.checks(cells, [plan, *extras], parts.merged, name,
                                                                                others, parts.filters), [])
    # Werte aus anderen Dateien: nur der zuletzt geholte Stand ist prüfbar.
    report.findings += guard("Verknüpfungen", lambda: verknuepfungen.checks(cells, plan, name, {n: c for n, c in cells_of.items() if n != name},
                                                                          parts.links if parts is not None else None), [])
    # Ausgeblendete und weggefilterte Zeilen in Summen: was man sieht und was gezählt wird.
    seen, covered = guard("ausgeblendete Zeilen", lambda: sichtbar.checks(cells, plan, name, others), ([], set()))
    if seen or covered:
        report.findings = [f for f in report.findings if not (f.key.startswith("filter:") and
                                                              f.key.partition(":")[2] in covered)] + seen
    if parts is not None and charts:
        report.findings += guard("Pivot-Tabellen und Diagramme",
                                 lambda: pivotdiagramm.checks(name, parts, cells_of, tables, live), [])
    # Korrekturen und Formeln im Text so, wie die Mappe schreibt: [@Menge] und MwSt statt C5 und $I$1.
    done = bezuege.restyle(result(book, name, found, report), cells, parts.refs if parts is not None else None, name,
                           _german_formula)
    for hint in done["hints"]:
        hint["id"] = hint_id(hint, cells)
        for fix in [hint.get("fix"), *((c or {}).get("fix") for c in hint.get("choices") or [])]:
            _english(fix)
    return done


def _english(fix) -> None:
    """Eine Korrektur, die eine Formel ist, auch so, wie englisches Excel sie zeigt (=SUM(A1,B1) statt
    =SUMME(A1;B1)): Das Seitenfenster zeigt die, die zur Sprache von Excel passt. Eingesetzt wird ohnehin
    fix["cells"], das Excel in jeder Sprache versteht."""
    if isinstance(fix, dict) and str(fix.get("german") or "").startswith("="):
        fix["english"] = einsetzen.english(fix["german"])


def _in_table(cells: str, plan) -> bool:
    """Liegt die erste Zelle („C5“, „C5:C9“, „C5,C14“) in der Tabelle von plan, Kopfzeile ausgenommen?"""
    first = (cells or "").split(",")[0].split(":")[0].replace("$", "")
    match = re.fullmatch(r"([A-Z]{1,3})(\d+)", first)
    if match is None or not plan.headers:
        return False
    row, col = formel.position(first)
    return plan.header_row < row <= plan.last_row and min(plan.headers) <= col <= max(plan.headers)


# So viele Korrekturen setzt „Alle einsetzen“ höchstens zusammen ein, und so lange darf es dafür nachrechnen.
MAX_JOINT = 12
JOINT_BUDGET = 20.0


def together(book: str, name: str, cells_of: dict[str, dict], hints: list[dict], filters: list | None = None,
             parts: Sheet | None = None, rules: list[dict] | None = None) -> dict:
    """„Alle einsetzen“: die Korrekturen des Blatts zusammen eingesetzt und das Blatt danach neu geprüft, so wie Excel
    es nach dem Einsetzen rechnet. Eine Korrektur bleibt draußen, wenn ihr Fehler danach noch dasteht oder sie zusammen
    mit den anderen einen neuen Fehler macht, den sie allein nicht macht. Zwei Korrekturen für dieselbe Zelle: die
    erste (Fehler vor Prüfen). Wo Daten und Formel denselben Fehler beheben könnten (eine Zahl als Text, die SVERWEIS
    nicht findet), kommen die Daten zuerst: Das ist die Ursache, die Formel nur ein Umweg."""
    order = {"fehler": 0, "pruefen": 1, "info": 2}
    fixable = sorted((h for h in hints if h.get("fix")), key=lambda h: (
        order.get(h["level"], 3), all(isinstance(v, str) for v in h["fix"]["cells"].values())))
    chosen, taken = [], set()
    for hint in fixable:
        cells = set(hint["fix"]["cells"])
        if cells & taken or len(taken | cells) > einsetzen.MAX_CELLS or len(chosen) >= MAX_JOINT:
            continue
        chosen.append(hint)
        taken |= cells
    if len(chosen) < 2:
        return {"state": "none", "text": T("Hier gibt es nichts, das sich zusammen einsetzen lässt.",
                                           "There is nothing here that can be applied together.")}
    cells = cells_of[name]
    plan = layout(Sheet(name, cells, pivots=list(parts.pivots) if parts is not None else []))
    others = {n: cells_of[n] for n in _uses(cells, cells_of, name)}
    before = {h["key"] for h in hints}
    deadline = time.monotonic() + JOINT_BUDGET

    def after(group: list[dict]) -> list[dict]:
        # Nachgerechnet wird mit den Zellen ([@Menge] steht im Fenster, C5 rechnet Gegenprobe).
        fixes = [{**h["fix"], "cells": h["fix"]["a1"]} if "a1" in h["fix"] else h["fix"] for h in group]
        new = einsetzen.refilled(einsetzen.applied(cells, fixes), plan, name, others, cells)
        # Ohne Pivot-Tabellen und Diagramme: Wird eine Zahl der Daten richtig, ist die Pivot-Tabelle danach alt. Das
        # ist kein Grund, die Korrektur wegzulassen.
        return sheet_result(book, name, {**cells_of, name: new}, filters, parts, charts=False, rules=rules)["hints"]

    def trouble(group: list[dict], allowed: set[str] = frozenset()) -> bool:
        found = after(group)
        keys = {h["key"] for h in found}
        return any(h["key"] in keys for h in group) or any(
            h["level"] == "fehler" and h["key"] not in before and h["key"] not in allowed for h in found)

    if trouble(chosen):
        # Was eine Korrektur allein an neuen Fehlern zeigt (eine Zahl als Text war, jetzt rechnet sie mit), stört nicht.
        alone: set[str] = set()
        for hint in chosen:
            if time.monotonic() > deadline:
                return {"state": "none", "text": T("Das Blatt ist zu groß, um alles zusammen nachzurechnen. Setz die "
                                                   "Korrekturen einzeln ein.", "The sheet is too large to recalculate "
                                                   "everything together. Apply the fixes one by one.")}
            alone |= {h["key"] for h in after([hint]) if h["key"] not in before}
        if trouble(chosen, alone):
            kept: list[dict] = []
            for hint in chosen:
                if time.monotonic() > deadline:
                    break
                if not trouble([*kept, hint], alone):
                    kept.append(hint)
            chosen = kept
    if not chosen:
        return {"state": "none", "text": T("Die Korrekturen stören sich gegenseitig. Setz sie einzeln ein, dann prüfe ich "
                                           "jedes Mal neu.", "The fixes get in each other's way. Apply them one by one, "
                                           "then I check again each time.")}
    return {"state": "ok", "fix": einsetzen.merged([h["fix"] for h in chosen]), "count": len(chosen),
            "left": len(fixable) - len(chosen)}


def _levels(done: dict) -> dict:
    levels = [h["level"] for h in done.get("hints", [])]
    return {"fehler": levels.count("fehler"), "pruefen": levels.count("pruefen")}


def _remembered(cells: dict, key, compute):
    """Bei gelesenen Zellen (xlsx.Cells, die sich nicht mehr ändern) einmal ausrechnen, sonst jedes Mal."""
    memo = getattr(cells, "memo", None)
    if memo is None:
        return compute()
    if key not in memo:
        memo[key] = compute()
    return memo[key]


def _cells_key(cells: dict) -> int:
    return _remembered(cells, "key", lambda: hash((frozenset(cells.items()), cut_of(cells))))


def _uses(cells: dict, cells_of: dict[str, dict], name: str) -> list[str]:
    """Die anderen Blätter, die in den Formeln dieses Blatts vorkommen (Tarife!A1, 'Daten 2024'!B3)."""
    text = _remembered(cells, "formulas", lambda: "\n".join(c.formula for c in cells.values() if c.formula).casefold())
    if "!" not in text:
        return []
    found = []
    for other in cells_of:
        if other == name:
            continue
        plain = other.casefold()
        quoted = "'" + plain.replace("'", "''") + "'!"
        if plain + "!" in text or quoted in text:
            found.append(other)
    return found


class Checker:
    """Ein Prüfer für die ganze App. check() aus einem Thread, nie zwei zugleich."""

    def __init__(self, link, data_dir: Path | None = None) -> None:
        self.link = link
        # Wo die eigenen Regeln je Spalte liegen (spaltenregeln.py), und die zuletzt gelesenen: Zeitstempel, Regeln, Schlüssel.
        self._data_dir = data_dir
        self._rules_read: tuple[float | None, list[dict], int] = (None, [], hash("[]"))
        self._lock = threading.Lock()
        self._tracker = WorkbookTracker()
        self._book = ""
        self._cells: dict[str, dict] = {}
        self._last: dict = {}
        self._report: tuple[int, Report] | None = None
        # Für das Add-in: je Blatt (Schlüssel, Ergebnis).
        self._sheets: dict[str, tuple[int, dict]] = {}
        # Die zuletzt geschickte Mappe (Blatt → Zellen), für die Herleitung einer Zelle.
        self._cells_of: dict[str, dict] = {}
        self._filters: dict[str, list] = {}
        # Je Blatt seine Pivot-Tabellen und Diagramme, und die Excel-Tabellen der Mappe.
        self._parts: dict[str, Sheet] = {}
        self._tables: dict = {}
        # Die zuletzt hineingezogene Datei: Name, je Blatt (Name, Ergebnis) und die Blätter, für ihren Bericht.
        self._file: tuple[str, list[tuple[str, dict]], list[Sheet]] | None = None
        # Die zuletzt beschriebene Mappe (Mappe in Worten): Schlüssel und Ergebnis.
        self._described: tuple[int, dict] | None = None
        # Was das Seitenfenster zuletzt je Blatt geschickt hat (Stempel, Blatt): Ein unverändertes Blatt schickt es
        # danach nur noch mit seinem Stempel, große Datenblätter also nicht bei jeder Änderung neu.
        self._received: dict[str, tuple[str, Sheet]] = {}
        # Große andere Blätter prüft ein Thread im Hintergrund, wenn Ruhe ist. Jede Anfrage zählt generation hoch:
        # Dann hört er nach dem laufenden Prüfschritt auf, damit die Anfrage schnell bleibt.
        self._generation = 0
        self._idle: threading.Thread | None = None
        self._idle_generation = -1
        self._active = ""
        # Je Blatt die Hinweise, die jemand mit „Passt so“ ausgeblendet hat (hint_id); das Fenster hält sie in der Mappe.
        self._passt: dict[str, frozenset[str]] = {}

    def check(self) -> dict:
        if not self._lock.acquire(blocking=False):
            return self._last or {"state": "connecting", "text": STATES["connecting"]}
        try:
            self._last = self._check()
            return self._last
        finally:
            self._lock.release()

    def _check(self) -> dict:
        link = self.link
        if link is None or not link.supported:
            return {"state": "off", "text": STATES["off"]}
        if not link.running():
            return {"state": "closed", "text": STATES["closed"]}
        snapshot = link.read()
        if snapshot is None:
            state = link.status.state
            if state == "busy" and self._last.get("state") == "ok":
                return self._last  # Was gelesen ist, gilt weiter.
            return {"state": state, "text": STATES.get(state, STATES["error"])}
        if snapshot.book != self._book:
            self._book, self._tracker, self._cells = snapshot.book, WorkbookTracker(), {}
        if snapshot.cells is not None:
            cells = snapshot.cells
            if snapshot.window is not None:
                # Bei großen Blättern liest Excel nur das Fenster: Der Rest bleibt, wie er war.
                top, left, bottom, right = snapshot.window
                old = self._cells.get(snapshot.sheet, {})
                kept = {p: c for p, c in old.items() if not (top <= p[0] <= bottom and left <= p[1] <= right)}
                cells = {**kept, **cells}
            self._cells[snapshot.sheet] = cells
            self._tracker.update(snapshot.book, [Sheet(snapshot.sheet, cells)], time.monotonic(), parts=False, partial=True,
                                 diff=False)
        else:
            book = link.book(snapshot)
            if book is None:
                return {"state": "error", "text": T("Die Zellen gibt Excel nicht her, und die gespeicherte Datei finde ich nicht.",
                                                    "Excel does not hand over the cells, and I cannot find the saved file.")}
            self._tracker.update(snapshot.book, book.sheets, time.monotonic(), parts=False, diff=False)
        self._tracker.focus(snapshot.sheet, None)
        found = material(self._tracker)
        state = self._tracker.active_state()
        if found is None or state is None:
            return {"state": "empty", "text": T("Auf diesem Blatt steht noch nichts.", "There is nothing on this sheet yet.")}
        _, shown, plan = state
        live = snapshot.cells is not None
        report = self._analyse(self._cells.get(snapshot.sheet) if live else shown, plan, live)
        return result(found.book, found.sheet, found.facts, report)

    # --- Eigene Regeln je Spalte ---

    def rules(self) -> tuple[list[dict], int]:
        """Die Regeln und ein Schlüssel dafür, neu gelesen, wenn sich die Datei geändert hat (auch aus der anderen App)."""
        if self._data_dir is None:
            return [], self._rules_read[2]
        try:
            stamp = spaltenregeln.path(self._data_dir).stat().st_mtime_ns
        except OSError:
            stamp = None
        if stamp != self._rules_read[0] or stamp is None:
            rules = spaltenregeln.load(self._data_dir)
            self._rules_read = (stamp, rules, hash(json.dumps(rules, sort_keys=True)))
        return self._rules_read[1], self._rules_read[2]

    def rule_list(self) -> dict:
        return {"state": "ok", "regeln": spaltenregeln.public(self.rules()[0])}

    def rule_save(self, rule: dict, before: dict | None = None) -> dict:
        if self._data_dir is None:
            return {"state": "error", "text": T("Regeln speichern geht hier nicht (kein Datenordner).",
                                                "Saving rules does not work here (no data folder).")}
        try:
            spaltenregeln.save(self._data_dir, rule, before)
        except (ValueError, OSError) as err:
            return {"state": "error", "text": str(err)}
        return self.rule_list()

    def rule_remove(self, column: str, sheet: str = "") -> dict:
        if self._data_dir is None:
            return {"state": "error", "text": T("Regeln löschen geht hier nicht (kein Datenordner).",
                                                "Deleting rules does not work here (no data folder).")}
        try:
            spaltenregeln.remove(self._data_dir, column, sheet)
        except OSError as err:
            return {"state": "error", "text": str(err)}
        return self.rule_list()

    def rule_column(self, sheet: str, cell: str, column: str = "", only: str | None = None) -> dict:
        """Für das Formular: die Spalte der markierten Zelle in der zuletzt geprüften Mappe (spaltenregeln.column), oder
        mit column die Spalte dieses Namens (zuerst auf dem Blatt sheet), um ihre Regel zu ändern."""
        with self._lock:
            cells_of, parts_of = dict(self._cells_of), dict(self._parts)
        if not cells_of or (not column and sheet not in cells_of):
            return {"state": "again"}
        rules = self.rules()[0]

        def analysed(name: str) -> Analysed:
            parts = parts_of.get(name)
            plan = layout(Sheet(name, cells_of[name], pivots=list(parts.pivots) if parts is not None else []))
            return Analysed(cells_of[name], plan, name)

        if column:
            wanted = spaltenregeln.norm(column)
            order = sorted(cells_of, key=lambda n: (n != sheet, spaltenregeln.norm(n) != spaltenregeln.norm(only or "")))
            for name in order:
                found = analysed(name)
                col = next((c for c, h in found.plan.headers.items() if spaltenregeln.norm(h) == wanted), None)
                if col is not None:
                    return spaltenregeln.column(found, (found.plan.header_row, col), rules, only)
            rule = spaltenregeln.find(rules, column, only or "")
            return {"state": "ok", "spalte": column, "blatt": only or "", "art": rule["art"] if rule else "", "spalten": [],
                    "jetzt": T(f"In der offenen Mappe gibt es keine Spalte „{column}“.",
                               f"There is no column “{column}” in the open workbook."),
                    "regel": spaltenregeln.public([rule])[0] if rule else None}
        try:
            pos = formel.position(cell.replace("$", "").upper())
        except (AttributeError, ValueError, KeyError, TypeError, IndexError):
            return {"state": "none", "text": T("Markiere in Excel eine Zelle der Spalte.", "Select a cell of the column in Excel.")}
        return spaltenregeln.column(analysed(sheet), pos, rules)

    def check_payload(self, payload: dict) -> dict:
        """Was das Seitenfenster schickt (addin.sheets_from), mit den Blättern, die es als unverändert meldet."""
        self._generation += 1
        self._passt = addin.passt_from(payload)
        try:
            book, active, sheets = addin.sheets_from(payload, self._known)
        except addin.Resend:
            return {"state": "again"}
        finally:
            # Die gelesenen Listen werden nicht mehr gebraucht: bei großen Blättern viel Speicher.
            payload.clear()
        return self.check_book(book, active, sheets)

    def _known(self, name: str, stamp: str) -> Sheet | None:
        found = self._received.get(name)
        return found[1] if found is not None and found[0] == stamp else None

    def others(self) -> dict:
        """Wie weit die übrigen Blätter geprüft sind, ohne neu zu rechnen: für das Fenster, solange der Hintergrund prüft."""
        with self._lock:
            return {"state": "ok", "book": self._book, "sheet": self._active, "others": self._others(self._active)}

    def _others(self, active: str | None, deadline: float | None = None) -> list[dict]:
        out = []
        for name, cells in self._cells_of.items():
            parts = self._parts.get(name)
            if name == active or not (cells or parts is not None and (parts.charts or parts.pivots)):
                continue
            big = deadline is not None and len(cells) > OTHERS_CELLS
            done = self._sheet_result(self._book, name, self._cells_of, time.monotonic() - 1 if big or deadline is None else deadline,
                                      self._filters.get(name))
            hidden = {"hidden": True} if getattr(parts, "hidden", False) else {}
            if done is None:
                out.append({"name": name, "fehler": None, "pruefen": None, "pending": True, **hidden})
            else:
                out.append({"name": name, **_levels(passed(done, self._passt.get(name))), **hidden})
        return out

    def _start_idle(self) -> None:
        """Die noch nicht geprüften Blätter nach und nach im Hintergrund prüfen, bis die nächste Anfrage kommt."""
        generation = self._generation
        if self._idle is not None and self._idle.is_alive() and self._idle_generation == generation:
            return
        # Ein älterer hört von selbst auf, sobald er die neue Zahl sieht.
        self._idle_generation = generation
        # In der Sprache der Anfrage, die ihn startet (sprache.py).
        self._idle = threading.Thread(target=contextvars.copy_context().run, args=(self._idle_work, generation), daemon=True,
                                      name="pruefer-hintergrund")
        self._idle.start()

    def _idle_work(self, generation: int) -> None:
        time.sleep(IDLE_WAIT)
        while self._generation == generation:
            with self._lock:
                if self._generation != generation:
                    return
                book, cells_of, filters_of = self._book, self._cells_of, self._filters
                todo = [n for n, c in cells_of.items() if c and self._cached(book, n, cells_of, filters_of.get(n)) is None]
                if not todo:
                    return
                name = todo[0]
                key = self._key(name, cells_of, filters_of.get(name))
                parts, tables, rules = self._parts.get(name), self._tables, self.rules()[0]
            try:
                with analyse_stoppable(lambda: self._generation != generation):
                    done = sheet_result(book, name, cells_of, filters_of.get(name), parts, tables, live=True, rules=rules)
            except Stopped:
                return
            with self._lock:
                if self._generation != generation or self._book != book or self._cells_of is not cells_of:
                    return
                self._sheets[name] = (key, done)

    def check_book(self, book: str, active: str, sheets: list[Sheet]) -> dict:
        """Die Mappe, wie das Excel-Add-in sie schickt: Das aktive Blatt wird genau geprüft, die übrigen
        nur gezählt (so weit die Zeit reicht), damit man sieht, wo sonst noch etwas ist. Große übrige Blätter
        prüft danach der Hintergrund (others() sagt, wie weit er ist)."""
        with self._lock:
            sheet = next((s for s in sheets if s.name == active), None) or next((s for s in sheets if s.cells), None)
            if sheet is None or not (sheet.cells or sheet.charts or sheet.pivots):
                return {"state": "empty", "book": book, "sheet": active, "text": T("Auf diesem Blatt steht noch nichts.", "There is nothing on this sheet yet.")}
            if book != self._book:
                self._book, self._sheets = book, {}
            cells_of = {s.name: s.cells for s in sheets}
            self._cells_of = cells_of
            filters = {s.name: [f.area for f in s.filters if f.filters] for s in sheets}
            self._filters = filters
            self._parts = {s.name: s for s in sheets}
            self._tables = pivotdiagramm.tables_of(sheets)
            self._received = {s.name: (s.stamp, s) for s in sheets if s.stamp}
            self._active = sheet.name
            for name in set(self._sheets) - set(cells_of):
                del self._sheets[name]
            found = dict(passed(self._sheet_result(book, sheet.name, cells_of, filters=filters.get(sheet.name)),
                                self._passt.get(sheet.name)))
            others = self._others(sheet.name, time.monotonic() + OTHERS_BUDGET)
            found["sheets"] = [s.name for s in sheets]
            found["others"] = others
            if any(o.get("pending") for o in others):
                self._start_idle()
            return found

    def check_file(self, name: str, data: bytes) -> dict:
        """Eine .xlsx-Datei prüfen, ohne sie in Excel zu öffnen: alle Blätter, nur im Arbeitsspeicher. Eine CSV- oder
        Textdatei (Export) als ein Blatt, ihre Daten (einlesen.csv_sheets)."""
        if einlesen.is_text(name):
            try:
                sheets = einlesen.csv_sheets(name, data)
            except ValueError as err:
                return {"state": "error", "text": str(err)}
        else:
            try:
                sheets = read_xlsx(data)
            except ValueError as err:
                return {"state": "error", "text": _unreadable(err)}
        cells_of = {s.name: s.cells for s in sheets}
        tables = pivotdiagramm.tables_of(sheets)
        rules = self.rules()[0]
        found = [(s.name, sheet_result(name, s.name, cells_of, [f.area for f in s.filters if f.filters], s, tables, rules=rules))
                 for s in sheets if s.cells or s.charts or s.pivots]
        with self._lock:
            self._file = (name, found, sheets)
        levels = [h["level"] for _, r in found for h in r["hints"]]
        hidden = {s.name for s in sheets if s.hidden}
        count = T(f"{len(found)} {'Blatt' if len(found) == 1 else 'Blätter'} geprüft",
                  f"{len(found)} {'sheet' if len(found) == 1 else 'sheets'} checked")
        if not found:
            text = T("In dieser Datei steht nichts.", "There is nothing in this file.")
        elif not levels:
            text = T(f"{count}, mir fällt nichts auf.", f"{count}, nothing stands out to me.")
        else:
            info, errors = levels.count("info"), levels.count("fehler")
            parts = [T(f"{errors} Fehler", f"{errors} {'error' if errors == 1 else 'errors'}") if errors else "",
                     T(f"{levels.count('pruefen')} zum Prüfen", f"{levels.count('pruefen')} to review") if "pruefen" in levels else "",
                     T(f"{info} {'Hinweis' if info == 1 else 'Hinweise'}", f"{info} {'note' if info == 1 else 'notes'}") if info else ""]
            text = f"{count}: {', '.join(p for p in parts if p)}."
        return {"state": "ok", "book": name, "text": text,
                "sheets": [{"name": n, "text": r["text"], "hints": r["hints"], "stats": r["stats"],
                            **({"failed": r["failed"]} if r.get("failed") else {}),
                            **({"hidden": True} if n in hidden else {})} for n, r in found]}

    def report(self, with_data: bool, file: bool = False) -> str | None:
        """Der Prüfbericht über alle Blätter der zuletzt geprüften Mappe (bericht.py), None ohne Mappe.
        file: über die zuletzt hineingezogene Datei."""
        self._generation += 1
        if file:
            with self._lock:
                done = self._file
            if done is None:
                return None
            name, found, sheets = done
            hints = {n: _levels(r) for n, r in found}
            mappe = inworten.describe(name, sheets, hints) if with_data else None
            used = spaltenregeln.used({s.name: s.cells for s in sheets}, self.rules()[0])
            return bericht.render(name, found, with_data, mappe=mappe, rules=used)
        with self._lock:
            if not self._cells_of:
                return None
            found = [(name, passed(self._sheet_result(self._book, name, self._cells_of, filters=self._filters.get(name)),
                                   self._passt.get(name)))
                     for name, cells in self._cells_of.items()
                     if cells or (name in self._parts and (self._parts[name].charts or self._parts[name].pivots))]
            mappe = self._describe({n: _levels(r) for n, r in found}) if with_data else None
            used = spaltenregeln.used(self._cells_of, self.rules()[0])
            return bericht.render(self._book, found, with_data, mappe=mappe, rules=used)

    def describe_book(self) -> dict:
        """Mappe in Worten (inworten.py) über die zuletzt geschickte Mappe, mit den Hinweisen je Blatt, soweit die
        Zeit reicht, sie zu zählen."""
        self._generation += 1
        with self._lock:
            if not self._cells_of:
                return {"state": "again"}
            deadline = time.monotonic() + DESCRIBE_BUDGET
            hints = {}
            for name in self._cells_of:
                done = passed(self._sheet_result(self._book, name, self._cells_of, deadline, self._filters.get(name)),
                              self._passt.get(name))
                if done is not None:
                    hints[name] = _levels(done)
            return self._describe(hints)

    def _describe(self, hints: dict[str, dict]) -> dict:
        sheets = [self._parts.get(name) or Sheet(name, cells) for name, cells in self._cells_of.items()]
        key = hash((self._book, tuple((s.name, _cells_key(s.cells), repr((s.charts, s.pivots))) for s in sheets),
                    tuple(sorted((n, h["fehler"], h["pruefen"]) for n, h in hints.items())), sprache.key()))
        if self._described is None or self._described[0] != key:
            self._described = (key, inworten.describe(self._book, sheets, hints))
        return self._described[1]

    def compare(self, name: str, data: bytes) -> dict:
        """Die zuletzt geprüfte Mappe mit einer früheren Version als Datei vergleichen (vergleich.py)."""
        with self._lock:
            if not self._cells_of:
                return {"state": "again"}
            cells_of, book = dict(self._cells_of), self._book
        try:
            old = read_xlsx(data)
        except ValueError as err:
            return {"state": "error", "text": _unreadable(err)}
        return vergleich.compare({s.name: s.cells for s in old}, cells_of, name, book)

    def explain(self, sheet: str, cell: str) -> dict:
        """Woraus die Zelle entsteht und wohin sie fließt (herleitung.py), mit der zuletzt geprüften Mappe."""
        with self._lock:
            if not self._cells_of:
                return {"state": "again"}
            return herleitung.explain(self._cells_of, sheet, cell)

    def _key(self, name: str, cells_of: dict[str, dict], filters: list | None) -> int:
        """Woran sich zeigt, ob das gemerkte Ergebnis eines Blatts noch gilt: seine Zellen, die Blätter, auf die es
        verweist, Filter, Pivot-Tabellen und Diagramme, die eigenen Regeln."""
        cells = cells_of[name]
        uses = _uses(cells, cells_of, name)
        parts = self._parts.get(name)
        if parts is not None and (parts.charts or parts.pivots):
            # Ihre Daten können auf jedem Blatt stehen.
            uses = [n for n in cells_of if n != name]
            shape = repr((parts.charts, parts.pivots, sorted(self._tables.items(), key=lambda t: t[0])))
        else:
            shape = ""
        if parts is not None and parts.merged:
            shape += repr((parts.merged, parts.filters))
        return hash((_cells_key(cells), tuple(sorted((n, _cells_key(cells_of[n])) for n in uses)), tuple(filters or ()), shape,
                     self.rules()[1], sprache.key()))

    def _cached(self, book: str, name: str, cells_of: dict[str, dict], filters: list | None) -> dict | None:
        cached = self._sheets.get(name)
        return cached[1] if cached is not None and cached[0] == self._key(name, cells_of, filters) else None

    def _sheet_result(self, book: str, name: str, cells_of: dict[str, dict], deadline: float | None = None,
                      filters: list | None = None) -> dict | None:
        """Hinweise für ein Blatt der Mappe, gemerkt, solange sich das Blatt und die Blätter, auf die es
        verweist, nicht ändern. Mit deadline: None, wenn es neu gerechnet werden müsste und die Zeit um ist."""
        key = self._key(name, cells_of, filters)
        cached = self._sheets.get(name)
        if cached is not None and cached[0] == key:
            return cached[1]
        if deadline is not None and time.monotonic() > deadline:
            return None
        done = sheet_result(book, name, cells_of, filters, self._parts.get(name), self._tables, live=True, rules=self.rules()[0])
        self._sheets[name] = (key, done)
        return done

    def joint(self, name: str) -> dict:
        """„Alle einsetzen“ für ein Blatt der zuletzt geschickten Mappe (together)."""
        self._generation += 1
        with self._lock:
            if name not in self._cells_of:
                return {"state": "again", "text": T("Ich prüfe zuerst neu, dann kannst du es noch einmal versuchen.",
                                                    "I check again first, then you can try once more.")}
            filters = self._filters.get(name)
            done = self._sheet_result(self._book, name, self._cells_of, filters=filters)
            return together(self._book, name, self._cells_of, done["hints"], filters, self._parts.get(name), self.rules()[0])

    def _analyse(self, cells: dict, plan, exact: bool) -> Report:
        """Die tiefe Prüfung (analyse.py) nur, wenn sich das Blatt geändert hat."""
        rules, rules_key = self.rules()
        key = hash((frozenset(cells.items()), exact, rules_key, sprache.key()))
        if self._report is None or self._report[0] != key:
            self._report = (key, analyse(cells, plan, exact=exact, rules=rules))
        return self._report[1]
