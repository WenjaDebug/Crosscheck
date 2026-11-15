"""Prüfungen für Datenauswertungen: was in Tabellen mit Fällen, Patientinnen oder Messwerten typisch schiefgeht.

Fehlende Werte und Codes wie 999 für „fehlt“, unmögliche Werte (Alter 212, Austritt vor Eintritt), doppelte
Fallnummern und doppelt erfasste Zeilen, zwei Einheiten in einer Spalte (1,72 und 172), MITTELWERT und Co.
über einen Bereich, der Zeilen auslässt, ein Nenner, der beim Kopieren mitrutscht, und SUMME über eine
gefilterte Tabelle. Alles ohne KI, aus den Zellen, die Excel zeigt (analyse.Sheet). Bereiche, die nicht zu
ihrer Kennzahl passen, prüft bereiche.py, falsch gewichtete Quoten und Wachstum quoten.py.
"""

from __future__ import annotations

import math
import re
import statistics
from bisect import bisect_left, bisect_right
from collections import Counter
from datetime import date

from . import alstext, bereiche, einsetzen, sprache, fallen, formel, gruppen, kategorien, kopf, massstab, quoten, woerter
from .analyse import _CHANGES, Finding, Sheet, _fmt, _german, address, links
from .sprache import T
from .xlsx import Cell, column_letters

Pos = tuple[int, int]
Area = tuple[int, int, int, int]

MIN_ROWS = 10
MAX_LISTED = 3
# Zahlen, die in Datensätzen oft für „fehlt“ oder „unbekannt“ stehen.
CODES = (-1.0, -9.0, -99.0, -999.0, -9999.0, 99.0, 999.0, 9999.0, 99999.0)
_PLACEHOLDERS = {"na", "n/a", "n.a.", "nan", "k.a.", "ka", "k. a.", "-", "--", "?", "fehlt", ".", "null", "none", "#nv", "#n/a",
                 "x"} | woerter.PLACEHOLDERS
_AGGREGATES = {"SUM", "AVERAGE", "MEDIAN", "COUNT", "COUNTA", "MIN", "MAX", "STDEV", "STDEV.S", "STDEV.P", "VAR", "VAR.S",
               "VAR.P", "SUBTOTAL", "AGGREGATE", "AVERAGEA", "PERCENTILE", "QUARTILE"}
_TOTAL_WORDS = ("summe", "total", "gesamt", "mittelwert", "durchschnitt", "median", "anzahl gesamt", "sum", "mean", "average")
_CHANGE_WORDS = _CHANGES

# Werte, die es nicht geben kann: Muster im Spaltenkopf, kleinster und größter möglicher Wert, Erklärung (deutsch,
# englisch).
_LIMITS = [
    (r"\balter\b|\bage\b|^alter|" + woerter.AGE, 0.0, 120.0,
     ("ein Alter unter 0 oder über 120 Jahren", "age below 0 or above 120 years")),
    (r"\bbmi\b|" + woerter.BMI, 10.0, 80.0, ("einen BMI unter 10 oder über 80", "BMI below 10 or above 80")),
    (r"temperatur|\btemp\b|°c|" + woerter.TEMPERATURE, 25.0, 45.0,
     ("eine Körpertemperatur unter 25 oder über 45 °C", "body temperature below 25 or above 45 °C")),
    (r"\bpuls\b|herzfrequenz|\bhf\b|heart.?rate|" + woerter.PULSE, 20.0, 250.0,
     ("einen Puls unter 20 oder über 250", "pulse below 20 or above 250")),
    (r"systol|" + woerter.SYSTOLIC, 40.0, 300.0,
     ("einen systolischen Blutdruck unter 40 oder über 300", "systolic blood pressure below 40 or above 300")),
    (r"diastol", 20.0, 200.0,
     ("einen diastolischen Blutdruck unter 20 oder über 200", "diastolic blood pressure below 20 or above 200")),
    # „Dosis pro kg“ und „mg/kg“ sind kein Gewicht.
    (r"gewicht.*kg|(?<!/)(?<!pro )(?<!je )(?<!per )\bkg\b|\bweight\b", 0.2, 400.0,
     ("ein Gewicht unter 0,2 oder über 400 kg", "weight below 0,2 or above 400 kg")),
    (r"sättigung|spo2|sao2", 50.0, 100.0,
     ("eine Sauerstoffsättigung unter 50 oder über 100 %", "oxygen saturation below 50 or above 100 %")),
]
_NOT_NEGATIVE = (r"anzahl|menge|dauer|liegedauer|tage\b|stunden|minuten|\bcount\b|\bn\b|kosten|betrag|preis|gewicht|größe|groesse|"
                 r"alter|" + woerter.NOT_NEGATIVE + "|" + woerter.AGE)
_START = r"eintritt|aufnahme|beginn|\bstart|\bvon\b|admission|eingang|op.?datum|diagnose|" + woerter.START
_END = r"austritt|entlassung|\bende\b|\bend\b|\bbis\b|discharge|abschluss|tod|exitus|verstorben|" + woerter.END
_BIRTH = r"geburt|geb\.|\bdob\b|birth|" + woerter.BIRTH
_ID = r"\bid\b|-id\b|\bnr\b|nr\.|nummer|fall|patient|\bpid\b|\bmrn\b|kennung|versicherten|\bpat\b|" + woerter.ID


def checks(sheet: Sheet, filtered: list[Area] | None = None) -> tuple[list[Finding], set[str]]:
    """Die Hinweise und die Schlüssel älterer Hinweise (analyse.py), die sie genauer sagen und ersetzen."""
    plan = sheet.plan
    if plan.header_row is None:
        return _other_ranges(sheet), set()
    rows = _data_rows(sheet)
    found: list[Finding] = _other_ranges(sheet)
    columns = {col: _column(sheet, col, rows) for col in sorted(plan.headers)}
    unfilled, blank = _unfilled(sheet, rows)
    found += unfilled
    # In einem Zeilenmodell (Finanzplan) ist eine Spalte ein Jahr: Leere Zwischenüberschriften fehlen dort nicht, und
    # Codes, unmögliche Werte oder zwei Einheiten je Spalte gibt es so nicht.
    ungrouped: set[str] = set()
    if not plan.across:
        found += _missing(sheet, columns, rows, blank)
        for col, info in columns.items():
            header = plan.headers[col]
            found += _codes(sheet, col, header, info)
            found += _impossible(sheet, col, header, info)
            found += _units(sheet, col, header, info)
            found += _odd_rate(sheet, col, header, info)
        # Tippfehler und anders geschriebene Namen, Zeilen gegen eine Regel zwischen zwei Spalten (Code und Diagnose).
        found += kategorien.checks(sheet, columns, rows)
        # Laborwerte im Langformat: Kalium 45 fällt nur gegen die übrigen Kaliumwerte auf.
        grouped, ungrouped = gruppen.checks(sheet, columns, rows)
        found = [f for f in found if f.key not in ungrouped] + grouped
    found += _dates(sheet, columns)
    # Eine negative Dauer, die nur aus einem Ende vor dem Anfang folgt, sagt der Hinweis dazu (order:).
    follows = sheet.memo.get("auswertung.order") or set()
    found = [f for f in found if not (f.key.startswith("impossible:") and _only_follows(sheet, f, follows))]
    astext = alstext.checks(sheet, columns)
    found += astext
    found += _ids(sheet, columns)
    found += _same_rows(sheet, rows)
    # Mitrutschende Bereiche erklären alle Zellen, die heruntergezogen sind; was dort sonst auffällt, folgt daraus.
    slid, moving = bereiche.sliding(sheet, rows)
    held, still = bereiche.frozen(sheet, rows)
    slid, moving = slid + held, moving | still
    # Falsch gewichtet sagt mehr als „passt nicht zum Namen“: Die Ursache steht dabei.
    weighted = [f for f in quoten.checks(sheet, rows, columns) if f.key.partition(":")[2] not in moving]
    moving |= {f.key.partition(":")[2] for f in weighted}
    named = [f for f in bereiche.checks(sheet, rows, columns) if f.key.partition(":")[2] not in moving]
    # Was der Name der Kennzahl genauer sagt, ersetzt den allgemeinen Hinweis zum Bereich derselben Zelle.
    exact = {f.key.partition(":")[2] for f in named if f.key.startswith(("label:", "count:"))} | moving
    found += [f for f in _ranges(sheet, rows) if f.key.partition(":")[2] not in exact] + named + slid + weighted
    found += massstab.checks(sheet, rows)
    found += kopf.checks(sheet, rows)
    taken = {f.key.partition(":")[2] for f in found}
    found += [f for f in fallen.checks(sheet, rows) if f.key.partition(":")[2] not in taken]
    found += _denominators(sheet, covered := set())
    found += _filtered(sheet, filtered or [])
    for item in found:
        kind, _, where = item.key.partition(":")
        if kind == "units":
            # Die Werte in der anderen Einheit sind keine Ausreißer.
            covered |= {f"outlier:{cell}" for cell in re.findall(r"\b[A-Z]{1,3}\d+\b", item.text)}
        if kind in ("placeholder", "astext"):
            # Zahlen als Text sagt alstext.py genauer, mit den Formeln, die dadurch anders rechnen.
            covered.add(f"text:{plan.headers[formel.position(where)[1]]}")
    return found, covered | ungrouped


def table(sheet: Sheet) -> tuple[list[int] | None, dict | None]:
    """Die Datenzeilen und je Spalte, was in ihr steht, für Prüfungen außerhalb (summen.py)."""
    if sheet.plan.header_row is None:
        return None, None
    rows = _data_rows(sheet)
    return rows, {col: _column(sheet, col, rows) for col in sorted(sheet.plan.headers)}


# Zahlen, die etwas benennen statt messen: Postleitzahl, Jahr, Monat, Code. Ausreißer und Einheiten sagen dort nichts.
_NAMING = (r"\bplz\b|postleitzahl|\bzip\b|postal|telefon|\bphone\b|\bjahr\b|\bmonat\b|\btag\b|\bwoche\b|\bkw\b|quartal|"
           r"\bstunde\b|\byear\b|\bmonth\b|\bday\b|\bweek\b|quarter|\bhour\b|\bcode\b|\bicd\b|\bnr\b|" + woerter.NAMING)


def is_naming(header: str) -> bool:
    """Eine Spalte mit Nummern, die etwas benennen (Postleitzahl, Monat, Code), nicht mit Messwerten."""
    return is_id(header) or bool(re.search(_NAMING, header.lower()))


def is_id(header: str) -> bool:
    """Fall-Nr, Patienten-ID, Kennung: eine Spalte, in der jede Zeile ihre eigene Nummer haben sollte."""
    return bool(re.search(_ID, header.lower()))


def codes_in(values: list[float]) -> set[float]:
    """Zahlen wie 999 oder -99, die mehrmals vorkommen und weit weg von allen übrigen Werten liegen."""
    if len(values) < MIN_ROWS:
        return set()
    counts = Counter(values)
    found = set()
    for code in CODES:
        if counts.get(code, 0) < 2:
            continue
        rest = [v for v in values if v not in CODES]
        if len(rest) < 5:
            continue
        low, high = min(rest), max(rest)
        spread = max(high - low, abs(statistics.median(rest)) * 0.1, 1e-9)
        if code > high + 2 * spread or code < low - 2 * spread or (code < 0 <= low):
            found.add(code)
    return found


# --- Spalten ---


class _Info:
    """Was in einer Spalte steht: Zahlen, Datumswerte (als Excel-Zahl), Text, leere Zeilen."""

    def __init__(self) -> None:
        self.numbers: dict[int, float] = {}
        self.dates: dict[int, float] = {}
        self.texts: dict[int, str] = {}
        self.empty: list[int] = []
        self.formulas = 0

    @property
    def kind(self) -> str:
        filled = len(self.numbers) + len(self.dates) + len(self.texts)
        if not filled:
            return "leer"
        if len(self.dates) >= 0.8 * filled:
            return "datum"
        if len(self.numbers) >= 0.8 * filled:
            return "zahl"
        return "text"


def in_words(sheet: Sheet, pos: tuple[int, int]) -> str:
    """Für die Herleitung: welche Zeilen der Tabelle die Formel in pos nimmt, in einem Satz (bereiche.in_words)."""
    if sheet.plan.header_row is None or sheet.nodes.get(pos) is None:
        return ""
    rows = _data_rows(sheet)
    columns = {col: _column(sheet, col, rows) for col in sorted(sheet.plan.headers)}
    try:
        return bereiche.in_words(sheet, pos, rows, columns)
    except (formel.Unknown, ValueError, TypeError, KeyError, IndexError, RecursionError):
        return ""

def _data_rows(sheet: Sheet) -> list[int]:
    """Die Datenzeilen der Tabelle, ohne Summen- und Mittelwertzeilen am Ende. Einmal je Blatt gerechnet: die Liste
    nicht ändern."""
    if ("auswertung", "rows") not in sheet.memo:
        sheet.memo[("auswertung", "rows")] = _rows_of(sheet)
    return sheet.memo[("auswertung", "rows")]


def _rows_of(sheet: Sheet) -> list[int]:
    plan = sheet.plan
    rows = list(range(plan.header_row + 1, plan.last_row + 1))
    out = []
    total = False
    for r in rows:
        name = plan.names.get(r, "")
        if name and (name.strip().lower().startswith(_TOTAL_WORDS) or woerter.starts_with(name, woerter.SUMMARY_ROWS)):
            total = True
            continue
        # Eine ganz leere Zeile mitten in der Tabelle trennt zwei Teile (zwei Monate untereinander): Ihr fehlt kein
        # Wert, sie ist keine Datenzeile.
        if not any((cell := sheet.cells.get((r, c))) is not None and (cell.formula is not None or cell.value.strip())
                   for c in plan.headers):
            continue
        if any(_sums_above(sheet.nodes.get((r, c)), r, c) for c in plan.headers):
            total = True
            continue
        # „Fälle =ANZAHL2(A2:A41)“ unter der Spalte Kosten: eine Kennzahl über eine andere Spalte von oben an.
        typed = any(isinstance(sheet.inputs.get((r, c)), float) and (r, c) in sheet.cells and
                    not sheet.cells[(r, c)].formula for c in plan.headers)
        if not typed and any(_counts_from_top(sheet.nodes.get((r, c)), r, plan.header_row) and _alone(sheet, (r, c))
                             for c in plan.headers):
            total = True
            continue
        # Unter der Summe: „Letzter = B13“ holt nur einen Wert aus der Spalte darüber.
        if total and any(_picks_above(sheet.nodes.get((r, c)), r, c) for c in plan.headers):
            continue
        out.append(r)
    return out


_PICKS = ("INDEX", "INDIRECT", "OFFSET", "CHOOSE")


def _picks_above(node, row: int, col: int) -> bool:
    """Eine Zelle, die nur einen Wert aus ihrer eigenen Spalte darüber holt (=B13, =INDEX(B2:B13;12), INDIREKT,
    BEREICH.VERSCHIEBEN): keine Datenzeile, sondern eine Kennzahl daraus. Was damit rechnet (MwSt = B15 * 0,19),
    bleibt eine Zeile für sich."""
    if node is None or not (node.kind == "ref" or node.kind == "call" and node.value in _PICKS):
        return False
    cells, ranges, _ = formel.references(node)
    if not (cells or ranges or formel.dynamic(node)):
        return False
    return all(c == col and r < row for r, c in cells) and all(c1 == c2 == col and r2 < row for _, c1, r2, c2 in ranges)


def _sums_above(node, row: int, col: int) -> bool:
    """Eine Summen- oder Mittelwertzeile: die Zelle fasst ihre eigene Spalte darüber zusammen (D9 = SUMME(D2:D8))."""
    if not _aggregate(node):
        return False
    return any(a.kind == "range" and a.value[1] <= col <= a.value[3] and a.value[2] < row for a in node.args)


def _counts_from_top(node, row: int, header_row: int) -> bool:
    """Eine Kennzahl über eine ganze Spalte darüber, von der ersten Datenzeile an (=ANZAHL2(A2:A41) neben dem Namen
    „Fälle“): keine Datenzeile. Ein gleitender Mittelwert (=MITTELWERT(B3:B5)) fängt nicht oben an."""
    if not _aggregate(node):
        return False
    return any(a.kind == "range" and a.value[0] <= header_row + 2 and a.value[2] < row and a.value[2] - a.value[0] >= 2
               for a in node.args)


def _alone(sheet: Sheet, pos: Pos) -> bool:
    """Eine Formel, die nicht mit der darüber oder darunter heruntergezogen ist (keine Laufsumme, kein Fenster)."""
    from .workbook import pattern

    shape = pattern(sheet.cells[pos].formula or "", pos)
    for near in ((pos[0] - 1, pos[1]), (pos[0] + 1, pos[1])):
        cell = sheet.cells.get(near)
        if cell is not None and cell.formula and pattern(cell.formula, near) == shape:
            return False
    return True


def _aggregate(node) -> bool:
    return node is not None and node.kind == "call" and node.value in _AGGREGATES


def _column(sheet: Sheet, col: int, rows: list[int]) -> _Info:
    info = _Info()
    for r in rows:
        cell = sheet.cells.get((r, col))
        if cell is None or cell.value.strip() == "":
            if cell is None or cell.formula is None:
                info.empty.append(r)
            continue
        if cell.formula is not None:
            info.formulas += 1
        shown, used = sheet.shown.get((r, col)), sheet.inputs.get((r, col))
        if isinstance(shown, str) and isinstance(used, float) and formel.serial_of(shown) is not None:
            info.dates[r] = used
        elif isinstance(shown, float) and not isinstance(shown, bool):
            info.numbers[r] = shown
        elif isinstance(shown, str):
            info.texts[r] = shown
    return info


def _cells(col: int, rows: list[int]) -> str:
    shown = ", ".join(address((r, col)) for r in rows[:MAX_LISTED])
    return shown + (_more(len(rows) - MAX_LISTED) if len(rows) > MAX_LISTED else "")


def _more(count: int) -> str:
    """„ und 4 weitere“ am Ende einer Aufzählung."""
    return T(f" und {count} weitere", f" and {count} more")


def _fn(name: str) -> str:
    """Eine Funktion im Satz: MITTELWERT auf Deutsch, AVERAGE auf Englisch."""
    return T(_german(name), name)


def order_start(later: str, earlier: str) -> str:
    """Wie der Hinweis zu einem Ende vor dem Anfang beginnt („Austritt liegt vor Eintritt “): spaltenregeln.py erkennt
    daran das Paar seiner Regel."""
    return T(f"{later} liegt vor {earlier} ", f"{later} is before {earlier} ")


def _missing(sheet: Sheet, columns: dict[int, _Info], rows: list[int], blank: set[Pos] = frozenset()) -> list[Finding]:
    """Fehlende Werte je Spalte, und Platzhalter wie „k.A.“ in Zahlenspalten, in einem Hinweis. blank: leere Zellen,
    die ein anderer Hinweis schon erklärt (Formeln, die in neuen Zeilen fehlen)."""
    if len(rows) < MIN_ROWS:
        return []
    parts, first = [], None
    placeholders = []
    # Die Anfangszeile (Jahr 0, Anfangsbestand) ist fast leer, und das gehört so.
    start = rows[0] if _opening(sheet, rows[0], rows[1:], sheet.plan.headers) else None
    for col, info in columns.items():
        header = sheet.plan.headers[col]
        if info.kind in ("zahl", "datum"):
            marks = [r for r, t in info.texts.items() if t.strip().casefold() in _PLACEHOLDERS]
            if marks:
                placeholders.append((col, header, marks))
        empty = [r for r in info.empty if (r, col) not in blank and r != start]
        filled = len(rows) - len(info.empty)
        if not empty or filled < len(rows) / 2:
            continue
        # Eine Spalte aus lauter Formeln hat keine fehlenden Daten: Eine leere Zelle darin ist eine fehlende Formel
        # oder der Anfang einer Laufsumme, das sagen die Prüfungen der Formeln.
        gaps = set(info.empty)
        if all(sheet.cells[(r, col)].formula is not None for r in rows if r not in gaps and (r, col) in sheet.cells):
            continue
        share = len(empty) / len(rows)
        parts.append(T(f"{header} {len(empty)} von {len(rows)} ({_percent(share)}, {_cells(col, empty)})",
                       f"{header} {len(empty)} of {len(rows)} ({_percent(share)}, {_cells(col, empty)})"))
        first = first or (empty[0], col)
    found = []
    if parts:
        found.append(Finding(
            f"missing:{address(first)}",
            T(f"Fehlende Werte: {'; '.join(parts[:4])}{' und weitere Spalten' if len(parts) > 4 else ''}.",
              f"Missing values: {'; '.join(parts[:4])}{' and more columns' if len(parts) > 4 else ''}."),
            "info", T("MITTELWERT und SUMME lassen leere Zellen still weg, ANZAHL zählt sie nicht. Sag in der Auswertung, "
                      "wie viele Werte fehlen, und ob die Fehlenden anders sind als die übrigen.",
                      "AVERAGE and SUM silently skip empty cells, and COUNT does not count them. Say in the analysis how "
                      "many values are missing, and whether the missing ones differ from the others."),
        ))
    for col, header, marks in placeholders[:2]:
        text = ", ".join(T(f"„{sheet.cells[(r, col)].value}“", f"“{sheet.cells[(r, col)].value}”") for r in marks[:MAX_LISTED])
        found.append(Finding(
            f"placeholder:{address((marks[0], col))}",
            T(f"In {header} steht bei {_cells(col, marks)} {text} statt einer Zahl. MITTELWERT und SUMME lassen das still weg.",
              f"In {header}, {_cells(col, marks)} {'holds' if len(marks) == 1 else 'hold'} {text} instead of a number. "
              "AVERAGE and SUM silently skip that."),
            "pruefen", T("Fehlt der Wert wirklich? Dann lass die Zelle leer und zähle die fehlenden Werte, sonst trag die Zahl ein.",
                         "Is the value really missing? Then leave the cell empty and count the missing values; otherwise "
                         "enter the number."),
        ))
    return found


def _unfilled(sheet: Sheet, rows: list[int]) -> tuple[list[Finding], set[Pos]]:
    """Neue Zeilen unten, in denen die Formeln fehlen, die darüber in jeder Zeile stehen: Die Daten wurden angehängt,
    die Formeln nicht heruntergezogen. Mit der Korrektur, die sie aus der Zeile darüber herunterzieht. Gibt auch die
    leeren Zellen, die der Hinweis erklärt."""
    from .workbook import pattern
    from .xlsx import shift

    plan = sheet.plan
    if len(rows) < MIN_ROWS:
        return [], set()
    known = set(rows)
    by_rows: dict[tuple[int, ...], list[tuple[int, int]]] = {}
    for col in sorted(plan.headers):
        trailing = []
        for r in reversed(rows):
            if (r, col) in sheet.cells:
                break
            trailing.append(r)
        if not trailing:
            continue
        trailing.reverse()
        last = trailing[0] - 1
        cell = sheet.cells.get((last, col))
        # Unmittelbar darüber die Formel, heruntergezogen über fast die ganze Tabelle, und die neuen Zeilen schließen
        # ohne Lücke an.
        if last not in known or cell is None or not cell.formula or len(trailing) > max(1, len(rows) // 4) or \
                trailing != list(range(trailing[0], trailing[-1] + 1)):
            continue
        # Dieselbe Formel (heruntergezogen) in fast allen Zeilen darüber; eine abweichende dazwischen sagt _odd_ones.
        shape = pattern(cell.formula, (last, col))
        above = [r for r in rows if r < trailing[0]]
        same = sum(1 for r in above if (near := sheet.cells.get((r, col))) is not None and near.formula
                   and pattern(near.formula, (r, col)) == shape)
        if same < MIN_ROWS or same < 0.9 * len(above):
            continue
        by_rows.setdefault(tuple(trailing), []).append((col, last))
    found, blank = [], set()
    for trailing, columns in by_rows.items():
        last = trailing[0] - 1
        # Die neuen Zeilen haben Daten: In den Spalten ihrer eigenen Tabelle (ohne leere Spalte dazwischen; eine
        # Nachschlagetabelle daneben ist oft länger), in denen darüber keine Formel steht, ist mindestens die Hälfte
        # gefüllt.
        own = _side_by_side(sorted(plan.headers), {c for c, _ in columns})
        inputs = [c for c in own if not (sheet.cells.get((last, c)) and sheet.cells[(last, c)].formula)]
        if not inputs or any(sum(1 for c in inputs if sheet.shown.get((r, c)) not in (None, "")) * 2 < len(inputs)
                             for r in trailing):
            continue
        missing = {c for c, _ in columns}
        changes = {}
        for col, _ in columns:
            for r in trailing:
                formula = shift(sheet.cells[(last, col)].formula, r - last, 0)
                try:
                    cells, ranges, _ = formel.references(formel.parse(formula), sheet.name)
                except (ValueError, formel.Unknown):
                    cells, ranges = {(formel.MAX_ROW, 1)}, []
                # Liest die Formel eine Zeile unter der Tabelle (=B5-B4 für „zur nächsten Zeile“), fehlt sie in der
                # letzten Zeile mit Absicht. Und was sie in ihrer eigenen Zeile liest (die Fall-Nr in A), muss da sein.
                if any(p[0] > rows[-1] for p in cells) or any(rows[-1] < a[0] for a in ranges) or \
                        any(p[0] == r and p[1] not in missing and sheet.shown.get(p) in (None, "") for p in cells):
                    break
                changes[(r, col)] = formula
            else:
                continue
            columns = [c for c in columns if c[0] != col]
            for r in trailing:
                changes.pop((r, col), None)
        if not columns:
            continue
        names = [plan.headers[c] for c, _ in columns]
        one, single = len(columns) == 1, len(trailing) == 1
        side = [c for c, _ in columns] == list(range(columns[0][0], columns[-1][0] + 1))
        and_ = T(" und ", " and ")
        letters = (column_letters(columns[0][0]) if one else
                   T(f"{column_letters(columns[0][0])} bis {column_letters(columns[-1][0])}",
                     f"{column_letters(columns[0][0])} to {column_letters(columns[-1][0])}") if side else
                   ", ".join(column_letters(c) for c, _ in columns[:-1]) + and_ + column_letters(columns[-1][0]))
        listed = names[0] if one else ", ".join(names[:-1]) + and_ + names[-1]
        if sprache.english():
            where = f"row {trailing[0]}" if single else f"rows {trailing[0]} to {trailing[-1]}"
            text = (f"In {where}, the {'formula' if one else 'formulas'} of {listed} ({letters}) "
                    f"{'is' if one else 'are'} missing. Above, {'it is' if one else 'they are'} in every row down to {last}: "
                    f"the new {'row was' if single else 'rows were'} appended, but the formulas were not filled down. So "
                    f"the results are missing there, and whatever calculates over {'this column' if one else 'these columns'} "
                    f"leaves out {'the row' if single else 'the rows'}.")
        else:
            where = f"Zeile {trailing[0]}" if single else f"den Zeilen {trailing[0]} bis {trailing[-1]}"
            text = (f"In {where} {'fehlt die Formel' if one else 'fehlen die Formeln'} von {listed} ({letters}). Darüber "
                    f"{'steht sie' if one else 'stehen sie'} in jeder Zeile bis {last}: "
                    f"{'Die neue Zeile wurde' if single else 'Die neuen Zeilen wurden'} angehängt, die Formeln aber nicht "
                    f"heruntergezogen. So fehlen dort die Ergebnisse, und was über {'diese Spalte' if one else 'diese Spalten'} "
                    f"rechnet, lässt {'die Zeile' if single else 'die Zeilen'} aus.")
        span = address((last, columns[0][0])) + ("" if one or not side else f":{address((last, columns[-1][0]))}")
        tip = (T(f"Markier {span} und zieh {'die Formel' if one else 'die Formeln'} bis Zeile {trailing[-1]} herunter.",
                 f"Select {span} and fill {'the formula' if one else 'the formulas'} down to row {trailing[-1]}.")
               if one or side else T(f"Zieh die Formeln aus Zeile {last} bis Zeile {trailing[-1]} herunter.",
                                     f"Fill the formulas from row {last} down to row {trailing[-1]}."))
        found.append(Finding(f"unfilled:{address((trailing[0], columns[0][0]))}", text, "fehler", tip,
                             fix=einsetzen.formulas(sheet, changes)))
        blank |= set(changes)
    return found, blank


def _side_by_side(headers: list[int], columns: set[int]) -> list[int]:
    """Die Spalten mit Überschrift, die ohne leere Spalte dazwischen an columns grenzen: die eigene Tabelle."""
    runs, run = [], []
    for c in headers:
        if run and c != run[-1] + 1:
            runs.append(run)
            run = []
        run.append(c)
    if run:
        runs.append(run)
    return [c for run in runs if columns & set(run) for c in run]


def _percent(share: float) -> str:
    return f"{share * 100:.0f} %" if share >= 0.01 else T("unter 1 %", "under 1 %")


# Sätze, die eine Regel vorgibt (Steuer, Zins, AHV): in einer Liste meist überall gleich.
_RATE = r"(?<!um)satz|\bzins|steuer|mwst|\bust\b|\bvat\b|\btax\b|\brate\b|" + woerter.RATE


def _odd_rate(sheet: Sheet, col: int, header: str, info: _Info) -> list[Finding]:
    """Ein Satz, der in einer oder zwei Zeilen knapp anders ist als in allen übrigen: 7,7 % neben 29-mal 8,1 % ist
    meist der alte Satz oder vertippt. Ein ganz anderer Satz (7 % neben 19 %, 0 % steuerfrei) ist oft Absicht und
    fällt darum nicht auf. Nur, wo eine Formel mit der Zelle rechnet."""
    if info.kind != "zahl" or info.formulas or len(info.numbers) < MIN_ROWS or not re.search(_RATE, header.lower()):
        return []
    counts = Counter(info.numbers.values())
    usual, n = counts.most_common(1)[0]
    odd = sorted(r for r, v in info.numbers.items() if v != usual)
    if not odd or len(odd) > 2 or n < 0.9 * len(info.numbers) or usual == 0:
        return []
    found = []
    like = next(r for r, v in sorted(info.numbers.items()) if v == usual)
    right = sheet.cells[(like, col)].value.strip() or _fmt(usual)
    for r in odd:
        value, pos = info.numbers[r], (r, col)
        if value == 0 or abs(value - usual) > 0.25 * abs(usual) or not sheet.dependents.get(pos):
            continue
        now = sheet.cells[pos].value.strip() or _fmt(value)
        effect = kategorien._effects(sheet, {pos: usual})
        found.append(Finding(
            f"rate:{address(pos)}",
            T(f"{address(pos)} ({header}) {now} weicht ab: In den übrigen {n} Zeilen steht {right}. Ist ein anderer "
              f"Satz für genau diese Zeile nicht Absicht, ist es wohl der alte Satz oder vertippt.{effect} Richtig wäre "
              f"{right} statt {now}.",
              f"{address(pos)} ({header}) {now} is different: the other {n} rows have {right}. Unless a different rate "
              f"for exactly this row is intended, it is probably the old rate or a typo.{effect} Should be {right} "
              f"instead of {now}."),
            "pruefen", T(f"Korrigiert in {address(pos)}: {right}. Gilt für diese Zeile wirklich ein anderer Satz, lass es "
                         "so und schreib eine Notiz in die Zelle.",
                         f"Corrected in {address(pos)}: {right}. If a different rate really applies to this row, leave it "
                         "and write a note in the cell."),
            fix=einsetzen.values(sheet, {pos: usual}, (like, col), right)))
    return found


def _codes(sheet: Sheet, col: int, header: str, info: _Info) -> list[Finding]:
    if info.kind != "zahl":
        return []
    found = []
    for code in sorted(codes_in(list(info.numbers.values()))):
        where = [r for r, v in info.numbers.items() if v == code]
        found.append(Finding(
            f"code:{address((where[0], col))}",
            T(f"In {header} steht {len(where)}-mal {_fmt(code)} ({_cells(col, where)}), weit weg von den übrigen Werten. "
              f"Das sieht nach einem Code für „fehlt“ oder „unbekannt“ aus, aber MITTELWERT und SUMME rechnen damit.",
              f"In {header}, {_fmt(code)} appears {len(where)} times ({_cells(col, where)}), far from the other values. "
              f"That looks like a code for “missing” or “unknown”, but AVERAGE and SUM calculate with it."),
            "fehler", T(f"Lass die Zellen leer oder nimm die {_fmt(code)} vor dem Rechnen heraus (z. B. MITTELWERTWENN(…; \"<>{_fmt(code)}\")).",
                        f"Leave the cells empty or take the {_fmt(code)} out before calculating (e.g. AVERAGEIF(…, \"<>{_fmt(code)}\"))."),
            fix=einsetzen.cleared(sheet, [(r, col) for r in where]),
        ))
    return found[:1]


def _impossible(sheet: Sheet, col: int, header: str, info: _Info) -> list[Finding]:
    if info.kind != "zahl" or len(info.numbers) < 3:
        return []
    name = header.lower()
    if re.search(_DIFFERENCE_WORDS, name) and _differences(sheet, col, list(info.numbers)):
        return []  # Mengeneffekt, Fehler %, Abweichung: eine Differenz darf negativ sein, eine Dauer nicht
    skip = codes_in(list(info.numbers.values()))
    values = {r: v for r, v in info.numbers.items() if v not in skip}
    for pattern, low, high, what in _LIMITS:
        if re.search(pattern, name):
            if "%" in name or "anteil" in name or "prozent" in name:
                break
            bad = [r for r, v in values.items() if not low <= v <= high]
            if bad and len(bad) <= max(3, len(values) // 10):
                return [impossible_hint(sheet, col, header, values, bad, T(f"{what[0]} gibt es nicht", f"there is no {what[1]}"))]
            return []
    share = re.search(r"%|prozent|anteil|quote|percent|" + woerter.SHARE, name)
    fraction = sum(1 for v in values.values() if abs(v) <= 1) >= 0.8 * len(values)
    if (share or fraction and re.search(woerter.SHARE_IF_FRACTIONS, name)) and not any(w in name for w in _CHANGE_WORDS):
        top = 1.0 if fraction else 100.0
        bad = [r for r, v in values.items() if v < 0 or v > top * (1 + 1e-9)]
        if bad and len(bad) <= max(3, len(values) // 10):
            limit = _fmt(top * 100 if fraction else top)
            return [impossible_hint(sheet, col, header, values, bad, T(f"einen Anteil unter 0 oder über {limit} % gibt es nicht",
                                                                       f"there is no share below 0 or above {limit} %"))]
        return []
    if re.search(_NOT_NEGATIVE, name) and not any(w in name for w in _CHANGE_WORDS):
        bad = [r for r, v in values.items() if v < 0]
        if bad and len(bad) <= max(3, len(values) // 10):
            return [impossible_hint(sheet, col, header, values, bad, T("eine negative Zahl passt hier nicht",
                                                                       "a negative number does not fit here"))]
    return []


# Spalten, die eine Veränderung oder Abweichung zeigen: Sie dürfen negativ sein.
_DIFFERENCE_WORDS = (r"effekt|effect|abweichung|differenz|delta|fehler|error|veränderung|veraenderung|änderung|change|"
                     r"varianz|variance|bias|écart|ecart|scostamento|errore|erreur")


def _differences(sheet: Sheet, col: int, rows: list[int]) -> bool:
    """Rechnet die Spalte eine Differenz (=(D2-B2)*E2, =(C2-B2)/C2, =G2-F2)? Dann ist ein Minus kein Fehler."""
    def minus(node) -> bool:
        if node is None:
            return False
        if node.kind == "op" and node.value == "-" and len(node.args) == 2 and node.args[1].kind != "num":
            return True
        if node.kind == "op" and node.value in ("*", "/") and len(node.args) == 2:
            return minus(node.args[0]) or (node.value == "*" and minus(node.args[1]))
        return False

    formulas = [sheet.nodes.get((r, col)) for r in rows if (r, col) in sheet.nodes]
    hits = sum(minus(n) for n in formulas)
    # Eine Zeile wie MAPE oder Summe darunter rechnet anders; die Spalte bleibt eine Differenz.
    return hits >= max(2, len(rows) // 2) and 3 * hits >= 2 * len(formulas)


def impossible_hint(sheet: Sheet, col: int, header: str, values: dict[int, float], bad: list[int], why: str) -> Finding:
    # Für folgen.py: Ist ein unmöglicher BMI nur die Folge einer falschen Größe, fällt seine Zeile hier weg.
    sheet.memo.setdefault("auswertung.bad", {})[f"impossible:{address((bad[0], col))}"] = (col, header, values, bad, why)
    listed = ", ".join(f"{_fmt(values[r])} in {address((r, col))}" for r in bad[:MAX_LISTED])
    more = _more(len(bad) - MAX_LISTED) if len(bad) > MAX_LISTED else ""
    return Finding(
        f"impossible:{address((bad[0], col))}",
        T(f"In {header} steht {listed}{more}: {why[0].upper()}{why[1:]}.",
          f"In {header}: {listed}{more}. {why[0].upper()}{why[1:]}."),
        "fehler", T("Ein Tippfehler, eine andere Einheit oder ein Code? Korrigier den Wert oder lass ihn für die Auswertung weg.",
                    "A typo, a different unit or a code? Correct the value or leave it out of the analysis."),
    )


def _units(sheet: Sheet, col: int, header: str, info: _Info) -> list[Finding]:
    """Zwei Gruppen von Werten, die genau um 10, 100 oder 1000 auseinanderliegen: zwei Einheiten in einer Spalte."""
    # Nur getippte Werte: Eine Formel (BMI aus Gewicht 999) erbt ihre Ausreißer von dem, womit sie rechnet.
    values = sorted((v, r) for r, v in info.numbers.items()
                    if v > 0 and v not in CODES and not sheet.cells[(r, col)].formula)
    if info.kind != "zahl" or len(values) < 8 or is_naming(header):
        return []
    logs = [math.log10(v) for v, _ in values]
    gap, at = max((logs[i + 1] - logs[i], i) for i in range(len(logs) - 1))
    low, high = logs[: at + 1], logs[at + 1:]
    if gap < 0.7 or min(len(low), len(high)) < 2 or min(len(low), len(high)) < 0.05 * len(logs):
        return []
    if gap < 2 * max(low[-1] - low[0], high[-1] - high[0]):
        return []
    ratio = statistics.median(high) - statistics.median(low)
    name = header.lower()
    glucose = re.search(r"glu[ck]ose|blutzucker|\bbz\b", name) and abs(10 ** ratio - 18) < 2.5
    if abs(ratio - round(ratio)) > 0.12 and not glucose:
        return []
    factor = 10 ** ratio if glucose else 10.0 ** round(ratio)
    # Umgerechnet decken beide Einheiten dieselben Werte ab (1,62 m bis 1,90 m wie 158 cm bis 195 cm).
    if values[at][0] * factor < values[at + 1][0] or values[0][0] * factor > values[-1][0]:
        return []
    small = [r for v, r in values[: at + 1]]
    big = [r for v, r in values[at + 1:]]
    few, many = (sorted(small), big) if len(small) <= len(big) else (sorted(big), small)
    hint = T("mmol/l und mg/dl", " (mmol/l and mg/dl)") if glucose else _unit_guess(name, round(ratio))
    high_, low_ = _fmt(statistics.median(info.numbers[r] for r in many)), _fmt(statistics.median(info.numbers[r] for r in few))
    factor = _fmt(round(10 ** ratio, 1) if glucose else 10.0 ** round(ratio))
    return [Finding(
        f"units:{address((few[0], col))}",
        T(f"In {header} liegen die Werte in zwei Gruppen, um {high_} und um {low_} ({_cells(col, few)}), um den Faktor "
          f"{factor} auseinander. Das sieht nach zwei Einheiten aus{hint}.",
          f"In {header}, the values fall into two groups, around {high_} and around {low_} ({_cells(col, few)}), a factor "
          f"of {factor} apart. That looks like two units{hint}."),
        "pruefen", T("Rechne alle Werte in dieselbe Einheit um, bevor du Mittelwerte oder Summen bildest.",
                     "Convert all values to the same unit before you calculate averages or sums."),
    )]


def _unit_guess(name: str, power: int) -> str:
    if re.search(r"größe|groesse|länge|height|\bcm\b|\bm\b", name) and power == 2:
        return T(" (m und cm)", " (m and cm)")
    if re.search(r"gewicht|weight|\bkg\b|\bg\b", name) and power == 3:
        return T(" (kg und g)", " (kg and g)")
    if re.search(r"kosten|betrag|preis|chf|eur|umsatz", name) and power == 3:
        return T(" (Franken und Tausend Franken)", " (francs and thousands of francs)")
    if power == 2 and re.search(r"%|anteil|quote|prozent", name):
        return T(" (Bruch und Prozent)", " (fraction and percent)")
    return ""


def _dates(sheet: Sheet, columns: dict[int, _Info]) -> list[Finding]:
    """Ein Ende vor dem Anfang (Austritt vor Eintritt) und Geburtsdaten in der Zukunft oder nach einem Eintritt."""
    plan = sheet.plan
    dated = {c: i for c, i in columns.items() if i.kind == "datum"}
    found = []
    today = (date.today() - date(1899, 12, 30)).days
    starts = [c for c in dated if re.search(_START, plan.headers[c].lower())]
    ends = [c for c in dated if re.search(_END, plan.headers[c].lower())]
    births = [c for c in dated if re.search(_BIRTH, plan.headers[c].lower())]
    pairs = [(s, e) for s in starts for e in ends if s != e]
    pairs += [(b, o) for b in births for o in dated if o != b and o not in births]
    for first, then in pairs:
        a, b = dated[first].dates, dated[then].dates
        bad = [r for r in sorted(a) if r in b and b[r] < a[r] - 1e-9]
        # Nur Uhrzeiten ohne Datum: Eine Nachtschicht endet am nächsten Morgen (22:00 bis 06:30). Rechnet die Dauer
        # daneben das richtig (REST(Ende-Beginn;1)), ist nichts falsch; auffällig ist nur, wo sie negativ wird, oder
        # ohne Dauer eine einzelne Zeile unter lauter anderen.
        clock = bool(bad) and all(0 <= a[r] < 1 and 0 <= b[r] < 1 for r in bad)
        if clock:
            computed = {r for r in bad if _reads_both(sheet, r, first, then)}
            negative = {q[0] for q in _negative_durations(sheet, bad, first, then)}
            alone = [r for r in bad if r not in computed]
            # Rechnen andere Zeilen über Mitternacht richtig, sind die Zeiten gewollt: Falsch ist dann die Formel der
            # negativen Zeile (ohne REST), und das sagt die Prüfung des Formelmusters mit Korrektur.
            bad = [] if computed - negative else [r for r in bad if r in negative or r in alone and len(alone) <= 2]
        if not bad:
            continue
        cells = ", ".join(T(f"Zeile {r}", f"row {r}") for r in bad[:MAX_LISTED]) + (_more(len(bad) - MAX_LISTED) if len(bad) > MAX_LISTED else "")
        follows = _negative_durations(sheet, bad, first, then)
        # Die negative Dauer in derselben Zeile ist die Folge, nicht ein eigener Fehler: Sie steht hier dabei.
        sheet.memo.setdefault("auswertung.order", set()).update(follows)
        more = _more(len(follows) - MAX_LISTED) if len(follows) > MAX_LISTED else ""
        if not follows:
            effect = T(" Dauern daraus werden negativ.", " Durations calculated from it become negative.")
        elif sprache.english():
            effect = (f" That is why {_listed_words([f'{sheet.label(q)} shows {_fmt(sheet.inputs[q])}' for q in follows[:MAX_LISTED]])}"
                      f"{more}.")
        else:
            shown = [f"{sheet.label(q)} {_fmt(sheet.inputs[q])}" for q in follows[:MAX_LISTED]]
            effect = f" Darum {'zeigt' if len(follows) == 1 else 'zeigen'} {_listed_words(shown)}{more}."
        tip = (T("Vertippt, oder geht die Zeit über Mitternacht? Dann rechnet REST(Ende-Beginn;1) die Dauer richtig.",
                 "A typo, or does the time run past midnight? Then MOD(end-start,1) calculates the duration correctly.")
               if clock else T("Sind Tag und Monat vertauscht oder die Spalten verwechselt?",
                               "Are day and month swapped, or the columns mixed up?"))
        key = f"order:{address((bad[0], then))}"
        if follows:
            # Was daraus folgt, nimmt spaltenregeln.py in den Hinweis zu einer Regel für dasselbe Paar.
            sheet.memo.setdefault("auswertung.effect", {})[key] = effect.strip()
        example = (sheet.cells[(bad[0], then)].value, sheet.cells[(bad[0], first)].value)
        found.append(Finding(
            key,
            order_start(plan.headers[then], plan.headers[first]) +
            T(f"({cells}, zum Beispiel {example[0]} vor {example[1]}).{effect}",
              f"({cells}, for example {example[0]} before {example[1]}).{effect}"),
            "fehler", tip,
        ))
        if len(found) >= 2:
            break
    for col in births:
        future = [r for r, v in dated[col].dates.items() if v > today + 1]
        if future:
            found.append(Finding(
                f"future:{address((future[0], col))}",
                T(f"In {plan.headers[col]} liegt {_cells(col, future)} in der Zukunft ({sheet.cells[(future[0], col)].value}).",
                  f"In {plan.headers[col]}, {_cells(col, future)} {'is' if len(future) == 1 else 'are'} in the future "
                  f"({sheet.cells[(future[0], col)].value})."),
                "fehler", T("Ein Tippfehler im Jahr?", "A typo in the year?"),
            ))
    return found


def _only_follows(sheet: Sheet, item: Finding, follows: set[Pos]) -> bool:
    """Sind alle unmöglichen Werte dieses Hinweises Folgen eines Endes vor dem Anfang?"""
    entry = (sheet.memo.get("auswertung.bad") or {}).get(item.key)
    if not entry or not follows:
        return False
    col, bad = entry[0], entry[3]
    return all((r, col) in follows for r in bad)


def _negative_durations(sheet: Sheet, bad: list[int], first: int, then: int) -> list[Pos]:
    """Formeln in den Zeilen mit dem Ende vor dem Anfang, die beides lesen und darum negativ werden (Stunden =
    (Ende-Beginn)*24-Pause)."""
    out = []
    for r in bad:
        both = set(sheet.dependents.get((r, first), ())) & set(sheet.dependents.get((r, then), ()))
        out += sorted(q for q in both if q[0] == r and _negative(sheet.inputs.get(q)))
    return out


def _reads_both(sheet: Sheet, r: int, first: int, then: int) -> bool:
    both = set(sheet.dependents.get((r, first), ())) & set(sheet.dependents.get((r, then), ()))
    return any(q[0] == r for q in both)


def _negative(value) -> bool:
    return isinstance(value, float) and not isinstance(value, bool) and math.isfinite(value) and value < 0


def _listed_words(items: list[str]) -> str:
    return items[0] if len(items) == 1 else ", ".join(items[:-1]) + T(f" und {items[-1]}", f" and {items[-1]}")


def _ids(sheet: Sheet, columns: dict[int, _Info]) -> list[Finding]:
    """Eine Spalte mit Fall- oder Patientennummern, in der fast alles nur einmal vorkommt, bis auf wenige."""
    plan = sheet.plan
    found = []
    for col, info in columns.items():
        header = plan.headers[col]
        if not is_id(header):
            continue
        values = {r: (sheet.cells[(r, col)].value.strip().casefold()) for r in list(info.numbers) + list(info.texts)}
        if len(values) < MIN_ROWS:
            continue
        counts = Counter(values.values())
        if len(counts) < 0.9 * len(values):
            continue
        twice = [v for v, k in counts.items() if k > 1]
        if not twice:
            continue
        person = re.search(r"patient|person|versicherten|\bpid\b|\bpat\b|kunde|mitglied|" + woerter.PERSON, header.lower())
        # In einer Liste von Behandlungen, Besuchen oder Buchungen (mit Datum) kommt dieselbe Person mit Absicht
        # mehrmals vor. Nur eine Fall- oder Rechnungsnummer muss dort einmalig sein.
        if person and not re.search(r"fall|rechnung|beleg|auftrag|" + woerter.UNIQUE, header.lower()) and \
                any(other.kind == "datum" for c, other in columns.items() if c != col):
            continue
        # Zählen die Nummern Zeile für Zeile hoch, sagt die Lücke, welche Nummer die doppelte sein sollte.
        repaired = _renumbered(sheet, col, sorted(values), {r for r, x in values.items() if x in twice})
        for r, (other, right, put) in sorted(repaired.items())[:MAX_LISTED]:
            pos, now = (r, col), sheet.cells[(r, col)].value.strip()
            fix = (einsetzen.values(sheet, {pos: put}, (other, col), right) if isinstance(put, float) else
                   kategorien._fix(sheet, {pos: put}))
            found.append(Finding(
                f"id:{address(pos)}",
                T(f"{address(pos)} ({header}) „{now}“ steht schon in {address((other, col))}. Die Nummern zählen sonst "
                  f"Zeile für Zeile hoch, und „{right}“ fehlt: wohl kopiert und nicht angepasst. Eine Suche nach der "
                  f"Nummer (SVERWEIS, XVERWEIS) findet nur die erste der beiden Zeilen. Richtig wäre „{right}“ statt "
                  f"„{now}“.",
                  f"{address(pos)} ({header}) “{now}” is already in {address((other, col))}. Otherwise the numbers count "
                  f"up row by row, and “{right}” is missing: probably copied and not adjusted. A lookup of the number "
                  f"(VLOOKUP, XLOOKUP) finds only the first of the two rows. Should be “{right}” instead of “{now}”."),
                "fehler", T(f"Korrigiert in {address(pos)}: {right}.", f"Corrected in {address(pos)}: {right}."), fix=fix))
        twice = [v for v in twice if not any(values[r] == v for r in repaired)]
        if not twice:
            continue
        listed = []
        for v in twice[:MAX_LISTED]:
            where = sorted(r for r, x in values.items() if x == v)
            value = sheet.cells[(where[0], col)].value
            listed.append(T(f"„{value}“ in Zeilen {', '.join(map(str, where[:4]))}",
                            f"“{value}” in rows {', '.join(map(str, where[:4]))}"))
        first = min(r for r, x in values.items() if x == twice[0])
        found.append(Finding(
            f"id:{address((first, col))}",
            T(f"In {header} kommt fast jede Nummer einmal vor, {len(twice)} aber mehrmals: {'; '.join(listed)}"
              f"{' und weitere' if len(twice) > MAX_LISTED else ''}.",
              f"In {header}, almost every number appears once, but {len(twice)} "
              f"{'appears' if len(twice) == 1 else 'appear'} more than once: {'; '.join(listed)}"
              f"{' and more' if len(twice) > MAX_LISTED else ''}."),
            "pruefen", T("Ist das dieselbe Person mit mehreren Fällen, oder derselbe Fall doppelt erfasst?",
                         "Is this the same person with several cases, or the same case entered twice?") if person
            else T("Ist derselbe Fall doppelt erfasst? Dann zählt er in Summen und Anzahlen zweimal.",
                   "Is the same case entered twice? Then it counts twice in sums and counts."),
        ))
    return found[:4]


_NUMBERED = re.compile(r"(.*?)(\d+)(\D*)")
_BLANK = Cell("")


def _renumbered(sheet: Sheet, col: int, rows: list[int], doubled: set[int]) -> dict[int, tuple[int, str, str | float]]:
    """Doppelte Nummern in einer Spalte, die Zeile für Zeile um eins hochzählt (A-2008, A-2009, A-2010): Wo eine
    Nummer aus der Reihe fällt, die schon weiter oben steht, und die Nummer, die in die Reihe gehört, nirgends steht,
    ist es diese. Zeile → (wo die Nummer schon steht, die richtige Nummer, was einzusetzen ist)."""
    parts = {}
    for r in rows:
        cell = sheet.cells.get((r, col))
        match = _NUMBERED.fullmatch(cell.value.strip()) if cell is not None and cell.formula is None else None
        if match:
            digits = match.group(2)
            # Mit Null vorn („0042“) hat die Nummer eine feste Länge, sonst nicht (99, 100).
            parts[r] = (match.group(1), int(digits), len(digits) if digits[0] == "0" else 0, match.group(3))

    def same(a, b) -> bool:
        return a in parts and b in parts and (parts[a][0], parts[a][3]) == (parts[b][0], parts[b][3])

    steps = [parts[b][1] - parts[a][1] for a, b in zip(rows, rows[1:]) if same(a, b)]
    if len(steps) < MIN_ROWS or sum(1 for d in steps if d == 1) < 0.8 * len(steps):
        return {}
    taken = {(p[0], p[1], p[3]) for p in parts.values()}
    out = {}
    for i, r in enumerate(rows):
        if r not in doubled or r not in parts:
            continue
        before, after = (rows[i - 1] if i else None), (rows[i + 1] if i + 1 < len(rows) else None)
        near = [n for n in (before, after) if same(n, r)]
        if not near:
            continue
        if before in near and after in near and parts[after][1] - parts[before][1] != 2:
            continue
        want = parts[before][1] + 1 if before in near else parts[after][1] - 1
        prefix, number, _, suffix = parts[r]
        width = parts[near[0]][2]
        if want == number or want < 0 or (prefix, want, suffix) in taken:
            continue
        other = next((o for o in rows if o != r and o in parts and parts[o] == parts[r]), None)
        if other is None or other > r:
            continue
        # Ist die ganze Zeile gleich, ist sie doppelt eingefügt (_same_rows): Eine neue Nummer machte es schlimmer.
        if all((sheet.cells.get((r, c)) or _BLANK).value == (sheet.cells.get((other, c)) or _BLANK).value
               for c in sheet.plan.headers):
            continue
        right = f"{prefix}{want:0{width}d}{suffix}"
        cell = sheet.cells[(r, col)]
        numeric = isinstance(sheet.inputs.get((r, col)), float) and not (prefix or suffix) and cell.text is None
        out[r] = (other, right, float(want) if numeric else right)
    return out


def _same_source(sheet: Sheet, a: int, b: int, columns: list[int]) -> bool:
    filled = [c for c in columns if c != sheet.plan.name_column and ((a, c) in sheet.cells or (b, c) in sheet.cells)]
    return bool(filled) and all(
        (a, c) in sheet.cells and (b, c) in sheet.cells and sheet.cells[(a, c)].formula and sheet.cells[(b, c)].formula
        and sheet.precedents.get((a, c)) == sheet.precedents.get((b, c)) for c in filled
        if (sheet.cells.get((a, c)) and sheet.cells[(a, c)].formula) or (sheet.cells.get((b, c)) and sheet.cells[(b, c)].formula))


def _same_rows(sheet: Sheet, rows: list[int]) -> list[Finding]:
    """Zwei Zeilen, in denen alles genau gleich ist: oft doppelt eingefügt."""
    columns = sorted(sheet.plan.headers)
    if len(columns) < 4 or len(rows) < MIN_ROWS:
        return []
    # Nur wenn eine Spalte die Zeilen unterscheidet (Nummer, Messwert): sonst sind gleiche Zeilen normal.
    if not any(len({sheet.cells[(r, c)].value for r in rows if (r, c) in sheet.cells}) >= 0.9 * len(rows) for c in columns):
        return []
    seen: dict[tuple, int] = {}
    pairs = []
    for r in rows:
        key = tuple(sheet.cells[(r, c)].value if (r, c) in sheet.cells else "" for c in columns)
        if sum(1 for v in key if v) < 3:
            continue
        if key in seen:
            # =C29 in jeder Spalte übernimmt die Zeile mit Absicht (Jahresüberschuss im Cashflow). Rechnen beide
            # Zeilen mit denselben Formeln aus denselben Zellen (Abschreibungen im Cashflow und im Free Cash Flow),
            # ist es derselbe Posten an zwei Stellen, nicht eingefügt.
            if not links(sheet, r, seen[key]) and not _same_source(sheet, r, seen[key], columns):
                pairs.append((seen[key], r))
        else:
            seen[key] = r
    if not pairs:
        return []
    if len(pairs) == 1:
        text = T(f"Zeile {pairs[0][1]} ist genau gleich wie Zeile {pairs[0][0]}. Doppelt eingefügt?",
                 f"Row {pairs[0][1]} is exactly the same as row {pairs[0][0]}. Pasted twice?")
    else:
        listed = ", ".join(T(f"{b} wie {a}", f"{b} like {a}") for a, b in pairs[:MAX_LISTED])
        text = T(f"{len(pairs)} Zeilen sind genau gleich wie eine frühere (Zeile {listed}{' und weitere' if len(pairs) > MAX_LISTED else ''}). Doppelt eingefügt?",
                 f"{len(pairs)} rows are exactly the same as an earlier one (row {listed}{' and more' if len(pairs) > MAX_LISTED else ''}). Pasted twice?")
    return [Finding(
        f"same:{address((pairs[0][1], columns[0]))}",
        text,
        "pruefen", T("Wenn es derselbe Fall ist, lösch die zweite Zeile, sonst zählt er doppelt.",
                     "If it is the same case, delete the second row; otherwise it counts twice."),
    )]


def _filled_down(sheet: Sheet, pos: Pos) -> bool:
    """Steht darüber oder darunter dieselbe Formel, heruntergezogen (eine laufende Summe)? Dann endet ihr Bereich
    nicht mit Absicht bei einer Gruppe."""
    from .workbook import pattern

    shape = pattern(sheet.cells[pos].formula or "", pos)
    return any((near := sheet.cells.get((row, pos[1]))) is not None and near.formula
               and pattern(near.formula, (row, pos[1])) == shape for row in (pos[0] - 1, pos[0] + 1))


def _moves(sheet: Sheet, pos: Pos, area: tuple) -> bool:
    """Ist der Bereich mit der Formel heruntergezogen und sein Ende ohne $ vor der Zeile, rutscht es Zeile für Zeile mit."""
    shown = bereiche._kept(sheet.cells[pos].formula or "").get(("range", area), "")
    end = shown.rpartition(":")[2]
    return bool(end) and not re.search(r"\$\d", end) and _filled_down(sheet, pos)


def _also(node: formel.Node, call: formel.Node, col: int, area: tuple) -> set[int]:
    """Zeilen der Spalte col, die neben dem Bereich area auch noch dazukommen: weitere Teile desselben Aufrufs
    (MITTELWERT(D2:D12;D13)) und bei SUMME, was daneben addiert wird (SUMME(D2:D12)+D13). Zeilen im Bereich selbst
    zählen nicht: Bei einer Laufsumme in jeder Zeile wären das sonst Tausende je Formel. Beim Kapitalwert ebenso:
    =B2+NBW(B9;B3:B7) nimmt Jahr 0 unabgezinst dazu, so gehört es."""
    parts = list(call.args)
    if call.value in ("SUM", "NPV"):
        terms = lambda n: terms(n.args[0]) + terms(n.args[1]) if n.kind == "op" and n.value == "+" else [n]  # noqa: E731
        top = terms(node)
        if any(t is call for t in top):
            parts += [t for t in top if t is not call]
    rows: set[int] = set()
    for part in parts:
        if part.kind == "ref" and part.value[1] == col:
            rows.add(part.value[0])
        elif part.kind == "range" and part.value[1] == part.value[3] == col:
            rows.update(range(part.value[0], min(part.value[2], area[0] - 1) + 1))
            rows.update(range(max(part.value[0], area[2] + 1), part.value[2] + 1))
    return rows


def _longer(sheet: Sheet, pos: Pos, call: formel.Node, area: tuple, wanted: tuple[int, int], rows: set[int],
            also: set[int]) -> dict | None:
    """Die Korrektur für einen zu kurzen Bereich: bis zur letzten Datenzeile gezogen. Nur wenn dazwischen nichts
    anderes steht (eine Zwischensumme), nichts doppelt zählt (MITTELWERT(D2:D10;D12)) und die Formel nicht
    in ihren eigenen Bereich gerät."""
    r1, c1, r2, _ = area
    start, end = wanted
    added = [r for r in range(start, end + 1) if not r1 <= r <= r2]
    if any(r in also or ((r, c1) in sheet.cells and r not in rows) for r in added):
        return None
    # Heruntergezogen: die ganze Spalte gleich, sonst stünden darin zwei verschiedene Formeln.
    places = _block(sheet, pos) if _filled_down(sheet, pos) else [pos]
    if any(area not in sheet.ranges.get(p, ()) for p in places):
        return None
    sides = [a.value for a in call.args if a.kind == "range" and (a.value[0], a.value[2]) == (r1, r2)]
    if any(start <= p[0] <= end and left <= p[1] <= right for p in places for _, left, _, right in sides):
        return None
    return einsetzen.stretched(sheet, places, call.value, area, wanted)


def _block(sheet: Sheet, pos: Pos) -> list[Pos]:
    """Die Zellen über und unter pos mit derselben Formel, heruntergezogen."""
    from .workbook import pattern

    shape = pattern(sheet.cells[pos].formula or "", pos)

    def same(row: int) -> bool:
        near = sheet.cells.get((row, pos[1]))
        return near is not None and bool(near.formula) and pattern(near.formula, (row, pos[1])) == shape

    # Höchstens eine Zelle mehr, als ein Klick schreibt: Ist die Spalte länger, wird keine Korrektur daraus.
    top, bottom = pos[0], pos[0]
    while top > 1 and same(top - 1) and bottom - top < einsetzen.MAX_CELLS:
        top -= 1
    while same(bottom + 1) and bottom - top < einsetzen.MAX_CELLS:
        bottom += 1
    return [(row, pos[1]) for row in range(top, bottom + 1)]


def _without_summary(sheet: Sheet, col: int, data: list[int]) -> list[int]:
    """Die Zahlen einer Spalte ohne die Kennzahlen darunter: Mittelwert, Standardabweichung, n und ein Intervall
    daraus („95%-KI … =B28-1,96*B29/WURZEL(B30)“) rechnen mit anderen Zeilen derselben Spalte. Sie sind Ergebnisse,
    keine Daten, auch wenn sie in Zeilen mit einem Namen vorne stehen."""
    from .workbook import pattern

    while data:
        node = sheet.nodes.get((data[-1], col))
        if node is None:
            break
        # Eine heruntergezogene Spalte (Laufsumme =C40+B41) ist kein Kennzahl-Block: dort hat jede Zeile ihre Form.
        above = sheet.cells.get((data[-1] - 1, col))
        here = sheet.cells[(data[-1], col)].formula
        if above is not None and above.formula and here and \
                pattern(above.formula, (data[-1] - 1, col)) == pattern(here, (data[-1], col)):
            break
        cells, areas, _ = formel.references(node)
        if not (any(c == col and r != data[-1] for r, c in cells)
                or any(r1 < r2 and not r1 <= data[-1] <= r2 for r1, _, r2, _ in areas)):
            break
        data = data[:-1]
    return data


def _ranges(sheet: Sheet, rows: list[int]) -> list[Finding]:
    """MITTELWERT, MEDIAN, ANZAHL und Co. über einen Bereich, der Datenzeilen auslässt oder die Summenzeile mitnimmt."""
    plan = sheet.plan
    found = []
    filled: dict[int, list[int]] = {}
    for col in plan.headers:
        filled[col] = _without_summary(sheet, col, [r for r in rows if isinstance(sheet.inputs.get((r, col)), float)
                                                    and (r, col) in sheet.cells])
    # Die Zeilen sind sortiert, dazu als Menge: Bei Tausenden Formeln je Spalte zählt es so die Zeilen außerhalb,
    # ohne sie durchzugehen.
    known = {col: set(data) for col, data in filled.items()}
    data_rows = set(rows)
    totals: dict[int, list[int]] = {}
    # Heruntergezogene gleiche Formeln mit demselben zu kurzen Bereich sind ein Fehler, ein Hinweis: Sonst füllen
    # zehn gleiche Hinweise die Liste, und ein anderer Fehler weiter unten fällt hinten heraus.
    covered: set[Pos] = set()
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown or pos in covered:
            continue
        for call in formel.calls(node):
            if call.value not in _AGGREGATES and call.value not in _CASHFLOWS:
                continue
            for arg in call.args:
                if arg.kind != "range":
                    continue
                r1, c1, r2, c2 = arg.value
                if c1 != c2 or c1 not in plan.headers or len(filled[c1]) < 3 or r2 >= formel.MAX_ROW:
                    continue
                # Eine SUMME direkt unter ihrer Spalte prüft schon task.py („zählt … nicht mit“).
                if call.value == "SUM" and pos[1] == c1:
                    continue
                data = filled[c1]
                if not (r1 <= data[-1] and r2 >= data[0]) or r1 > plan.header_row + 2:
                    continue
                also = _also(node, call, c1, arg.value)
                lo, hi = bisect_left(data, r1), bisect_right(data, r2)
                outside = lo + len(data) - hi - sum(1 for r in also if r in known[c1] and not r1 <= r <= r2)
                # Wenige fehlende Zeilen am Ende: meist kamen Daten dazu, und der Bereich wurde nicht nachgezogen.
                # Heruntergezogen und das Ende rutscht mit (laufende Summe D$2:D42, gleitender Mittelwert): Der
                # Bereich endet mit Absicht in der eigenen Zeile.
                # Fehlt nur die erste Zeile, und die ist eine Anfangszeile (Jahr 0 mit der Investition, ein
                # Anfangsbestand: fast leer gegenüber den Zeilen darunter), lässt der Bereich sie mit Absicht aus.
                # Der interne Zinsfuß braucht Jahr 0 aber immer: ohne die Investition gibt es keinen.
                opening = hi == len(data) and lo and call.value not in _NEEDS_START \
                    and all(_opening(sheet, r, data[lo:hi], plan.headers) for r in data[:lo])
                if outside and r1 <= data[0] + 1 and outside <= max(3, len(data) // 5) and not _moves(sheet, pos, arg.value) \
                        and not opening:
                    left = [r for r in data[:lo] + data[hi:] if r not in also]
                    # Bis zur letzten Zeile der Tabelle, auch wenn dort eine Zahl als Text steht (die zählt nach dem
                    # Umwandeln mit).
                    end = max(r for r in rows if (r, c1) in sheet.cells)
                    fix = _longer(sheet, pos, call, arg.value, (min(r1, data[0]), max(r2, end)), set(rows), also)
                    same = [p for p in (_block(sheet, pos) if _filled_down(sheet, pos) else [pos])
                            if arg.value in sheet.ranges.get(p, ())]
                    covered.update(same)
                    more = (T(f" Ebenso die {len(same) - 1} gleichen Formeln bis {address(same[-1])}.",
                              f" Likewise the {len(same) - 1} identical formulas down to {address(same[-1])}.")
                            if len(same) > 1 else "")
                    span, whole = f"{address((r1, c1))}:{address((r2, c2))}", f"{column_letters(c1)}:{column_letters(c1)}"
                    found.append(Finding(
                        f"range:{address(pos)}",
                        T(f"{_german(call.value)} in {sheet.figure_label(pos)} nimmt {span}, in "
                          f"{plan.headers[c1]} stehen Zahlen aber bis {address((data[-1], c1))}: {len(left)} "
                          f"{'Zeile fehlt' if len(left) == 1 else 'Zeilen fehlen'} ({_cells(c1, left)}).{more}",
                          f"{call.value} in {sheet.figure_label(pos)} takes {span}, but in {plan.headers[c1]} there are "
                          f"numbers down to {address((data[-1], c1))}: {len(left)} "
                          f"{'row is' if len(left) == 1 else 'rows are'} missing ({_cells(c1, left)}).{more}"),
                        "fehler", T(f"Zieh den Bereich bis {address((end, c1))}, oder nimm die ganze Spalte "
                                    f"({whole}), wenn darunter nichts steht.",
                                    f"Extend the range to {address((end, c1))}, or take the whole column ({whole}) if "
                                    "nothing is below it."),
                        fix=fix,
                    ))
                elif outside and not _filled_down(sheet, pos) and (group := bereiche.group(sheet, rows, r1, r2, c1)) \
                        and not (also.issuperset(r for r in range(group[2], group[3] + 1) if not r1 <= r <= r2)
                                 and group[2] <= r1 and r2 <= group[3]):
                    found.append(bereiche.group_finding(sheet, pos, call.value, (r1, c1, r2, c2), *group))
                if call.value == "SUM":
                    continue
                if c1 not in totals:
                    totals[c1] = [r for r in range(plan.header_row + 1, plan.last_row + 1)
                                  if r not in data_rows and _aggregate(sheet.nodes.get((r, c1)))]
                among = totals[c1]
                # TEILERGEBNIS über TEILERGEBNIS: Excel lässt die inneren weg (Daten → Teilergebnis).
                inside = [r for r in among[bisect_left(among, max(r1, plan.header_row + 1)):bisect_right(among, min(r2, plan.last_row))]
                          if (r, c1) != pos and not (call.value in _SKIPS_NESTED and sheet.nested((r, c1)))]
                if inside:
                    found.append(Finding(
                        f"total:{address(pos)}",
                        T(f"{_german(call.value)} in {sheet.figure_label(pos)} nimmt {address((inside[0], c1))} mit, und dort steht "
                          f"selbst schon eine Summe oder ein Mittelwert.",
                          f"{call.value} in {sheet.figure_label(pos)} includes {address((inside[0], c1))}, and that cell "
                          f"already holds a sum or an average itself."),
                        "fehler", T(f"Lass den Bereich vor {address((inside[0], c1))} enden.",
                                    f"Let the range end before {address((inside[0], c1))}."),
                    ))
    return found[:MAX_RANGES]


MAX_RANGES = 6
# Über die Jahre einer Spalte: Kapitalwert und interner Zinsfuß lassen sonst still ein Jahr aus.
_CASHFLOWS = {"NPV", "IRR", "MIRR"}
_NEEDS_START = {"IRR", "MIRR"}


def _opening(sheet: Sheet, row: int, others: list[int], columns) -> bool:
    """Ist row eine Anfangszeile vor den Zeilen others: Jahr 0 mit Investition und Darlehen, ein Anfangsbestand?
    Sie füllt höchstens halb so viele Spalten der Tabelle wie die Zeilen darunter und hat dabei einen Betrag. Eine
    Zeile mit nur einem Text ist eine Datenzeile mit Lücken."""
    def filled(r: int) -> int:
        return sum(1 for c in columns if (cell := sheet.cells.get((r, c))) is not None and (cell.formula or cell.value not in (None, "")))

    if len(others) < 3 or not any(isinstance(sheet.inputs.get((row, c)), float) and (row, c) in sheet.cells for c in columns):
        return False
    counts = sorted(filled(r) for r in others)
    return filled(row) * 2 <= counts[len(counts) // 2]


# Größer prüft der Prüfer ein anderes Blatt für einen Bereich darauf nicht noch einmal ganz durch.
MAX_OTHER_CELLS = 60_000
_CONDITIONAL = {"SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS", "MAXIFS", "MINIFS"}
_SKIPS_NESTED = {"SUBTOTAL", "AGGREGATE"}


def _other_table(sheet: Sheet, name: str):
    """Das andere Blatt name als geprüftes Blatt mit seinen Datenzeilen, oder None (unbekannt, zu groß, ohne Tabelle).
    Einmal je Mappe: Die Zellen eines gelesenen Blatts ändern sich nicht mehr (xlsx.Cells.memo)."""
    from .workbook import layout
    from .xlsx import Sheet as Plain

    cells = sheet.book.get(name.casefold())
    if cells is None or len(cells) > MAX_OTHER_CELLS:
        return None
    memo = getattr(cells, "memo", None)
    if memo is not None and ("auswertung", "table") in memo:
        return memo[("auswertung", "table")]
    if ("auswertung", "other", name.casefold()) in sheet.memo:
        return sheet.memo[("auswertung", "other", name.casefold())]
    other = Sheet(cells, layout(Plain(name, cells)), name, {})
    done = (other, _data_rows(other)) if other.plan.header_row is not None else None
    if memo is not None:
        memo[("auswertung", "table")] = done
    sheet.memo[("auswertung", "other", name.casefold())] = done
    return done


def _other_column(other: Sheet, rows: list[int], col: int) -> tuple[list[int], list[int]]:
    """Die Datenzeilen mit einem Wert in der Spalte col des anderen Blatts und die Zeilen darin, die selbst
    zusammenfassen (Summe, Mittelwert). Einmal je Spalte: Jeder Bereich darauf fragt dasselbe."""
    key = ("auswertung", "column", col)
    if key not in other.memo:
        data = [r for r in rows if (r, col) in other.cells and other.shown.get((r, col)) is not None]
        known = set(rows)
        among = [r for r in range(other.plan.header_row + 1, max(other.extent(None)[0], other.plan.header_row) + 1)
                 if r not in known and _aggregate(other.nodes.get((r, col)))]
        other.memo[key] = (data, among)
    return other.memo[key]


def _sheet_ref(name: str, text: str) -> str:
    return f"{name}!{text}" if re.fullmatch(r"[A-Za-zÄÖÜäöüß_][\wÄÖÜäöüß.]*", name) else f"'{name}'!{text}"


def _other_ranges(sheet: Sheet) -> list[Finding]:
    """Wie _ranges, für Bereiche auf einem anderen Blatt (MITTELWERT(Fälle!I4:I69) auf dem Blatt „Auswertung“):
    Der Bereich lässt die letzten Datenzeilen aus, oder er nimmt die Summenzeile darunter mit. Gleiche Formeln einer
    Spalte (ZÄHLENWENN(Fälle!$D$2:$D$4001;A2) für jede Station) und gleich hohe Bereiche einer Formel
    (SUMMEWENN(Fälle!$D$2:$D$4001;A2;Fälle!$F$2:$F$4001)) bekommen einen Hinweis zusammen."""
    from .workbook import pattern

    found: list[Finding] = []
    seen: set[tuple] = set()
    short: dict[tuple, dict] = {}
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or pos in sheet.unknown:
            continue
        for call in formel.calls(node):
            if call.value not in _AGGREGATES and call.value not in _CONDITIONAL:
                continue
            for arg in call.args:
                if arg.kind != "xrange":
                    continue
                name, (r1, c1, r2, c2) = arg.value
                done = _other_table(sheet, name) if c1 == c2 and r2 < formel.MAX_ROW else None
                if done is None:
                    continue
                other, rows = done
                plan = other.plan
                data, among = _other_column(other, rows, c1)
                if c1 not in plan.headers or len(data) < 3 or r1 > plan.header_row + 2 or not r1 <= data[-1]:
                    continue
                left = [r for r in data if r > r2]
                # Heruntergezogen mit mitrutschendem Ende ist Absicht (laufende Summe über das andere Blatt).
                moving = _filled_down(sheet, pos) and not re.search(r":\$?[A-Z]+\$\d", sheet.cells[pos].formula or "")
                if left and r1 <= data[0] + 1 and len(left) <= max(3, len(data) // 5) and not moving:
                    key = (pos[1], pattern(sheet.cells[pos].formula or "", pos), name.casefold(), r1, r2)
                    entry = short.setdefault(key, {"pos": pos, "call": call, "name": name, "area": (r1, c1, r2, c2),
                                                   "end": data[-1], "other": other, "rows": rows, "among": among,
                                                   "left": left, "columns": [], "members": []})
                    if pos == entry["pos"] and c1 not in entry["columns"]:
                        entry["columns"].append(c1)
                    if pos not in entry["members"]:
                        entry["members"].append(pos)
                here = _sheet_ref(name, f"{address((r1, c1))}:{address((r2, c2))}")
                inside = [r for r in among[bisect_left(among, max(r1, plan.header_row + 1)):bisect_right(among, r2)]
                          if not (call.value in _SKIPS_NESTED and other.nested((r, c1)))]
                if inside and ("total", pos) not in seen:
                    seen.add(("total", pos))
                    at = _sheet_ref(name, address((inside[0], c1)))
                    found.append(Finding(
                        f"total:{address(pos)}",
                        T(f"{_german(call.value)} in {sheet.figure_label(pos)} nimmt {here}, und darin steht in {at} "
                          f"schon die Summe (oder ein Mittelwert) der Zeilen darüber. Sie zählt mit, als wäre sie ein Fall.",
                          f"{call.value} in {sheet.figure_label(pos)} takes {here}, and in it {at} already holds the sum "
                          f"(or an average) of the rows above. It counts as if it were a case."),
                        "fehler", T(f"Lass den Bereich vor {at} enden.", f"Let the range end before {at}."),
                    ))
    found += [_short_other(sheet, entry) for entry in short.values()]
    found.sort(key=lambda f: formel.position(f.key.partition(":")[2]))
    return found[:6]


def _short_other(sheet: Sheet, entry: dict) -> Finding:
    """Der Hinweis zu Bereichen auf einem anderen Blatt, die vor den letzten Fällen aufhören: eine Formel oder alle
    gleichen einer Spalte, mit allen gleich hohen Bereichen der Formel."""
    pos, call, name, (r1, c1, r2, c2), end = entry["pos"], entry["call"], entry["name"], entry["area"], entry["end"]
    plan, left, columns, members = entry["other"].plan, entry["left"], entry["columns"], entry["members"]
    and_ = T(" und ", " and ")
    heres = and_.join(_sheet_ref(name, f"{address((r1, c))}:{address((r2, c))}") for c in columns)
    one = len(members) == 1
    lead = f"{_fn(call.value)} in {sheet.figure_label(pos)}"
    if not one:
        lead += T(f" und {len(members) - 1} {'gleicher Formel' if len(members) == 2 else 'gleichen Formeln'} darunter",
                  f" and {len(members) - 1} identical {'formula' if len(members) == 2 else 'formulas'} below")
    lacking = T(f"{len(left)} {'Zeile fehlt' if len(left) == 1 else 'Zeilen fehlen'}",
                f"{len(left)} {'row is' if len(left) == 1 else 'rows are'} missing")
    more = _more(len(left) - MAX_LISTED) if len(left) > MAX_LISTED else ""
    if len(columns) == 1:
        last = _sheet_ref(name, address((end, c1)))
        cells = ", ".join(_sheet_ref(name, address((r, c1))) for r in left[:MAX_LISTED])
        text = T(f"{lead} {'nimmt' if one else 'nehmen'} {heres}, auf „{name}“ stehen in "
                 f"{plan.headers[c1]} aber Werte bis {last}: {lacking} ({cells}{more}).",
                 f"{lead} {'takes' if one else 'take'} {heres}, but on “{name}” {plan.headers[c1]} has values down to "
                 f"{last}: {lacking} ({cells}{more}).")
        tip = T(f"Zieh den Bereich bis {last}.", f"Extend the range to {last}.")
    else:
        heads = and_.join(plan.headers.get(c, column_letters(c)) for c in columns)
        listed = ", ".join(str(r) for r in left[:MAX_LISTED]) + more
        text = T(f"{lead} {'nimmt' if one else 'nehmen'} {heres}, auf „{name}“ gehen {heads} aber bis "
                 f"Zeile {end}: {lacking} ({'Zeile' if len(left) == 1 else 'Zeilen'} {listed}).",
                 f"{lead} {'takes' if one else 'take'} {heres}, but on “{name}” {heads} go down to row {end}: {lacking} "
                 f"({'row' if len(left) == 1 else 'rows'} {listed}).")
        tip = T(f"Zieh die Bereiche bis Zeile {end}.", f"Extend the ranges to row {end}.")
    fix = _longer_other(sheet, pos, call, name, (r1, c1, r2, c2), end, entry["other"], entry["rows"], entry["among"],
                        members)
    return Finding(f"range:{address(pos)}", text, "fehler", tip, fix=fix)


def _longer_other(sheet: Sheet, pos: Pos, call: formel.Node, name: str, area: tuple, end: int, other: Sheet,
                  rows: list[int], among: list[int], members: list[Pos] | None = None) -> dict | None:
    """Wie _longer für einen Bereich auf dem Blatt name: bis zur letzten Datenzeile dort gezogen, wenn dazwischen nur
    Datenzeilen stehen (keine Summenzeile, nichts anderes). members: die gleichen Formeln der Spalte, alle zusammen."""
    r1, c1, r2, _ = area
    known = set(rows)
    if any(r in among or (r not in known and any(p[0] == r for p in other.cells)) for r in range(r2 + 1, end + 1)):
        return None
    places = _block(sheet, pos) if _filled_down(sheet, pos) else [pos]
    places = sorted(set(places) | set(members or ()))
    return einsetzen.stretched(sheet, places, call.value, area, (r1, end), other=name)


def _denominators(sheet: Sheet, covered: set[str]) -> list[Finding]:
    """Anteile, deren Nenner beim Herunterkopieren mitrutscht (C10, C11, … statt C$10). Dass der Nenner
    dann auf leere Zellen zeigt, sagt dieser Hinweis mit (covered)."""
    by_column: dict[int, list[tuple[Pos, Pos]]] = {}
    for pos, node in sorted(sheet.nodes.items()):
        if node is None or node.kind != "op" or node.value != "/" or len(node.args) != 2:
            continue
        top, bottom = node.args
        if bottom.kind != "ref" or top.kind not in ("ref", "call", "op"):
            continue
        by_column.setdefault(pos[1], []).append((pos, bottom.value))
    found = []
    for col, uses in by_column.items():
        if len(uses) < 3:
            continue
        shifts = {d[0] - p[0] for p, d in uses}
        moving = len({d for _, d in uses}) == len(uses)
        if len(shifts) != 1 or 0 in shifts or not moving:
            continue
        (first, base), (second, then) = uses[0], uses[1]
        if not _aggregate(sheet.nodes.get(base)):
            continue
        covered |= {f"emptyref:{address(p)}" for p, _ in uses}
        fixed = _anchored(sheet, [p for p, _ in uses], base)
        key, fix = f"slide:{address(first)}", None
        if fixed:
            german = (f"Korrigiert in {address(first)}: ={bereiche._german_formula(fixed)}, dann bis {address(uses[-1][0])} "
                      "herunterziehen. Nachgerechnet teilt dann jede Zeile durch die Summe.")
            tip = T(german, f"Corrected in {address(first)}: ={bereiche._german_formula(fixed)}, then fill down to "
                            f"{address(uses[-1][0])}. Recalculated, every row then divides by the total.")
            # Die Korrektur liest einsetzen.py aus dem deutschen Tipp, in jeder Sprache dieselbe.
            fix = einsetzen.fix(sheet, Finding(key, "", "fehler", german))
        else:
            tip = T(f"Setz $ vor die Zeile des Nenners ({column_letters(base[1])}${base[0]}) und kopier die Formel neu herunter.",
                    f"Put $ before the row of the denominator ({column_letters(base[1])}${base[0]}) and copy the formula "
                    "down again.")
        found.append(Finding(
            key,
            T(f"Der Nenner rutscht beim Kopieren mit: {address(first)} teilt durch {address(base)} (die Summe), "
              f"{address(second)} aber schon durch {address(then)}. Nur die erste Zeile ist ein Anteil am Total.",
              f"The denominator shifts when copied: {address(first)} divides by {address(base)} (the total), but "
              f"{address(second)} already by {address(then)}. Only the first row is a share of the total."),
            "fehler", tip, fix=fix,
        ))
    return found[:2]


def _anchored(sheet: Sheet, cells: list[Pos], base: Pos) -> str:
    """Die erste Formel mit festem Nenner (D9 → D$9), auf Englisch. Heruntergezogen muss jede Zeile durch base teilen
    und sich ausrechnen lassen, sonst ""."""
    from .xlsx import shift

    first = cells[0]
    head, sep, tail = (sheet.cells[first].formula or "").rpartition("/")
    if not sep or tail.strip().replace("$", "") != address(base):
        return ""
    english = f"{head}/{'$' if tail.strip().startswith('$') else ''}{column_letters(base[1])}${base[0]}"
    calc, _ = sheet.evaluator()
    try:
        total = calc.run(formel.Node("ref", base))
        for pos in cells:
            node = formel.parse(shift(english, pos[0] - first[0], 0))
            if node.kind != "op" or node.value != "/" or node.args[1] != formel.Node("ref", base):
                return ""
            got, part = calc.run(node), calc.run(node.args[0])
            if not all(isinstance(v, float) and not isinstance(v, bool) for v in (got, part, total)) \
                    or abs(got - part / total) > 1e-9 * max(1.0, abs(got)):
                return ""
    except (formel.Unknown, ValueError, TypeError, ZeroDivisionError):
        return ""
    return english


def _filtered(sheet: Sheet, filtered: list[Area]) -> list[Finding]:
    """SUMME, MITTELWERT und Co. über eine gefilterte Tabelle zählen die ausgeblendeten Zeilen mit."""
    if not filtered:
        return []
    found = []
    for pos, node in sorted(sheet.nodes.items()):
        if node is None:
            continue
        for call in formel.calls(node):
            if call.value not in _AGGREGATES - {"SUBTOTAL", "AGGREGATE"}:
                continue
            ranges = [a.value for a in call.args if a.kind == "range"]
            area = next((f for f in filtered for r in ranges if _overlaps(f, r) and r[2] > r[0]), None)
            if area is None:
                continue
            code = {"SUM": 109, "AVERAGE": 101, "COUNT": 102, "COUNTA": 103, "MAX": 104, "MIN": 105, "STDEV": 107,
                    "STDEV.S": 107, "VAR": 110, "VAR.S": 110}.get(call.value)
            tip = (T(f"Für nur die sichtbaren Zeilen: TEILERGEBNIS({code}; …).", f"For only the visible rows: SUBTOTAL({code}, …).")
                   if code else T("Für nur die sichtbaren Zeilen: AGGREGAT mit Option 5.",
                                  "For only the visible rows: AGGREGATE with option 5."))
            table = f"{address(area[:2])}:{address(area[2:])}"
            found.append(Finding(
                f"filter:{address(pos)}",
                T(f"Die Tabelle {table} ist gefiltert, aber {_german(call.value)} in "
                  f"{sheet.label(pos)} zählt auch die ausgeblendeten Zeilen mit.",
                  f"The table {table} is filtered, but {call.value} in {sheet.label(pos)} also counts the hidden rows."),
                "pruefen", tip + T(" Soll es die ganze Tabelle sein, ist alles richtig.",
                                   " If it is meant to be the whole table, everything is right."),
            ))
            break
        if len(found) >= 2:
            break
    return found


def _overlaps(a: Area, b: Area) -> bool:
    return a[0] <= b[2] and b[0] <= a[2] and a[1] <= b[3] and b[1] <= a[3]
