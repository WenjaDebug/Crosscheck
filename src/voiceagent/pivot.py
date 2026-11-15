"""Eine Pivot-Übersicht mit echten Formeln: SUMMEWENNS, ZÄHLENWENNS, MITTELWERTWENNS, MINWENNS, MAXWENNS.

Gegenprobe rechnet jede Zahl vorher selbst aus, für die Vorschau und zum Nachprüfen, und schreibt
dann Formeln, die auf die Daten zeigen. So rechnet Excel mit, wenn sich die Daten ändern, und
jede Zahl lässt sich in Excel nachvollziehen. Geschrieben wird immer in ein neues Blatt.

Die Formeln stehen englisch mit Komma, so wie Excel sie über AppleScript und COM annimmt
(Range.Formula); Excel zeigt sie dann in seiner Sprache an.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .analyse import value_of
from .auftrag import Auftrag, Wert, function_name
from .sprache import T
from .workbook import Layout
from .xlsx import Cell, column_letters, number_text

# Mehr Zeilen ist keine Übersicht mehr.
MAX_GROUPS = 1000
# Höchstens so viele Spalten in einer Kreuztabelle.
MAX_ACROSS = 30
# „pro Partei und Jahr“ wird eine Kreuztabelle, wenn das Jahr höchstens so viele Werte hat.
AUTO_ACROSS = 12
PREVIEW_ROWS = 12
# Hier beginnt die Übersicht im neuen Blatt (darüber Titel und Herkunft).
TOP = 4

# Wie die Formeln in der deutschen Excel heißen, für die Erklärung.
GERMAN = {"SUMIFS": "SUMMEWENNS", "COUNTIFS": "ZÄHLENWENNS", "AVERAGEIFS": "MITTELWERTWENNS",
          "MINIFS": "MINWENNS", "MAXIFS": "MAXWENNS"}
_IFS = {"summe": "SUMIFS", "anzahl": "COUNTIFS", "mittelwert": "AVERAGEIFS", "min": "MINIFS", "max": "MAXIFS"}
# Excel-Konstanten für echte Pivot-Tabellen (XlConsolidationFunction).
PIVOT_FUNCTIONS = {"summe": -4157, "anzahl": -4112, "mittelwert": -4106, "min": -4139, "max": -4136}

_DATE = re.compile(r"^(\d{2})\.(\d{2})\.(\d{4})(?: (\d{2}):(\d{2}))?$")
_TOTAL = re.compile(r"^(summe|summen|total|totals|gesamt|gesamtsumme|insgesamt|zwischensumme|sum|grand total|"
                    r"gesamtergebnis|endsumme|somme|somma|totale|totaux|sous-total|subtotale|total général|totale generale)\b", re.I)
_TOTAL_FORMULA = re.compile(r"^\s*(SUM|SUBTOTAL|AGGREGATE|SUMME|TEILERGEBNIS)\(", re.I)
_INVALID_NAME = re.compile(r"[\[\]:*?/\\]")
_RANGE = re.compile(r"\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)")


class PivotError(ValueError):
    """Daraus wird keine Übersicht, mit einem Satz, warum."""


@dataclass
class Group:
    key: tuple
    label: str  # wie in der Vorschau
    entry: str  # wie es in die Zelle geschrieben wird (Range.Formula)
    literal: str | None = None  # Kriterium als fester Text statt Zellbezug
    date: bool = False


@dataclass
class Draft:
    """Was in das neue Blatt kommt, Zelle für Zelle, und was dabei herauskommen muss."""

    name: str
    title: str
    # Ab A1, Zeile für Zeile, so wie Range.Formula sie annimmt.
    grid: list[list[str]]
    # Zeile, Spalte → was Gegenprobe selbst ausgerechnet hat (Zahl oder "" für leer).
    expected: dict[tuple[int, int], float | str]
    preview: list[list[str]]
    text: str
    notes: list[str] = field(default_factory=list)
    bold: list[str] = field(default_factory=list)
    formats: list[tuple[str, str]] = field(default_factory=list)
    # Für eine echte Pivot-Tabelle: Quelle und Felder.
    pivot: dict | None = None
    formulas: int = 0


def quote(sheet: str) -> str:
    return "'" + sheet.replace("'", "''") + "'"


def _text_entry(text: str) -> str:
    """Text so, dass Excel ihn nicht als Zahl, Datum oder Formel liest."""
    if not text:
        return ""
    if text[0] in "=+-@'#" or any(ch.isdigit() for ch in text) or text.upper() in ("TRUE", "FALSE", "WAHR", "FALSCH"):
        return "'" + text
    return text


def _string(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def _number_entry(value: float) -> str:
    return str(int(value)) if value == int(value) and abs(value) < 1e15 else repr(value)


def key_of(cell: Cell | None) -> tuple:
    """Wie SUMMEWENNS vergleicht: Text ohne Groß/klein, Zahlen als Zahl, leer als leer."""
    text = cell.value if cell is not None else ""
    if not text.strip():
        return ("e",)
    match = _DATE.match(text)
    if match:
        day, month, year, hour, minute = match.groups()
        return ("d", int(year), int(month), int(day), int(hour or 0), int(minute or 0))
    value = value_of(cell)
    if isinstance(value, bool):
        return ("b", value)
    if isinstance(value, float):
        return ("n", value)
    return ("s", text.casefold())


_ORDER = {"n": 0, "d": 1, "s": 2, "b": 3, "e": 4}


def _sort_key(key: tuple):
    return (_ORDER[key[0]], key[1:])


def _group(key: tuple, text: str) -> Group:
    kind = key[0]
    if kind == "e":
        blank = T("(leer)", "(blank)")
        return Group(key, blank, blank, literal='""')
    if kind == "d":
        _, year, month, day, hour, minute = key
        entry = f"=DATE({year},{month},{day})" + (f"+TIME({hour},{minute},0)" if hour or minute else "")
        return Group(key, text, entry, date=True)
    if kind == "b":
        return Group(key, T("WAHR" if key[1] else "FALSCH", "TRUE" if key[1] else "FALSE"), "TRUE" if key[1] else "FALSE")
    if kind == "n":
        return Group(key, number_text(key[1]), _number_entry(key[1]))
    group = Group(key, text, _text_entry(text))
    if any(ch in text for ch in "*?~") or text[0] in "<>=":
        # Sonst liest SUMMEWENNS * und ? als Platzhalter und < > = als Vergleich.
        escaped = re.sub(r"([*?~])", r"~\1", text)
        group.literal = _string("=" + escaped)
    return group


def show(value) -> str:
    """Eine Zahl für die Vorschau, wie die deutsche Excel sie zeigt."""
    if isinstance(value, str):
        return value
    if value is None:
        return ""
    return number_text(round(value, 2) if value != int(value) else value)


def sheet_name(base: str, existing: list[str]) -> str:
    """Ein neuer Blattname: höchstens 31 Zeichen, ohne [ ] : * ? / \\ und noch nicht vergeben."""
    base = " ".join(_INVALID_NAME.sub(" ", base).split()).strip("'")[:28].strip() or T("Übersicht", "Summary")
    taken = {name.casefold() for name in existing}
    name, number = base, 2
    while name.casefold() in taken:
        name = f"{base[:31 - len(str(number)) - 1]} {number}"
        number += 1
    return name


@dataclass
class _Source:
    sheet: str
    cells: dict[tuple[int, int], Cell]
    first: int
    last: int

    def range(self, col: int) -> str:
        letters = column_letters(col)
        return f"{quote(self.sheet)}!${letters}${self.first}:${letters}${self.last}"

    def rows(self) -> range:
        return range(self.first, self.last + 1)


def _data_rows(cells: dict[tuple[int, int], Cell], plan: Layout, columns: list[int]) -> tuple[int, int, list[str]]:
    """Erste und letzte Datenzeile, ohne eine Gesamtsumme darunter."""
    notes: list[str] = []
    rows = [r for r in range(plan.header_row + 1, plan.last_row + 1)
            if any(cells.get((r, c)) is not None and (cells[(r, c)].value or cells[(r, c)].formula) for c in plan.headers)]
    if not rows:
        raise PivotError(T("Unter der Kopfzeile stehen keine Daten.", "There is no data below the header row."))

    def total(row: int) -> bool:
        for col in {plan.name_column, *columns} - {None}:
            cell = cells.get((row, col))
            if cell is None:
                continue
            if cell.formula is None and _TOTAL.match(cell.value.strip()):
                return True
            if cell.formula and _TOTAL_FORMULA.match(cell.formula) and _sums_above(cell.formula, row, col):
                return True
        return False

    while len(rows) > 1 and total(rows[-1]):
        notes.append(T(f"Zeile {rows[-1]} sieht wie eine Gesamtsumme aus, die lasse ich weg.",
                       f"Row {rows[-1]} looks like a grand total, so I leave it out."))
        rows.pop()
    middle = [r for r in rows[:-1] if total(r)]
    if middle:
        notes.append(T(f"Zeile {middle[0]} sieht wie eine Zwischensumme aus. Sie wird in der Übersicht zu einer eigenen Zeile.",
                       f"Row {middle[0]} looks like a subtotal. It becomes a row of its own in the summary."))
    return rows[0], rows[-1], notes


def _sums_above(formula: str, row: int, col: int) -> bool:
    """Zählt die Formel mehrere Zellen derselben Spalte darüber zusammen (eine Summe unter der Tabelle)?"""
    for match in _RANGE.finditer(formula):
        left, top, right, bottom = match.groups()
        if left == right == column_letters(col) and int(bottom) < row and int(bottom) > int(top):
            return True
    return False


def _numbers(source: _Source, rows: list[int], col: int) -> tuple[list[float], int]:
    """Die Zahlen (wie SUMMEWENNS sie zählt) und wie viele Zellen nicht leer sind."""
    numbers, filled = [], 0
    for row in rows:
        cell = source.cells.get((row, col))
        if cell is None or cell.value == "":
            continue
        filled += 1
        value = value_of(cell)
        if isinstance(value, float):
            numbers.append(value)
    return numbers, filled


def _aggregate(source: _Source, rows: list[int], wert: Wert) -> float | str:
    if wert.spalte is None:
        return float(len(rows))
    numbers, filled = _numbers(source, rows, wert.spalte)
    if wert.funktion == "summe":
        return float(sum(numbers))
    if wert.funktion == "anzahl":
        return float(filled)
    if wert.funktion == "mittelwert":
        return sum(numbers) / len(numbers) if numbers else ""
    if not filled:
        return ""
    if not numbers:
        return 0.0
    return float(min(numbers) if wert.funktion == "min" else max(numbers))


def _formula(source: _Source, wert: Wert, criteria: list[tuple[int, str]]) -> str:
    """Die Formel für eine Zelle der Übersicht. criteria: (Spalte der Daten, Kriterium)."""
    pairs = ",".join(f"{source.range(col)},{crit}" for col, crit in criteria)
    if wert.spalte is None:
        if not criteria:
            raise ValueError("Anzahl ohne Kriterium")
        return f"=COUNTIFS({pairs})"
    values = source.range(wert.spalte)
    if wert.funktion == "anzahl":
        return f"=COUNTIFS({pairs}{',' if pairs else ''}{values},\"<>\")"
    function = _IFS[wert.funktion]
    if wert.funktion == "mittelwert":
        return f"=IFERROR({function}({values},{pairs}),\"\")"
    if wert.funktion in ("min", "max"):
        return f"=IF(COUNTIFS({pairs},{values},\"<>\")=0,\"\",{function}({values},{pairs}))"
    return f"={function}({values},{pairs})"


def _overall(source: _Source, wert: Wert, column_cells: str) -> str:
    """Gesamt über alle Daten. column_cells: die Zellen der Übersicht darüber (für Summe und Anzahl)."""
    values = source.range(wert.spalte) if wert.spalte is not None else ""
    if wert.funktion in ("summe", "anzahl") or wert.spalte is None:
        return f"=SUM({column_cells})"
    if wert.funktion == "mittelwert":
        return f"=IFERROR(AVERAGE({values}),\"\")"
    return f"={'MIN' if wert.funktion == 'min' else 'MAX'}({values})"


def _overall_value(source: _Source, wert: Wert, above: list[float | str]) -> float | str:
    if wert.funktion in ("summe", "anzahl") or wert.spalte is None:
        return float(sum(v for v in above if isinstance(v, float)))
    numbers, _ = _numbers(source, list(source.rows()), wert.spalte)
    if wert.funktion == "mittelwert":
        return sum(numbers) / len(numbers) if numbers else ""
    if not numbers:
        return 0.0
    return float(min(numbers) if wert.funktion == "min" else max(numbers))


def caption(wert: Wert, headers: dict[int, str]) -> str:
    """Wie Excel eine Wertspalte der Pivot-Tabelle nennt: „Summe von Sitze“, „Anzahl“."""
    if wert.spalte is None:
        return function_name("anzahl")
    name = function_name(wert.funktion)
    return T(f"{name} von {headers[wert.spalte]}", f"{name} of {headers[wert.spalte]}")


def _resolve(auftrag: Auftrag, source: _Source, headers: dict[int, str], notes: list[str]) -> list[Wert]:
    """Rechenart je Wert festlegen: Zahlen ergeben eine Summe, Text eine Anzahl."""
    werte = []
    rows = list(source.rows())
    for wert in auftrag.werte:
        function = wert.funktion
        if wert.spalte is not None:
            numbers, filled = _numbers(source, rows, wert.spalte)
            numeric = bool(numbers) and len(numbers) * 2 >= filled
            if function is None:
                function = "summe" if numeric else "anzahl"
            elif function != "anzahl" and not numbers:
                notes.append(T(f"In {headers[wert.spalte]} stehen keine Zahlen, darum zähle ich die Einträge.",
                               f"{headers[wert.spalte]} has no numbers, so I count the entries."))
                function = "anzahl"
        resolved = Wert(function or "anzahl", wert.spalte)
        if resolved not in werte:
            werte.append(resolved)
    return werte


def _groups(source: _Source, col: int, headers: dict[int, str], data_cols: list[int]) -> tuple[list[Group], dict]:
    """Die verschiedenen Werte einer Spalte und in welchen Zeilen sie stehen."""
    seen: dict[tuple, str] = {}
    members: dict[tuple, list[int]] = {}
    with_data_but_empty = False
    for row in source.rows():
        cell = source.cells.get((row, col))
        key = key_of(cell)
        members.setdefault(key, []).append(row)
        if key == ("e",):
            if any(source.cells.get((row, c)) is not None and source.cells[(row, c)].value for c in data_cols if c != col):
                with_data_but_empty = True
            continue
        seen.setdefault(key, cell.value)
    groups = [_group(k, seen[k]) for k in sorted(seen, key=_sort_key)]
    if with_data_but_empty:
        groups.append(_group(("e",), ""))
    if len(groups) > MAX_GROUPS:
        raise PivotError(T(f"In {headers[col]} stehen {len(groups)} verschiedene Werte. Das wird keine Übersicht. "
                           "Nimm eine Spalte mit weniger verschiedenen Werten.",
                           f"{headers[col]} has {len(groups)} different values. That is too many for a summary. "
                           "Choose a column with fewer different values."))
    return groups, members


def _criterion(group: Group, address: str) -> str:
    return group.literal if group.literal is not None else address


def build(cells: dict[tuple[int, int], Cell], plan: Layout, sheet: str, existing: list[str], auftrag: Auftrag) -> Draft:
    """Die Übersicht für das neue Blatt: Formeln, erwartete Werte und Vorschau."""
    if plan.header_row is None or not plan.headers:
        raise PivotError(T("Ich finde auf diesem Blatt keine Kopfzeile mit Spaltennamen.",
                           "I cannot find a header row with column names on this sheet."))
    headers = plan.headers
    used = [*auftrag.zeilen, *([auftrag.spalte] if auftrag.spalte else []), *(w.spalte for w in auftrag.werte if w.spalte)]
    if any(col not in headers for col in used):
        raise PivotError(T("Diese Spalte gibt es auf dem Blatt nicht mehr.", "This column is no longer on the sheet."))
    first, last, notes = _data_rows(cells, plan, [w.spalte for w in auftrag.werte if w.spalte])
    source = _Source(sheet, cells, first, last)
    werte = _resolve(auftrag, source, headers, notes)
    data_cols = sorted(headers)
    rows_fields = list(auftrag.zeilen)
    across = auftrag.spalte
    if across is None and len(rows_fields) == 2 and len(werte) == 1:
        groups, _ = _groups(source, rows_fields[1], headers, data_cols)
        if len(groups) <= AUTO_ACROSS:
            across = rows_fields.pop()
    if across is not None and len(werte) > 1:
        raise PivotError(T("Eine Kreuztabelle geht mit einem Wert. Nimm einen, zum Beispiel nur die Summe.",
                           "A cross table works with one value. Pick one, for example just the sum."))
    if across is not None and len(rows_fields) > 1:
        raise PivotError(T("Eine Kreuztabelle hat eine Spalte in den Zeilen und eine in den Spalten.",
                           "A cross table has one column in the rows and one in the columns."))
    if across is not None:
        draft = _crosstab(source, headers, rows_fields[0], across, werte[0], auftrag, existing, data_cols)
    else:
        draft = _table(source, headers, rows_fields, werte, auftrag, existing, data_cols)
    draft.notes = notes + draft.notes
    draft.pivot = _pivot_spec(cells, plan, first, last, rows_fields, across, werte, sheet, draft.notes)
    return draft


def _ranked(values: list[float | str], order: str) -> list[int]:
    """Reihenfolge nach Wert, leere Werte immer zuletzt."""
    filled = sorted((i for i, v in enumerate(values) if v != ""), key=lambda i: values[i], reverse=order == "absteigend")
    return filled + [i for i, v in enumerate(values) if v == ""]


def _title(werte: list[Wert], headers: dict[int, str], rows: list[int], across: int | None) -> str:
    what = T(" und ", " and ").join(caption(w, headers) for w in werte)
    by = T(" und ", " and ").join(headers[c] for c in rows)
    return T(f"{what} pro {by}" + (f", Spalten nach {headers[across]}" if across else ""),
             f"{what} per {by}" + (f", columns by {headers[across]}" if across else ""))


def _head(title: str, source: _Source, width: int) -> list[list[str]]:
    origin = T(f"Aus Blatt {source.sheet}, Zeilen {source.first} bis {source.last}. Die Zahlen sind Formeln und rechnen mit, "
               "wenn sich die Daten ändern.",
               f"From sheet {source.sheet}, rows {source.first} to {source.last}. The numbers are formulas and update "
               "when the data changes.")
    rows = [[title], [_text_entry(origin)], []]
    return [r + [""] * (width - len(r)) for r in rows]


def _formula_names(werte: list[Wert]) -> str:
    names = []
    for wert in werte:
        english = "COUNTIFS" if wert.spalte is None else _IFS[wert.funktion]
        name = T(GERMAN[english], english)
        if name not in names:
            names.append(name)
    return T(" und ", " and ").join(names)


def _table(source: _Source, headers: dict[int, str], rows_fields: list[int], werte: list[Wert], auftrag: Auftrag,
           existing: list[str], data_cols: list[int]) -> Draft:
    if len(rows_fields) == 1:
        groups, members = _groups(source, rows_fields[0], headers, data_cols)
        lines = [([g], members.get(g.key, [])) for g in groups]
    else:
        outer, outer_members = _groups(source, rows_fields[0], headers, data_cols)
        inner, inner_members = _groups(source, rows_fields[1], headers, data_cols)
        lines = []
        for a in outer:
            for b in inner:
                both = sorted(set(outer_members.get(a.key, [])) & set(inner_members.get(b.key, [])))
                if both:
                    lines.append(([a, b], both))
        if len(lines) > MAX_GROUPS:
            raise PivotError(f"Das ergibt {len(lines)} Zeilen. Das wird keine Übersicht.")
    labels = len(rows_fields)
    width = labels + len(werte)
    values = [[_aggregate(source, rows, w) for w in werte] for _, rows in lines]
    if auftrag.sortierung:
        order = _ranked([v[0] for v in values], auftrag.sortierung)
        lines, values = [lines[i] for i in order], [values[i] for i in order]
    title = _title(werte, headers, rows_fields, None)
    grid = _head(title, source, width)
    grid.append([headers[c] for c in rows_fields] + [caption(w, headers) for w in werte])
    expected: dict[tuple[int, int], float | str] = {}
    formats: list[tuple[str, str]] = []
    start = TOP + 1
    for i, (line_groups, _) in enumerate(lines):
        row = start + i
        cells = [g.entry for g in line_groups]
        for j, wert in enumerate(werte):
            criteria = [(rows_fields[k], _criterion(g, f"${column_letters(k + 1)}{row}")) for k, g in enumerate(line_groups)]
            cells.append(_formula(source, wert, criteria))
            expected[(row, labels + j + 1)] = values[i][j]
        grid.append(cells)
        for k, g in enumerate(line_groups):
            if g.date:
                formats.append((f"{column_letters(k + 1)}{row}", "dd.mm.yyyy"))
    end = start + len(lines) - 1
    total_row = end + 1
    total = [T("Gesamt", "Total")] + [""] * (labels - 1)
    for j, wert in enumerate(werte):
        letters = column_letters(labels + j + 1)
        total.append(_overall(source, wert, f"{letters}{start}:{letters}{max(end, start)}"))
        expected[(total_row, labels + j + 1)] = _overall_value(source, wert, [v[j] for v in values])
    grid.append(total)
    for j, wert in enumerate(werte):
        if wert.funktion == "mittelwert":
            letters = column_letters(labels + j + 1)
            formats.append((f"{letters}{start}:{letters}{total_row}", "0.00"))
    last_col = column_letters(width)
    preview = [grid[TOP - 1]]
    for i, (line_groups, _) in enumerate(lines[:PREVIEW_ROWS]):
        preview.append([g.label for g in line_groups] + [show(v) for v in values[i]])
    if len(lines) > PREVIEW_ROWS:
        preview.append(["…"] + [""] * (width - 1))
    preview.append([T("Gesamt", "Total")] + [""] * (labels - 1) + [show(expected[(total_row, labels + j + 1)]) for j in range(len(werte))])
    name = sheet_name(f"Pivot {caption(werte[0], headers).split(T(' von ', ' of '))[-1]}", existing)
    text = T(f"Ich lege das neue Blatt „{name}“ an: {title}, {len(lines)} {'Zeile' if len(lines) == 1 else 'Zeilen'} "
             f"und eine Gesamtzeile. Die Zahlen sind Formeln ({_formula_names(werte)}) auf Blatt „{source.sheet}“, "
             f"Zeilen {source.first} bis {source.last}. Deine bestehenden Zellen ändere ich nicht.",
             f"I will add the new sheet “{name}”: {title}, {len(lines)} {'row' if len(lines) == 1 else 'rows'} "
             f"and a total row. The numbers are formulas ({_formula_names(werte)}) on sheet “{source.sheet}”, "
             f"rows {source.first} to {source.last}. I will not change your existing cells.")
    return Draft(
        name=name, title=title, grid=[r + [""] * (width - len(r)) for r in grid], expected=expected, preview=preview,
        text=text, bold=["A1", f"A{TOP}:{last_col}{TOP}", f"A{total_row}:{last_col}{total_row}"], formats=formats,
        formulas=len(lines) * len(werte) + len(werte),
    )


def _crosstab(source: _Source, headers: dict[int, str], row_field: int, across: int, wert: Wert, auftrag: Auftrag,
              existing: list[str], data_cols: list[int]) -> Draft:
    groups, members = _groups(source, row_field, headers, data_cols)
    columns, column_members = _groups(source, across, headers, data_cols)
    if len(columns) > MAX_ACROSS:
        raise PivotError(T(f"In {headers[across]} stehen {len(columns)} verschiedene Werte, das sind zu viele Spalten. "
                           f"Nimm sie lieber in die Zeilen: … pro {headers[row_field]} und {headers[across]}.",
                           f"{headers[across]} has {len(columns)} different values, that is too many columns. "
                           f"Put it in the rows instead: … per {headers[row_field]} and {headers[across]}."))
    width = len(columns) + 2
    rows_of = {g.key: set(members.get(g.key, [])) for g in groups}
    cols_of = {c.key: set(column_members.get(c.key, [])) for c in columns}
    values = [[_aggregate(source, sorted(rows_of[g.key] & cols_of[c.key]), wert) for c in columns] for g in groups]
    row_totals = [_aggregate(source, sorted(rows_of[g.key]), wert) for g in groups]
    column_totals = [_aggregate(source, sorted(cols_of[c.key]), wert) for c in columns]
    if auftrag.sortierung:
        order = _ranked(row_totals, auftrag.sortierung)
        groups, values, row_totals = [groups[i] for i in order], [values[i] for i in order], [row_totals[i] for i in order]
    title = _title([wert], headers, [row_field], across)
    grid = _head(title, source, width)
    grid[TOP - 2] = [caption(wert, headers), _text_entry(headers[across])] + [""] * (width - 2)
    grid.append([headers[row_field]] + [c.entry for c in columns] + [T("Gesamt", "Total")])
    expected: dict[tuple[int, int], float | str] = {}
    formats: list[tuple[str, str]] = []
    start = TOP + 1
    for j, c in enumerate(columns):
        if c.date:
            formats.append((f"{column_letters(j + 2)}{TOP}", "dd.mm.yyyy"))
    for i, g in enumerate(groups):
        row = start + i
        line = [g.entry]
        row_crit = (row_field, _criterion(g, f"$A{row}"))
        for j, c in enumerate(columns):
            col_crit = (across, _criterion(c, f"{column_letters(j + 2)}${TOP}"))
            line.append(_formula(source, wert, [row_crit, col_crit]))
            expected[(row, j + 2)] = values[i][j]
        line.append(_formula(source, wert, [row_crit]))
        expected[(row, width)] = row_totals[i]
        grid.append(line)
        if g.date:
            formats.append((f"A{row}", "dd.mm.yyyy"))
    end = start + len(groups) - 1
    total_row = end + 1
    total = [T("Gesamt", "Total")]
    for j, c in enumerate(columns):
        total.append(_formula(source, wert, [(across, _criterion(c, f"{column_letters(j + 2)}${TOP}"))]))
        expected[(total_row, j + 2)] = column_totals[j]
    letters = column_letters(width)
    total.append(_overall(source, wert, f"{letters}{start}:{letters}{max(end, start)}"))
    expected[(total_row, width)] = _overall_value(source, wert, row_totals)
    grid.append(total)
    if wert.funktion == "mittelwert":
        formats.append((f"B{start}:{letters}{total_row}", "0.00"))
    preview = [[headers[row_field]] + [c.label for c in columns] + [T("Gesamt", "Total")]]
    for i, g in enumerate(groups[:PREVIEW_ROWS]):
        preview.append([g.label] + [show(v) for v in values[i]] + [show(row_totals[i])])
    if len(groups) > PREVIEW_ROWS:
        preview.append(["…"] + [""] * (width - 1))
    preview.append([T("Gesamt", "Total")] + [show(v) for v in column_totals] + [show(expected[(total_row, width)])])
    name = sheet_name(f"Pivot {headers[row_field]} {headers[across]}", existing)
    text = T(f"Ich lege das neue Blatt „{name}“ an: {title}, {len(groups)} Zeilen und {len(columns)} Spalten, dazu die "
             f"Gesamtwerte. Die Zahlen sind Formeln ({_formula_names([wert])}) auf Blatt „{source.sheet}“, Zeilen "
             f"{source.first} bis {source.last}. Deine bestehenden Zellen ändere ich nicht.",
             f"I will add the new sheet “{name}”: {title}, {len(groups)} rows and {len(columns)} columns, plus the "
             f"totals. The numbers are formulas ({_formula_names([wert])}) on sheet “{source.sheet}”, rows "
             f"{source.first} to {source.last}. I will not change your existing cells.")
    return Draft(
        name=name, title=title, grid=[r + [""] * (width - len(r)) for r in grid], expected=expected, preview=preview,
        text=text, bold=["A1", f"A{TOP - 1}:{letters}{TOP}", f"A{total_row}:{letters}{total_row}", f"{letters}{TOP}:{letters}{total_row}"],
        formats=formats, formulas=(len(groups) + 1) * (len(columns) + 1),
    )


def _pivot_spec(cells: dict, plan: Layout, first: int, last: int, rows_fields: list[int], across: int | None,
                werte: list[Wert], sheet: str, notes: list[str]) -> dict | None:
    """Was eine echte Pivot-Tabelle braucht: den Bereich mit Kopfzeile und die Felder mit ihrem Namen in Excel."""
    left, right = min(plan.headers), max(plan.headers)
    names = {}
    for col in range(left, right + 1):
        cell = cells.get((plan.header_row, col))
        if cell is None or not cell.value.strip() or cell.formula is not None:
            return None
        names[col] = cell.value
    if len({n.casefold() for n in names.values()}) < len(names):
        return None
    data = []
    for wert in werte:
        field_col = wert.spalte if wert.spalte is not None else rows_fields[0]
        data.append({"field": names[field_col], "function": PIVOT_FUNCTIONS[wert.funktion if wert.spalte else "anzahl"],
                     "caption": caption(wert, plan.headers)})
    return {
        "sheet": sheet,
        "range": f"{column_letters(left)}{plan.header_row}:{column_letters(right)}{last}",
        "rows": [names[c] for c in rows_fields],
        "columns": [names[across]] if across is not None else [],
        "data": data,
    }
