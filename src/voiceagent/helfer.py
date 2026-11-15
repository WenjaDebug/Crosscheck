"""Excel-Helfer: Du schreibst, was du willst, Gegenprobe macht es in Excel.

Ablauf: Vorschau zuerst (was er tun wird und welche Zahlen herauskommen), geschrieben wird erst
nach deinem Klick, immer in ein neues Blatt. Danach liest er das neue Blatt zurück und prüft jede
Zahl gegen seine eigene Rechnung. Den Satz versteht er ohne KI.
"""

from __future__ import annotations

import hashlib
import logging
import secrets
import threading
from dataclasses import dataclass

from .auftrag import Auftrag, understand
from .einlesen import draft as import_draft
from .einlesen import tables_from
from .pivot import Draft, PivotError, build, sheet_name
from .pruefer import STATES
from .schreiben import MAX_WRONG, ExcelWriter, verify
from .workbook import Layout, _as_number, layout
from .xlsx import Cell, Sheet

logger = logging.getLogger(__name__)

# So viele Vorschauen merkt er sich.
KEEP = 5


@dataclass
class _Pending:
    text: str
    auftrag: Auftrag | None
    draft: Draft
    book: str
    sheet: str
    fingerprint: str
    # pivot: Übersicht aus dem aktiven Blatt, import: Tabelle aus einer Datei
    kind: str = "pivot"


def _fingerprint(cells: dict[tuple[int, int], Cell]) -> str:
    digest = hashlib.sha256()
    for (row, col), cell in sorted(cells.items()):
        digest.update(f"{row},{col},{cell.value},{cell.formula}\n".encode("utf-8", "replace"))
    return digest.hexdigest()


def numeric_columns(cells: dict[tuple[int, int], Cell], plan: Layout) -> set[int]:
    """Spalten, in denen unter der Kopfzeile meist Zahlen stehen."""
    result = set()
    for col in plan.headers:
        values = [cells[(r, col)].value for r in range(plan.header_row + 1, plan.last_row + 1)
                  if (r, col) in cells and cells[(r, col)].value]
        numbers = sum(1 for v in values if _as_number(v) is not None)
        if values and numbers * 2 >= len(values):
            result.add(col)
    return result


def examples(cells: dict[tuple[int, int], Cell], plan: Layout, numeric: set[int]) -> list[str]:
    """Beispielsätze mit den echten Spaltennamen: Zahlen zusammenfassen, nach Text oder Jahren aufteilen."""

    def years(col: int) -> bool:
        values = [_as_number(cells[(r, col)].value) for r in range(plan.header_row + 1, plan.last_row + 1) if (r, col) in cells]
        return bool(values) and all(v is not None and v == int(v) and 1900 <= v <= 2100 for v in values)

    numbers = [c for c in sorted(numeric) if not years(c)]
    groups = [c for c in sorted(plan.headers) if c not in numbers]
    if plan.name_column in groups:
        groups.remove(plan.name_column)
        groups.insert(0, plan.name_column)
    if not groups:
        return []
    name = plan.headers
    found = [f"Anzahl pro {name[groups[0]]}"]
    if numbers:
        found.insert(0, f"Summe von {name[numbers[0]]} pro {name[groups[0]]}")
        found.append(f"Durchschnitt {name[numbers[-1]]} pro {name[groups[0]]}, größte zuerst")
        if len(groups) > 1:
            found.append(f"Summe von {name[numbers[0]]} pro {name[groups[0]]} und {name[groups[1]]}")
    return found


class Helper:
    """Einer für die ganze App. Vorschau und Schreiben nie zugleich."""

    def __init__(self, link) -> None:
        self.link = link
        self._lock = threading.Lock()
        self._pending: dict[str, _Pending] = {}
        self._headers: list[str] = []
        # Blätter, die der Helfer selbst angelegt hat: (Mappe, Blatt) → Blatt mit den Daten.
        self._made: dict[tuple[str, str], str] = {}
        # Hochgeladene Dateien, nur im Arbeitsspeicher: Token → (Dateiname, Tabellen).
        self._files: dict[str, tuple[str, list]] = {}

    def _sheet(self) -> tuple[dict | None, tuple]:
        """Das aktive Blatt ganz gelesen: (Fehler, (Schnappschuss, Zellen))."""
        link = self.link
        if link is None or not link.supported:
            return {"ok": False, "state": "off", "text": "Der Excel-Helfer geht nur am Mac und unter Windows."}, ()
        if not link.running():
            return {"ok": False, "state": "closed", "text": "Öffne zuerst die Mappe mit deinen Daten in Excel."}, ()
        snapshot = link.read(full=True)
        if snapshot is None:
            state = link.status.state
            return {"ok": False, "state": state, "text": STATES.get(state, STATES["error"])}, ()
        cells = snapshot.cells
        if cells is None:
            book = link.book(snapshot)
            sheet = next((s for s in book.sheets if s.name == snapshot.sheet), None) if book is not None else None
            if sheet is None:
                return {"ok": False, "state": "error", "text": "Die Zellen gibt Excel nicht her, und die gespeicherte Datei finde ich nicht."}, ()
            cells = sheet.cells
        return None, (snapshot, cells)

    def preview(self, text: str) -> dict:
        if not self._lock.acquire(timeout=30):
            return {"ok": False, "state": "busy", "text": "Ich bin noch mit dem letzten Auftrag beschäftigt."}
        try:
            return self._preview(text)
        finally:
            self._lock.release()

    def _preview(self, text: str) -> dict:
        problem, found = self._sheet()
        if problem is not None:
            return problem
        snapshot, cells = found
        source = self._made.get((snapshot.book, snapshot.sheet))
        if source is not None:
            return {"ok": False, "state": "own", "text": f"Vorne ist das Blatt „{snapshot.sheet}“, das ich angelegt habe. "
                                                         f"Klick in Excel auf das Blatt mit deinen Daten, zum Beispiel „{source}“."}
        plan = layout(Sheet(snapshot.sheet, cells))
        if plan.header_row is None or not plan.headers:
            return {"ok": False, "state": "empty", "text": f"Auf Blatt „{snapshot.sheet}“ finde ich keine Tabelle mit "
                                                          "Spaltennamen. Klick in Excel auf das Blatt mit deinen Daten."}
        self._headers = list(plan.headers.values())
        numeric = numeric_columns(cells, plan)
        hints = examples(cells, plan, numeric)
        understood = understand(text, plan.headers, numeric)
        if understood.auftrag is None:
            return {"ok": False, "state": "unclear", "text": understood.text, "columns": self._headers, "examples": hints}
        try:
            draft = build(cells, plan, snapshot.sheet, snapshot.sheets, understood.auftrag)
        except PivotError as err:
            return {"ok": False, "state": "unclear", "text": str(err), "columns": self._headers, "examples": hints}
        token = secrets.token_hex(8)
        self._pending[token] = _Pending(text, understood.auftrag, draft, snapshot.book, snapshot.sheet, _fingerprint(cells))
        while len(self._pending) > KEEP:
            self._pending.pop(next(iter(self._pending)))
        return {
            "ok": True,
            "token": token,
            "book": snapshot.book,
            "source": snapshot.sheet,
            "sheet": draft.name,
            "title": draft.title,
            "text": draft.text,
            "notes": draft.notes,
            "preview": draft.preview,
            "formulas": draft.formulas,
            "pivot_possible": draft.pivot is not None,
            "columns": self._headers,
            "examples": hints,
        }

    def run(self, token: str, real_pivot: bool = False) -> dict:
        if not self._lock.acquire(timeout=30):
            return {"ok": False, "state": "busy", "text": "Ich bin noch mit dem letzten Auftrag beschäftigt."}
        try:
            return self._run(token, real_pivot)
        finally:
            self._lock.release()

    def _run(self, token: str, real_pivot: bool) -> dict:
        pending = self._pending.get(token)
        if pending is None:
            return {"ok": False, "state": "expired", "text": "Diese Vorschau kenne ich nicht mehr. Lass sie noch einmal zeigen."}
        if pending.kind == "import":
            return self._run_import(token, pending)
        problem, found = self._sheet()
        if problem is not None:
            return problem
        snapshot, cells = found
        if snapshot.book != pending.book or snapshot.sheet != pending.sheet:
            return {"ok": False, "state": "moved", "text": f"Klick in Excel bitte wieder auf das Blatt „{pending.sheet}“ "
                                                          f"in der Mappe „{pending.book}“. Von dort habe ich die Daten."}
        if _fingerprint(cells) != pending.fingerprint:
            # Die Daten haben sich geändert: neu rechnen, neu zeigen, nichts schreiben.
            self._pending.pop(token, None)
            fresh = self._preview(pending.text)
            fresh["changed"] = True
            if fresh.get("ok"):
                fresh["text"] = "Deine Daten haben sich seit der Vorschau geändert. Hier ist die neue Vorschau. " + fresh["text"]
            return fresh
        return self._write(token, pending, snapshot.sheets, real_pivot)

    def _run_import(self, token: str, pending: _Pending) -> dict:
        link = self.link
        if link is None or not link.supported or not link.running():
            return {"ok": False, "state": "closed", "text": "Öffne in Excel die Mappe, in die die Tabelle soll."}
        snapshot = link.read()
        if snapshot is None:
            state = link.status.state
            return {"ok": False, "state": state, "text": STATES.get(state, STATES["error"])}
        if snapshot.book != pending.book:
            return {"ok": False, "state": "moved", "text": f"Vorne ist jetzt die Mappe „{snapshot.book}“. Ich schreibe nur in "
                                                          f"„{pending.book}“, die Mappe aus der Vorschau. Hol sie wieder nach vorne."}
        return self._write(token, pending, snapshot.sheets, False)

    def _write(self, token: str, pending: _Pending, sheets: list[str], real_pivot: bool) -> dict:
        draft = pending.draft
        if draft.name.casefold() in {s.casefold() for s in sheets}:
            draft.name = sheet_name(draft.name, sheets)
        written = ExcelWriter(self.link).write(draft, pending.book, real_pivot and draft.pivot is not None)
        if not written.ok:
            return {"ok": False, "state": written.state, "text": written.text}
        self._pending.pop(token, None)
        if pending.kind == "pivot":
            self._made[(pending.book, written.sheet)] = pending.sheet
        what = "Werte" if pending.kind == "import" else "Zahlen"
        undo = " Gefällt es dir nicht: Rechtsklick unten auf den Blattnamen, Löschen. Rückgängig geht dafür nicht."
        if written.pivot:
            return {"ok": True, "state": "ok", "sheet": written.sheet, "pivot": True,
                    "text": f"Fertig: Excel hat auf Blatt „{written.sheet}“ eine echte Pivot-Tabelle angelegt. Ändern sich "
                            "die Daten, klick mit rechts hinein und wähle Aktualisieren." + undo}
        intro = ""
        if real_pivot and written.pivot_error:
            intro = (f"Eine echte Pivot-Tabelle hat Excel nicht angelegt ({written.pivot_error[:120]}). Darum steht die "
                     "Übersicht mit Formeln da. ")
        elif real_pivot and draft.pivot is None:
            intro = "Für eine echte Pivot-Tabelle braucht jede Spalte der Daten einen eigenen Namen. Darum steht die Übersicht mit Formeln da. "
        right, wrong = verify(draft, self.link.read(full=True), written.sheet)
        total = len(draft.expected)
        if wrong:
            examples = "; ".join(wrong[:MAX_WRONG])
            text = (f"{intro}Das Blatt „{written.sheet}“ ist da, aber {len(wrong)} von {total} {what} zeigt Excel anders, "
                    f"als ich erwartet habe: {examples}. Schau dir diese Zellen an, bevor du damit weiterarbeitest.")
            return {"ok": True, "state": "different", "sheet": written.sheet, "text": text + undo, "wrong": wrong}
        if right and pending.kind == "import":
            text = (f"Fertig: Das Blatt „{written.sheet}“ ist da. Ich habe alle {right} Werte in Excel nachgeprüft, "
                    "sie stehen so da wie in der Datei.")
        elif right:
            text = (f"{intro}Fertig: Das Blatt „{written.sheet}“ ist da. Ich habe alle {right} Zahlen in Excel nachgeprüft, "
                    "sie stimmen mit meiner eigenen Rechnung überein.")
        else:
            text = f"{intro}Fertig: Das Blatt „{written.sheet}“ ist da. Nachprüfen konnte ich es nicht, Excel gab das Blatt nicht her."
        return {"ok": True, "state": "ok", "sheet": written.sheet, "text": text + undo}

    def load(self, name: str, data: bytes) -> dict:
        """Eine Datei mit Tabellen: alle Tabellen nennen und die erste als Vorschau zeigen."""
        try:
            tables = tables_from(name, data)
        except ValueError as err:
            return {"ok": False, "state": "file", "text": str(err)}
        token = secrets.token_hex(8)
        self._files[token] = (name, tables)
        while len(self._files) > KEEP:
            self._files.pop(next(iter(self._files)))
        result = self.table(token, 0)
        result["file"] = token
        result["tables"] = [{"name": t.name, "rows": t.size[0], "columns": t.size[1], "kind": t.kind} for t in tables]
        return result

    def table(self, file: str, index: int) -> dict:
        """Vorschau für eine Tabelle aus der Datei: so kommt sie in ein neues Blatt."""
        found = self._files.get(file)
        if found is None:
            return {"ok": False, "state": "expired", "text": "Diese Datei kenne ich nicht mehr. Lade sie noch einmal."}
        name, tables = found
        if not 0 <= index < len(tables):
            return {"ok": False, "state": "expired", "text": "Diese Tabelle gibt es in der Datei nicht."}
        link = self.link
        if link is None or not link.supported:
            return {"ok": False, "state": "off", "text": "Der Excel-Helfer geht nur am Mac und unter Windows."}
        if not link.running():
            return {"ok": False, "state": "closed", "text": "Öffne in Excel die Mappe, in die die Tabelle soll (oder eine neue, leere)."}
        snapshot = link.read()
        if snapshot is None:
            state = link.status.state
            return {"ok": False, "state": state, "text": STATES.get(state, STATES["error"])}
        table = tables[index]
        try:
            draft = import_draft(table, snapshot.sheets, name)
        except ValueError as err:
            return {"ok": False, "state": "file", "text": str(err)}
        token = secrets.token_hex(8)
        self._pending[token] = _Pending(name, None, draft, snapshot.book, snapshot.sheet, "", kind="import")
        while len(self._pending) > KEEP:
            self._pending.pop(next(iter(self._pending)))
        return {"ok": True, "token": token, "book": snapshot.book, "sheet": draft.name, "text": draft.text,
                "notes": draft.notes, "preview": draft.preview, "pivot_possible": False, "index": index}
