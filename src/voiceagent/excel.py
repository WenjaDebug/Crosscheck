"""Excel live mitlesen: Gegenprobe fragt das laufende Excel nach der offenen Mappe.

Am Mac über AppleScript (JavaScript für Automation mit osascript), unter Windows über
PowerShell. Beides ist schon im System, es kommt kein Paket dazu. Gelesen werden das aktive
Blatt (Werte und Formeln), die Namen der Blätter, wo der Cursor steht und wo die Datei liegt.
Was nur in der gespeicherten Datei steht (Filter, Diagramme, Pivot-Tabellen …), liest
xlsx.py aus der Datei selbst, sobald Excel sie speichert. Beim Mitlesen ändert Gegenprobe nichts in Excel;
geschrieben wird nur auf Wunsch und nur in ein neues Blatt (schreiben.py).

Beim ersten Mal fragt macOS: „Terminal möchte Microsoft Excel steuern“. Erst nach „Erlauben“
geht es. Wer ablehnt, findet den Schalter unter Systemeinstellungen → Datenschutz & Sicherheit
→ Automation → Terminal → Microsoft Excel.
"""

from __future__ import annotations

import base64
import json
import logging
import os
import re
import subprocess
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote, urlparse

from .xlsx import MAX_FILE, Book, Cell, date_text, number_text, read_book

logger = logging.getLogger(__name__)

# So oft wird nachgesehen.
INTERVAL = 2.0
INTERVAL_WINDOWS = 3.0
# Beim ersten Mal wartet macOS auf „Erlauben“, so lange darf es dauern.
FIRST_TIMEOUT = 120.0
# Während jemand in einer Zelle tippt, antwortet Excel erst danach. So lange wird gewartet.
TIMEOUT = 20.0
# Größere oder langsamere Blätter (Sekunden für einmal Lesen): nur noch, was im Fenster zu sehen ist.
MAX_LIVE_CELLS = 30_000
SLOW = 4.0
# Das ganze Blatt auf Wunsch (Pivot-Übersicht): so lange darf es dauern.
FULL_TIMEOUT = 60.0
# So oft hintereinander ohne Antwort, dann gilt das Blatt als zu groß.
MAX_BUSY = 3
# Eine Datei, die gerade geschrieben wird, erst so lange danach lesen.
SETTLE = 1.0
# Beim Suchen einer OneDrive-Datei nach Namen höchstens so viele Einträge ansehen.
MAX_SEARCH = 20_000

_MAC_SCRIPT = r"""
function text(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) {
    const p = (n) => String(n).padStart(2, "0");
    let s = p(v.getDate()) + "." + p(v.getMonth() + 1) + "." + v.getFullYear();
    if (v.getHours() || v.getMinutes()) s += " " + p(v.getHours()) + ":" + p(v.getMinutes());
    return {d: s};
  }
  if (typeof v === "boolean") return v ? "WAHR" : "FALSCH";
  return v;
}
function grid(x) {
  if (x === null || x === undefined) return [[""]];
  if (!Array.isArray(x)) return [[x]];
  if (x.length && !Array.isArray(x[0])) return [x];
  return x;
}
function run(argv) {
  const xl = Application("Microsoft Excel");
  if (!xl.running()) return JSON.stringify({running: false});
  const out = {running: true};
  try { out.front = xl.frontmost(); } catch (e) {}
  try {
    const book = xl.activeWorkbook;
    out.book = book.name();
    try { out.path = book.fullName(); } catch (e) {}
    try { out.sheets = book.worksheets.name(); } catch (e) {}
  } catch (e) { out.nobook = String(e); return JSON.stringify(out); }
  try {
    const sheet = xl.activeSheet;
    out.sheet = sheet.name();
    const range = argv[0] === "visible" ? xl.activeWindow.visibleRange : sheet.usedRange;
    out.row = range.firstRowIndex();
    out.col = range.firstColumnIndex();
    out.formulas = grid(range.formula());
    out.values = grid(range.value()).map((r) => r.map(text));
  } catch (e) { out.error = String(e); }
  try { const cell = xl.activeCell; out.at = [cell.firstRowIndex(), cell.firstColumnIndex()]; } catch (e) {}
  return JSON.stringify(out);
}
"""

# Jede Zeile: Art, Tab, Inhalt. Zellen: „cell Zeile Spalte Formel Wert“, Werte mit Typ davor (n:, d:, b:, s:).
_WINDOWS_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$inv = [Globalization.CultureInfo]::InvariantCulture
function Clean($v) { return ([string]$v) -replace "[`t`r`n]", " " }
function Out($k, $v) { [Console]::Out.WriteLine($k + "`t" + (Clean $v)) }
function Typed($v) {
  if ($v -eq $null) { return '' }
  if ($v -is [DateTime]) { return 'd:' + $v.ToString('dd.MM.yyyy HH:mm', $inv) }
  if ($v -is [bool]) { if ($v) { return 'b:1' } else { return 'b:0' } }
  if ($v -is [double] -or $v -is [int] -or $v -is [decimal]) { return 'n:' + ([double]$v).ToString('R', $inv) }
  return 's:' + (Clean $v)
}
try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { Out 'running' '0'; exit 0 }
Out 'running' '1'
try {
  $book = $xl.ActiveWorkbook
  if ($book -eq $null) { exit 0 }
  Out 'book' $book.Name
  Out 'path' $book.FullName
  foreach ($s in $book.Worksheets) { Out 'sheets' $s.Name }
  $sheet = $xl.ActiveSheet
  Out 'sheet' $sheet.Name
  if ($mode -eq 'visible') { $range = $xl.ActiveWindow.VisibleRange } else { $range = $sheet.UsedRange }
  Out 'row' $range.Row
  Out 'col' $range.Column
  $rows = $range.Rows.Count
  $cols = $range.Columns.Count
  Out 'size' ("" + $rows + " " + $cols)
  $f = $range.Formula
  $v = $range.Value2
  if ($rows -eq 1 -and $cols -eq 1) {
    [Console]::Out.WriteLine("cell`t1`t1`t" + (Clean $f) + "`t" + (Typed $range.Value()))
  } else {
    $vals = $range.Value()
    for ($i = 1; $i -le $rows; $i++) { for ($j = 1; $j -le $cols; $j++) {
      $fi = Clean $f[$i, $j]
      $vi = Typed $vals[$i, $j]
      if ($fi -ne '' -or $vi -ne '') { [Console]::Out.WriteLine("cell`t$i`t$j`t$fi`t$vi") }
    } }
  }
  $cell = $xl.ActiveCell
  Out 'at' ("" + $cell.Row + " " + $cell.Column)
} catch { Out 'error' $_.Exception.Message }
"""

# macOS: „Not authorized to send Apple events“, -1743.
_DENIED = re.compile(r"-1743|not authori[sz]ed|nicht berechtigt", re.I)
# Excel ist beschäftigt (jemand tippt in einer Zelle, ein Dialog ist offen): später noch einmal.
_BUSY = re.compile(r"-1712|timed out|0x8001010A|0x80010001|RPC_E_CALL_REJECTED|RPC_E_SERVERCALL_RETRYLATER|rejected by callee", re.I)


@dataclass
class Snapshot:
    """Was Excel gerade zeigt."""

    book: str
    path: str = ""
    sheet: str = ""
    sheets: list[str] = field(default_factory=list)
    # None: Die Zellen ließen sich nicht lesen (dann bleibt nur die Datei).
    cells: dict[tuple[int, int], Cell] | None = None
    # Gelesener Bereich (oben, links, unten, rechts), wenn nur das Fenster gelesen wurde (großes Blatt).
    window: tuple[int, int, int, int] | None = None
    at: tuple[int, int] | None = None
    # Ist Excel vorne? None: unbekannt.
    front: bool | None = None


@dataclass
class Status:
    # off | closed | connecting | denied | empty | busy | ok | error
    state: str
    text: str = ""


def _number(value) -> str:
    if isinstance(value, bool):
        return "WAHR" if value else "FALSCH"
    if isinstance(value, (int, float)):
        return number_text(float(value))
    if isinstance(value, dict) and "d" in value:
        return str(value["d"])
    return " ".join(str(value).split())[:200]


def _cells(start_row: int, start_col: int, formulas: list, values: list) -> dict[tuple[int, int], Cell]:
    cells = {}
    for i, row in enumerate(formulas):
        for j, formula in enumerate(row):
            value = values[i][j] if i < len(values) and j < len(values[i]) else ""
            formula = str(formula or "")
            text = _number(value) if value not in ("", None) else ""
            if formula.startswith("=") and len(formula) > 1:
                cells[(start_row + i, start_col + j)] = Cell(text, formula[1:].strip()[:300])
            elif text:
                cells[(start_row + i, start_col + j)] = Cell(text)
    return cells


def parse_mac(output: str) -> tuple[Snapshot | None, Status]:
    """Die Antwort des AppleScripts als Zellen und Zustand."""
    try:
        data = json.loads(output or "{}")
    except ValueError:
        return None, Status("error", "Excel hat unverständlich geantwortet.")
    if not data.get("running"):
        return None, Status("closed")
    error = str(data.get("error") or data.get("nobook") or "")
    if _DENIED.search(error):
        return None, Status("denied", error)
    if _BUSY.search(error):
        return None, Status("busy", error)
    if not data.get("book"):
        # Ohne offene Mappe kommt AppleScript nicht an „active workbook“.
        return None, (Status("error", error) if data.get("error") else Status("empty"))
    snapshot = Snapshot(
        str(data["book"]), str(data.get("path") or ""), str(data.get("sheet") or ""),
        [str(s) for s in data.get("sheets") or []], front=data.get("front") if isinstance(data.get("front"), bool) else None,
    )
    if isinstance(data.get("formulas"), list) and isinstance(data.get("row"), int) and isinstance(data.get("col"), int):
        formulas = data["formulas"]
        snapshot.cells = _cells(data["row"], data["col"], formulas, data.get("values") or [])
        width = max((len(r) for r in formulas if isinstance(r, list)), default=1)
        snapshot.window = (data["row"], data["col"], data["row"] + max(len(formulas), 1) - 1, data["col"] + width - 1)
    at = data.get("at")
    if isinstance(at, list) and len(at) == 2 and all(isinstance(a, int) for a in at):
        snapshot.at = (at[0], at[1])
    return snapshot, Status("ok", error)


def _typed(value: str) -> str:
    kind, _, rest = value.partition(":")
    if kind == "n":
        try:
            return number_text(float(rest))
        except ValueError:
            return rest
    if kind == "d":
        try:
            return date_text(datetime.strptime(rest, "%d.%m.%Y %H:%M:%S" if rest.count(":") == 2 else "%d.%m.%Y %H:%M"))
        except ValueError:
            return rest
    if kind == "b":
        return "WAHR" if rest == "1" else "FALSCH"
    return rest


def parse_windows(output: str) -> tuple[Snapshot | None, Status]:
    """Die Antwort des PowerShell-Skripts als Zellen und Zustand."""
    fields: dict[str, list[str]] = {}
    cells: dict[tuple[int, int], Cell] = {}
    for line in (output or "").splitlines():
        kind, _, rest = line.partition("\t")
        if kind == "cell":
            parts = rest.split("\t")
            if len(parts) == 4 and parts[0].isdigit() and parts[1].isdigit():
                formula, value = parts[2], _typed(parts[3])
                position = (int(parts[0]), int(parts[1]))
                if formula.startswith("=") and len(formula) > 1:
                    cells[position] = Cell(value, formula[1:].strip()[:300])
                elif value:
                    cells[position] = Cell(value)
        elif kind:
            fields.setdefault(kind, []).append(rest)
    if fields.get("running", ["0"])[0] != "1":
        return None, Status("closed")
    error = (fields.get("error") or [""])[0]
    if _BUSY.search(error):
        return None, Status("busy", error)
    if not fields.get("book"):
        return None, (Status("error", error) if error else Status("empty"))
    snapshot = Snapshot(fields["book"][0], (fields.get("path") or [""])[0], (fields.get("sheet") or [""])[0], fields.get("sheets", []))
    try:
        top, left = int(fields["row"][0]), int(fields["col"][0])
        snapshot.cells = {(top + r - 1, left + c - 1): cell for (r, c), cell in cells.items()}
        rows, cols = (int(n) for n in fields["size"][0].split())
        snapshot.window = (top, left, top + rows - 1, left + cols - 1)
    except (KeyError, ValueError, IndexError):
        pass
    at = (fields.get("at") or [""])[0].split()
    if len(at) == 2 and all(a.isdigit() for a in at):
        snapshot.at = (int(at[0]), int(at[1]))
    return snapshot, Status("ok", error)


def _hfs(path: str) -> str:
    """„Macintosh HD:Users:anna:Mappe.xlsx“ (alte Mac-Schreibweise) → „/Users/anna/Mappe.xlsx“."""
    if path.startswith("/") or "://" in path or re.match(r"^[A-Za-z]:[\\/]", path) or ":" not in path:
        return path
    return "/" + "/".join(path.split(":")[1:])


def onedrive_roots(home: Path | None = None) -> list[Path]:
    home = home or Path.home()
    roots = []
    for variable in ("OneDrive", "OneDriveCommercial", "OneDriveConsumer"):
        if os.environ.get(variable):
            roots.append(Path(os.environ[variable]))
    for folder in (home / "Library" / "CloudStorage", home):
        try:
            roots += sorted(p for p in folder.iterdir() if p.name.startswith("OneDrive") and p.is_dir())
        except OSError:
            continue
    return list(dict.fromkeys(roots))


def cloud_denied(home: Path | None = None) -> bool:
    """Darf Gegenprobe den Ordner mit OneDrive und anderen Cloud-Speichern nicht öffnen (macOS fragt dafür)?"""
    folder = (home or Path.home()) / "Library" / "CloudStorage"
    try:
        next(folder.iterdir(), None)
    except PermissionError:
        return True
    except OSError:
        return False
    return False


def local_path(full_name: str, home: Path | None = None) -> Path | None:
    """Wo die offene Mappe auf diesem Rechner liegt, auch wenn Excel eine OneDrive-Adresse nennt."""
    name = _hfs(full_name.strip())
    if not name:
        return None
    if "://" not in name:
        path = Path(name)
        return path if path.is_absolute() and path.is_file() else None
    address = urlparse(name)
    parts = [unquote(p) for p in address.path.split("/") if p]
    if address.netloc.lower() == "d.docs.live.net":
        rest = parts[1:]  # vorne steht die Nummer des Kontos
    elif "Documents" in parts:
        rest = parts[parts.index("Documents") + 1 :]
    else:
        rest = parts[-1:]
    if not rest:
        return None
    roots = onedrive_roots(home)
    for root in roots:
        candidate = root.joinpath(*rest)
        if candidate.is_file():
            return candidate
    return _find(roots, rest[-1])


def _find(roots: list[Path], file_name: str) -> Path | None:
    """Die neueste Datei mit diesem Namen in den OneDrive-Ordnern."""
    found, seen = [], 0
    for root in roots:
        for folder, dirs, files in os.walk(root):
            seen += len(files) + len(dirs)
            dirs[:] = [d for d in dirs if not d.startswith(".")]
            if file_name in files:
                found.append(Path(folder) / file_name)
            if seen > MAX_SEARCH:
                break
    try:
        return max(found, key=lambda p: p.stat().st_mtime) if found else None
    except OSError:
        return None


class ExcelLink:
    """Fragt das laufende Excel. Ein Objekt für die ganze App, read() aus einem Thread."""

    def __init__(self, platform: str = sys.platform, run=subprocess.run) -> None:
        self.platform = platform
        self._run = run
        self.status = Status("off" if not self.supported else "closed")
        self.connected = False
        # Große Blätter (zu viele Zellen oder zu langsam): nur das Fenster lesen.
        self._window_only = False
        self._busy = 0
        self._file: tuple[str, float, Book | None] = ("", 0.0, None)
        self.file_note = ""

    @property
    def supported(self) -> bool:
        return self.platform == "darwin" or self.platform.startswith("win")

    @property
    def interval(self) -> float:
        return INTERVAL_WINDOWS if self.platform.startswith("win") else INTERVAL

    def _call(self, command: list[str], timeout: float) -> str:
        options = {}
        if self.platform.startswith("win"):
            options["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
        result = self._run(command, capture_output=True, timeout=timeout, **options)
        output = result.stdout.decode("utf-8", "replace") if isinstance(result.stdout, bytes) else result.stdout or ""
        if result.returncode != 0 and not output.strip():
            error = result.stderr.decode("utf-8", "replace") if isinstance(result.stderr, bytes) else result.stderr or ""
            raise RuntimeError(error.strip() or f"Fehlercode {result.returncode}")
        return output

    def running(self) -> bool:
        """Läuft Excel? Billig nachgesehen, ohne Excel zu fragen (und ohne macOS-Nachfrage)."""
        if not self.supported:
            return False
        try:
            if self.platform == "darwin":
                return self._run(["pgrep", "-x", "Microsoft Excel"], capture_output=True, timeout=5).returncode == 0
            output = self._call(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"], 10)
            return "EXCEL.EXE" in output.upper()
        except (OSError, subprocess.SubprocessError, RuntimeError):
            return False

    def read(self, full: bool = False) -> Snapshot | None:
        """Das aktive Blatt der offenen Mappe, None wenn es nichts zu lesen gibt (status sagt, warum).

        full: das ganze Blatt, auch wenn es groß ist (für eine Pivot-Übersicht aller Daten)."""
        if not self.supported:
            self.status = Status("off")
            return None
        timeout = TIMEOUT if self.connected else FIRST_TIMEOUT
        if full:
            timeout = max(timeout, FULL_TIMEOUT)
        mode = "visible" if self._window_only and not full else "used"
        started, known = time.monotonic(), self.connected
        try:
            if self.platform == "darwin":
                output = self._call(["osascript", "-l", "JavaScript", "-e", _MAC_SCRIPT, mode], timeout)
                snapshot, status = parse_mac(output)
            else:
                script = f"$mode = '{mode}'\n{_WINDOWS_SCRIPT}"
                encoded = base64.b64encode(script.encode("utf-16-le")).decode()
                output = self._call(["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded], timeout)
                snapshot, status = parse_windows(output)
        except subprocess.TimeoutExpired:
            self._busy += 1
            if self.connected and mode == "used" and self._busy >= MAX_BUSY and not full:
                # Immer wieder zu langsam: wohl ein großes Blatt, ab jetzt nur noch das Fenster.
                self._window_only = True
            self.status = Status("busy" if self.connected else "error", "Excel antwortet nicht.")
            return None
        except (OSError, RuntimeError) as err:
            text = str(err)
            state = "denied" if _DENIED.search(text) else "busy" if _BUSY.search(text) else "error"
            self.status = Status(state, text[:300])
            return None
        self.status = status
        if snapshot is None:
            return None
        self.connected, self._busy = True, 0
        # Das erste Mal zählt nicht als langsam: Da wartet macOS vielleicht auf „Erlauben“.
        slow = known and time.monotonic() - started > SLOW
        if snapshot.cells is not None and mode == "used" and not full and (len(snapshot.cells) > MAX_LIVE_CELLS or slow):
            self._window_only = True
        if mode == "used":
            # Das ganze Blatt: Was außerhalb fehlt, ist gelöscht.
            snapshot.window = None
        return snapshot

    def book(self, snapshot: Snapshot) -> Book | None:
        """Die gespeicherte Datei der Mappe, neu gelesen, wenn Excel sie gespeichert hat."""
        if not re.search(r"[/\\:]", snapshot.path):
            # Noch nie gespeichert („Mappe1“).
            self.file_note = "unsaved"
            return None
        try:
            path = local_path(snapshot.path)
        except PermissionError:
            self.file_note = "denied"
            return None
        if path is None:
            # Eine OneDrive-Adresse ohne Datei auf dem Rechner: nur in der Cloud, oder der Ordner ist gesperrt.
            cloud = "://" in snapshot.path
            self.file_note = "denied" if cloud and cloud_denied() else "cloud" if cloud else "unknown"
            return None
        if path.suffix.lower() not in (".xlsx", ".xlsm"):
            self.file_note = "format"
            return None
        try:
            stat = path.stat()
            if stat.st_size > MAX_FILE:
                self.file_note = "big"
                return None
            known_path, known_time, known = self._file
            if known_path == str(path) and known_time == stat.st_mtime:
                return known
            if time.time() - stat.st_mtime < SETTLE:
                return known if known_path == str(path) else None
            book = read_book(path.read_bytes())
        except PermissionError:
            self.file_note = "denied"
            return None
        except (OSError, ValueError):
            # Excel schreibt gerade, beim nächsten Mal.
            return self._file[2] if self._file[0] == str(path) else None
        self.file_note = ""
        self._file = (str(path), stat.st_mtime, book)
        return book
