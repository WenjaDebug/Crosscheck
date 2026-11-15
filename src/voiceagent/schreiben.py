"""In Excel schreiben: immer ein neues Blatt, nie in bestehende Zellen.

Am Mac über AppleScript (osascript), unter Windows über PowerShell und COM. Beides ist schon im
System, es kommt kein Paket dazu. Excel bekommt die Zellen so, wie Range.Formula sie annimmt
(englische Formeln, Text mit ' davor, wenn Excel ihn sonst als Zahl lesen würde).

Danach liest Gegenprobe das neue Blatt zurück und vergleicht jede Zahl mit dem, was sie selbst
ausgerechnet hat. So merkt er, wenn Excel etwas anders macht als erwartet.

Eine echte Pivot-Tabelle von Excel ist ein Versuch: Unter Windows über COM (PivotCaches), am Mac
über AppleScript, das Microsoft dafür kaum beschreibt. Klappt es nicht, schreibt Gegenprobe die
Übersicht mit Formeln in dasselbe neue Blatt.
"""

from __future__ import annotations

import json
import logging
import os
import re
import subprocess
import tempfile
from dataclasses import dataclass

from .excel import _BUSY, _DENIED, ExcelLink, Snapshot
from .pivot import Draft, show
from .workbook import _as_number
from .xlsx import column_letters

logger = logging.getLogger(__name__)

# So viele Zellen schickt Gegenprobe auf einmal (große Listen mag AppleScript nicht).
CHUNK = 400
# Fällt Excel auf Zelle für Zelle zurück, darf es dauern.
TIMEOUT = 180.0
# Höchstens so viele abweichende Zellen nennen.
MAX_WRONG = 3


@dataclass
class Written:
    ok: bool
    # ok | denied | busy | closed | nobook | wrongbook | error
    state: str
    text: str
    sheet: str = ""
    pivot: bool = False
    pivot_error: str = ""


def _as(text: str) -> str:
    """Ein Text als AppleScript-Ausdruck, nur aus ASCII: Umlaute als (character id 228).

    So kommt jeder Text heil an, egal wie osascript das Skript liest. Solche Ausdrücke dürfen
    nicht in einem „tell application“ stehen (dort meint „character“ etwas von Excel)."""
    parts, run = [], ""
    for ch in text:
        if 32 <= ord(ch) < 127 or ch in "\t\n\r":
            run += ch
        else:
            if run:
                parts.append(run)
                run = ""
            parts.append(ord(ch))
    if run or not parts:
        parts.append(run)

    def literal(chunk: str) -> str:
        escaped = chunk.replace("\\", "\\\\").replace('"', '\\"').replace("\r", "\\r").replace("\n", "\\n").replace("\t", "\\t")
        return f'"{escaped}"'

    pieces = [literal(p) if isinstance(p, str) else f"(character id {p})" for p in parts]
    if len(pieces) == 1:
        return pieces[0]
    if isinstance(parts[0], int):
        pieces.insert(0, '""')
    return "(" + " & ".join(pieces) + ")"


def _as_list(rows: list[list[str]]) -> str:
    return "{" + ", ".join("{" + ", ".join(_as(v) for v in row) + "}" for row in rows) + "}"


def blocks(draft: Draft) -> list[tuple[str, list[list[str]], int]]:
    """Die Zellen in Blöcken zu höchstens CHUNK Zellen: (Bereich, Zeilen, erste Zeile)."""
    width = max(len(r) for r in draft.grid)
    rows = [r + [""] * (width - len(r)) for r in draft.grid]
    size = max(1, CHUNK // width)
    result = []
    for start in range(0, len(rows), size):
        part = rows[start : start + size]
        address = f"A{start + 1}:{column_letters(width)}{start + len(part)}"
        result.append((address, part, start + 1))
    return result


_MAC_HELPERS = r"""
on putBlock(ws, addr, gridRows, firstRow, letterList)
	tell application "Microsoft Excel"
		try
			set formula of range addr of ws to gridRows
		on error
			-- Zelle für Zelle, wenn Excel den ganzen Block nicht nimmt.
			repeat with i from 1 to count of gridRows
				set rowItems to item i of gridRows
				repeat with j from 1 to count of rowItems
					set v to item j of rowItems
					if v is not "" then set formula of range ((item j of letterList) & ((firstRow + i - 1) as text)) of ws to v
				end repeat
			end repeat
		end try
	end tell
end putBlock

on newSheet(wb, newName)
	tell application "Microsoft Excel"
		try
			set ws to make new worksheet at end of wb
		on error
			try
				set ws to make new worksheet at beginning of wb
			on error
				set ws to make new worksheet
			end try
		end try
		set madeName to name of ws
		set renamed to false
		try
			tell ws to set name to newName
			set renamed to true
		end try
		-- Nach dem Umbenennen gilt nur noch der neue Name.
		if renamed then
			set ws to worksheet newName of wb
		else
			set ws to worksheet madeName of wb
		end if
		try
			activate object ws
		end try
		return ws
	end tell
end newSheet
"""

# osascript: Das Skript ließ sich nicht übersetzen (ein Wort, das Excel nicht kennt).
_COMPILE = re.compile(r"-27(40|41|50|53)|syntax error|Syntaxfehler|Expected |erwartet", re.I)


def applescript(draft: Draft, book: str, real_pivot: bool = False) -> str:
    """Das AppleScript für das neue Blatt. Texte stehen als Variablen vor den tell-Blöcken."""
    spec = draft.pivot if real_pivot else None
    width = max(len(r) for r in draft.grid)
    letters = [column_letters(c) for c in range(1, width + 1)]
    lines = [_MAC_HELPERS, "on run", '\tset out to ""', f"\tset bookName to {_as(book)}", f"\tset newName to {_as(draft.name)}"]
    if spec is not None:
        lines.append(f"\tset sourceName to {_as(spec['sheet'])}")
        names = [*spec["rows"], *spec["columns"], *(d["field"] for d in spec["data"])]
        for i, name in enumerate(names):
            lines.append(f"\tset field{i} to {_as(name)}")
    lines += ['\ttell application "Microsoft Excel"', "\t\tset wb to active workbook",
              '\t\tif name of wb is not bookName then return "wrongbook" & tab & (name of wb)', "\tend tell",
              "\tset ws to my newSheet(wb, newName)", "\tset pivotDone to false"]
    if spec is not None:
        count = len(spec["rows"])
        rows = ", ".join(f"field{i}" for i in range(count))
        columns = ", ".join(f"field{count + i}" for i in range(len(spec["columns"])))
        first_data = count + len(spec["columns"])
        fields = [f"add fields to pivot table pt row fields {{{rows}}}" + (f" column fields {{{columns}}}" if columns else "")]
        for i in range(len(spec["data"])):
            fields.append(f"set pivot field orientation of pivot field field{first_data + i} of pt to orient as data field")
        body = "\n".join(f"\t\t\t{line}" for line in fields)
        lines.append(f"""\ttell application "Microsoft Excel"
		try
			set src to range "{spec["range"]}" of worksheet sourceName of wb
			set pt to make new pivot table at beginning of ws with properties {{source data:src}}
{body}
			set pivotDone to true
		on error errText
			set out to out & "pivoterror" & tab & errText & linefeed
			try
				set display alerts to false
				delete ws
				set display alerts to true
			end try
		end try
	end tell
	if not pivotDone then set ws to my newSheet(wb, newName)""")
    lines.append("\tif not pivotDone then")
    lines.append("\t\tset letterList to {" + ", ".join(f'"{c}"' for c in letters) + "}")
    for address, part, top in blocks(draft):
        lines.append(f'\t\tmy putBlock(ws, "{address}", {_as_list(part)}, {top}, letterList)')
    lines.append('\t\ttell application "Microsoft Excel"')
    for address in draft.bold:
        lines.append(f'\t\t\ttry\n\t\t\t\tset bold of font object of range "{address}" of ws to true\n\t\t\tend try')
    for address, number_format in draft.formats:
        lines.append(f'\t\t\ttry\n\t\t\t\tset number format of range "{address}" of ws to "{number_format}"\n\t\t\tend try')
    lines.append(f'\t\t\ttry\n\t\t\t\tautofit columns of range "A:{letters[-1]}" of ws\n\t\t\tend try')
    lines.append("\t\tend tell\n\tend if")
    lines.append('\ttell application "Microsoft Excel"\n\t\ttry\n\t\t\tcalculate\n\t\tend try\n'
                 '\t\tset out to out & "sheet" & tab & (name of ws) & linefeed\n\tend tell')
    lines.append('\treturn out & "pivot" & tab & (pivotDone as text)\nend run')
    return "\n".join(lines)


_WINDOWS_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
function Out($k, $v) { [Console]::Out.WriteLine($k + "`t" + (([string]$v) -replace "[`t`r`n]", " ")) }
$data = $json | ConvertFrom-Json
function NewSheet($book) {
  $ws = $book.Worksheets.Add([Type]::Missing, $book.Worksheets.Item($book.Worksheets.Count))
  try { $ws.Name = $data.name } catch {}
  return $ws
}
try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { Out 'state' 'closed'; exit 0 }
try {
  $book = $xl.ActiveWorkbook
  if ($book -eq $null) { Out 'state' 'nobook'; exit 0 }
  if ($book.Name -ne $data.book) { Out 'wrongbook' $book.Name; exit 0 }
  $ws = NewSheet $book
  $done = $false
  if ($data.pivot -ne $null) {
    try {
      $src = $book.Worksheets.Item($data.pivot.sheet).Range($data.pivot.range)
      $cache = $book.PivotCaches().Create(1, $src)
      $pt = $cache.CreatePivotTable($ws.Range('A3'), 'Gegenprobe' + [DateTime]::Now.ToString('HHmmss'))
      $pos = 1
      foreach ($f in @($data.pivot.rows)) { $field = $pt.PivotFields($f); $field.Orientation = 1; $field.Position = $pos; $pos++ }
      foreach ($f in @($data.pivot.columns)) { $field = $pt.PivotFields($f); $field.Orientation = 2 }
      foreach ($d in @($data.pivot.data)) { [void]$pt.AddDataField($pt.PivotFields($d.field), $d.caption, [int]$d.function) }
      $ws.Range('A1').Value2 = $data.title
      $done = $true
    } catch {
      Out 'pivoterror' $_.Exception.Message
      try { $xl.DisplayAlerts = $false; $ws.Delete(); $xl.DisplayAlerts = $true } catch {}
      $ws = NewSheet $book
    }
  }
  if (-not $done) {
    foreach ($b in @($data.blocks)) {
      $rows = @($b.rows)
      $n = $rows.Count
      $m = @($rows[0]).Count
      $arr = New-Object 'object[,]' $n, $m
      for ($i = 0; $i -lt $n; $i++) { $r = @($rows[$i]); for ($j = 0; $j -lt $m; $j++) { $arr[$i, $j] = [string]$r[$j] } }
      $ws.Range($b.address).Formula = $arr
    }
    foreach ($r in @($data.bold)) { try { $ws.Range($r).Font.Bold = $true } catch {} }
    foreach ($f in @($data.formats)) { try { $ws.Range($f[0]).NumberFormat = $f[1] } catch {} }
    try { [void]$ws.Range($data.columns).EntireColumn.AutoFit() } catch {}
  }
  try { $xl.Calculate() } catch {}
  try { $ws.Activate() } catch {}
  Out 'sheet' $ws.Name
  Out 'pivot' ([int]$done)
} catch { Out 'error' $_.Exception.Message }
"""


def powershell(draft: Draft, book: str, real_pivot: bool = False) -> str:
    """Das PowerShell-Skript für das neue Blatt. Die Daten stehen als JSON darin (nur ASCII)."""
    width = column_letters(max(len(r) for r in draft.grid))
    data = {
        "book": book, "name": draft.name, "title": draft.title,
        "blocks": [{"address": address, "rows": part} for address, part, _ in blocks(draft)],
        "bold": draft.bold, "formats": [list(f) for f in draft.formats], "columns": f"A:{width}",
        "pivot": draft.pivot if real_pivot else None,
    }
    payload = json.dumps(data, ensure_ascii=True)
    return f"$json = @'\n{payload}\n'@\n{_WINDOWS_SCRIPT}"


def _fields(output: str) -> dict[str, str]:
    fields: dict[str, str] = {}
    for line in (output or "").splitlines():
        key, _, value = line.partition("\t")
        if key:
            fields[key.strip()] = value.strip()
    return fields


class ExcelWriter:
    """Schreibt ein neues Blatt in die offene Mappe. link: das ExcelLink der App (Plattform und Aufruf)."""

    def __init__(self, link: ExcelLink) -> None:
        self.link = link

    def write(self, draft: Draft, book: str, real_pivot: bool = False) -> Written:
        link = self.link
        if not link.supported:
            return Written(False, "off", "In Excel schreiben geht nur am Mac und unter Windows.")
        compile_error = ""
        try:
            if link.platform == "darwin":
                try:
                    output = self._mac(applescript(draft, book, real_pivot))
                except RuntimeError as err:
                    if not (real_pivot and _COMPILE.search(str(err))):
                        raise
                    # Kennt Excel die Wörter für Pivot-Tabellen nicht, fehlt dem Skript nur dieser Teil.
                    compile_error = "AppleScript kennt die Befehle für Pivot-Tabellen nicht"
                    output = self._mac(applescript(draft, book, False))
            else:
                output = self._windows(powershell(draft, book, real_pivot))
        except subprocess.TimeoutExpired:
            return Written(False, "busy", "Excel hat nicht rechtzeitig geantwortet. Schau nach, ob das neue Blatt ganz da ist.")
        except (OSError, RuntimeError) as err:
            text = str(err)
            if _DENIED.search(text):
                return Written(False, "denied", "Ich darf Excel nicht steuern. Erlaube es in den Systemeinstellungen unter "
                                                "Datenschutz und Sicherheit, Automation: bei Terminal Microsoft Excel einschalten.")
            if _BUSY.search(text):
                return Written(False, "busy", "Excel ist gerade beschäftigt. Beende die Eingabe in der Zelle (Enter) "
                                              "und versuch es noch einmal.")
            logger.warning("In Excel schreiben ging nicht: %s", text[:300])
            return Written(False, "error", f"Excel hat das Schreiben abgelehnt: {text[:200]}")
        fields = _fields(output)
        if "wrongbook" in fields:
            return Written(False, "wrongbook", f"In Excel ist jetzt die Mappe „{fields['wrongbook']}“ vorne, nicht "
                                               f"„{book}“. Ich schreibe nur in die Mappe aus der Vorschau.")
        if fields.get("state") in ("closed", "nobook"):
            return Written(False, fields["state"], "Excel ist zu oder hat keine offene Mappe.")
        if "error" in fields:
            return Written(False, "error", f"Excel hat das Schreiben abgelehnt: {fields['error'][:200]}")
        if "sheet" not in fields:
            return Written(False, "error", "Excel hat nicht gesagt, ob das neue Blatt da ist.")
        pivot = fields.get("pivot", "").lower() in ("1", "true")
        return Written(True, "ok", "", sheet=fields["sheet"], pivot=pivot, pivot_error=fields.get("pivoterror", "") or compile_error)

    def _mac(self, script: str) -> str:
        result = self.link._run(["osascript", "-"], input=script.encode("utf-8"), capture_output=True, timeout=TIMEOUT)
        return self._output(result)

    def _windows(self, script: str) -> str:
        handle, path = tempfile.mkstemp(suffix=".ps1", prefix="voiceagent-")
        try:
            with os.fdopen(handle, "w", encoding="utf-8-sig") as file:
                file.write(script)
            command = ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", path]
            options = {"creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)}
            result = self.link._run(command, capture_output=True, timeout=TIMEOUT, **options)
            return self._output(result)
        finally:
            try:
                os.unlink(path)
            except OSError:
                pass

    @staticmethod
    def _output(result) -> str:
        output = result.stdout.decode("utf-8", "replace") if isinstance(result.stdout, bytes) else result.stdout or ""
        if result.returncode != 0 and not output.strip():
            error = result.stderr.decode("utf-8", "replace") if isinstance(result.stderr, bytes) else result.stderr or ""
            raise RuntimeError(error.strip() or f"Fehlercode {result.returncode}")
        return output


def verify(draft: Draft, snapshot: Snapshot | None, sheet: str) -> tuple[int, list[str]]:
    """Das neue Blatt zurücklesen: Wie viele Zahlen stimmen, und welche nicht (Adresse, Excel, Gegenprobe)."""
    if snapshot is None or snapshot.cells is None or snapshot.sheet != sheet:
        return 0, []
    right, wrong = 0, []
    for (row, col), want in sorted(draft.expected.items()):
        cell = snapshot.cells.get((row, col))
        got = cell.value if cell is not None else ""
        if want == "":
            same = got.strip() == ""
        elif isinstance(want, str):
            # Text wie das Mitlesen ihn liefert: ohne doppelte Leerzeichen, höchstens 200 Zeichen.
            same = " ".join(got.split()) == " ".join(want.split())[:200]
        else:
            number = _as_number(got) if got else None
            same = number is not None and abs(number - want) <= 1e-6 * max(1.0, abs(want))
        if same:
            right += 1
        else:
            wrong.append(f"{column_letters(col)}{row}: Excel zeigt „{got or 'leer'}“, erwartet war {show(want) or 'leer'}")
    return right, wrong
