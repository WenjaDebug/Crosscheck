"""Das Excel-Add-in: der Prüfer als Seitenfenster direkt in Excel.

Excel zeigt ein Add-in als kleine Webseite rechts neben der Mappe. Diese Seite liefert Gegenprobe selbst
von diesem Rechner aus (https://localhost:8766), sie liest die Mappe über Excel (Office.js) und schickt
die Zellen nur an Gegenprobe auf demselben Rechner. Nichts davon geht ins Internet; nur office.js selbst
lädt Excel von Microsoft.

Microsoft verlangt für Add-ins eine gesicherte Verbindung. Darum legt die Einrichtung einmal eine eigene,
lokale Zertifizierungsstelle an, der nur dieser Rechner vertraut, und damit ein Zertifikat für localhost.
Einrichten am Mac: `python -m voiceagent.addin` (oder ExcelAddin.command doppelklicken), unter Windows ebenso
(oder ExcelAddin.bat doppelklicken; was dort anders ist, steht in windows.py).
"""

from __future__ import annotations

import argparse
import errno
import hashlib
import os
import re
import secrets
import shutil
import ssl
import subprocess
import sys
import tempfile
import threading
from datetime import datetime, timedelta, timezone
from pathlib import Path

from . import bezuege, dienst, formel, windows, zertifikat
from .xlsx import (MAX_BOOK_CELLS, MAX_CELLS, MAX_HIDDEN, MAX_MERGED, MAX_PARTS, MAX_SHEETS, AutoFilter, Both, Cell, Cells, Chart, ExcelTable,
                   FilterColumn, Pivot, PivotValue, Series, Sheet, _clean, _date_text, area_of, chart_kind, hide_rows, kept_text,
                   number_text, sheet_of)

PORT = 8766
ORIGIN = f"https://localhost:{PORT}"
# Fest, damit Excel das Add-in nach einer neuen Version wiedererkennt.
ADDIN_ID = "6f1d2c7e-3b8a-4f5e-9a41-2d7c0b9e5a13"
# Länger nimmt macOS ein Serverzertifikat nicht an; vorher wird es still erneuert.
CERT_DAYS = 397
RENEW_DAYS = 30
MAX_FORMULA = 2000
# So viele Namen aus dem Namensmanager nimmt Gegenprobe höchstens an.
MAX_NAMES = 500
# So heißt die eigene Zertifizierungsstelle (wie in _CA_CONFIG); so steht sie auch in der Sicherheitsfrage von Windows.
CA_NAME = "VoiceAgent lokal (nur dieser Rechner)"


def folder() -> Path:
    """Hier liegen Zertifikate und Schlüssel, unabhängig vom Datenordner (auch für data/uebung)."""
    override = os.environ.get("VOICEAGENT_ADDIN_DIR")
    return Path(override) if override else Path.home() / ".voiceagent" / "addin"


def wef_folder() -> Path:
    """Der Ordner, in dem Excel am Mac selbst hinzugefügte Add-ins sucht."""
    return Path.home() / "Library" / "Containers" / "com.microsoft.Excel" / "Data" / "Documents" / "wef"


# --- Zertifikate ---

_CA_CONFIG = """[req]
distinguished_name = dn
prompt = no
[dn]
CN = VoiceAgent lokal (nur dieser Rechner)
[ca]
basicConstraints = critical, CA:TRUE
keyUsage = critical, keyCertSign, cRLSign
subjectKeyIdentifier = hash
"""

_LEAF_CONFIG = """[req]
distinguished_name = dn
prompt = no
[dn]
CN = localhost
"""

_LEAF_EXTENSIONS = """basicConstraints = CA:FALSE
keyUsage = critical, digitalSignature, keyEncipherment
extendedKeyUsage = serverAuth
subjectAltName = DNS:localhost, IP:127.0.0.1
subjectKeyIdentifier = hash
authorityKeyIdentifier = keyid, issuer
"""


def _pure() -> bool:
    """Zertifikate ohne openssl (zertifikat.py): unter Windows immer, sonst nur, wenn openssl fehlt. Am Mac bleibt
    es bei openssl, so muss ein eingerichteter Mac dem Zertifikat nicht neu vertrauen."""
    return sys.platform == "win32" or shutil.which("openssl") is None


def _openssl(*args: str, cwd: Path) -> None:
    done = subprocess.run(["openssl", *args], cwd=cwd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=60)
    if done.returncode != 0:
        raise RuntimeError(f"openssl {args[0]}: {(done.stderr or done.stdout).strip()[:300]}")


def _expires(cert: Path) -> datetime | None:
    if _pure():
        try:
            return zertifikat.expires(cert)
        except (OSError, ValueError, IndexError):
            return None
    try:
        done = subprocess.run(["openssl", "x509", "-enddate", "-noout", "-in", str(cert)], capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30)
        return datetime.fromtimestamp(ssl.cert_time_to_seconds(done.stdout.strip().partition("=")[2]), timezone.utc)
    except (OSError, ValueError, subprocess.SubprocessError):
        return None


def _linked(cert: Path) -> bool:
    """Nennt das Zertifikat seine Zertifizierungsstelle (Authority Key Identifier)? Das openssl am Mac
    (LibreSSL) schreibt das nicht von selbst hinein, und strenge Prüfer wie Python 3.13 lehnen es dann ab."""
    if _pure():
        try:
            return zertifikat.linked(cert)
        except (OSError, ValueError, IndexError):
            return False
    try:
        done = subprocess.run(["openssl", "x509", "-noout", "-text", "-in", str(cert)], capture_output=True, text=True,
                              encoding="utf-8", errors="replace", timeout=30)
    except (OSError, subprocess.SubprocessError):
        return True
    return done.returncode != 0 or "Authority Key Identifier" in done.stdout


def certificates(base: Path | None = None) -> tuple[Path, Path, Path]:
    """Zertifizierungsstelle und Serverzertifikat, wenn nötig neu. Gibt (ca, cert, key) zurück."""
    base = base or folder()
    base.mkdir(parents=True, exist_ok=True)
    ca, ca_key = base / "ca.pem", base / "ca-key.pem"
    cert, key = base / "cert.pem", base / "key.pem"
    pure = _pure()
    if not (ca.exists() and ca_key.exists()):
        if pure:
            zertifikat.authority(ca, ca_key, CA_NAME)
        else:
            (base / "ca.cnf").write_text(_CA_CONFIG, encoding="utf-8")
            _openssl("req", "-x509", "-new", "-newkey", "rsa:2048", "-nodes", "-sha256", "-days", "3650",
                     "-keyout", ca_key.name, "-out", ca.name, "-config", "ca.cnf", "-extensions", "ca", cwd=base)
        cert.unlink(missing_ok=True)
    expires = _expires(cert) if cert.exists() and key.exists() else None
    if expires is None or expires - datetime.now(timezone.utc) < timedelta(days=RENEW_DAYS) or not _linked(cert):
        if pure:
            zertifikat.server(cert, key, ca, ca_key, CERT_DAYS)
        else:
            (base / "leaf.cnf").write_text(_LEAF_CONFIG, encoding="utf-8")
            (base / "leaf.ext").write_text(_LEAF_EXTENSIONS, encoding="utf-8")
            _openssl("req", "-new", "-newkey", "rsa:2048", "-nodes", "-keyout", key.name, "-out", "leaf.csr",
                     "-config", "leaf.cnf", cwd=base)
            _openssl("x509", "-req", "-in", "leaf.csr", "-CA", ca.name, "-CAkey", ca_key.name,
                     "-set_serial", f"0x{secrets.token_hex(12)}", "-days", str(CERT_DAYS), "-sha256",
                     "-extfile", "leaf.ext", "-out", cert.name, cwd=base)
    for private in (ca_key, key):
        private.chmod(0o600)
    return ca, cert, key


def ready(base: Path | None = None) -> tuple[Path, Path] | None:
    """(cert, key), wenn das Add-in eingerichtet ist. Ein bald ablaufendes Zertifikat wird erneuert."""
    base = base or folder()
    if not (base / "ca.pem").exists():
        return None
    try:
        _, cert, key = certificates(base)
    except (OSError, RuntimeError, ValueError, subprocess.SubprocessError) as err:
        print(f"Excel-Add-in: Zertifikat ging nicht ({err}).")
        return None
    return cert, key


def serve(app, cert: Path, key: Path, port: int = PORT) -> threading.Thread:
    """Dieselbe App noch einmal mit https für das Add-in, im Hintergrund, nur auf diesem Rechner."""
    import uvicorn

    config = uvicorn.Config(app, host="127.0.0.1", port=port, log_level="warning",
                            ssl_certfile=str(cert), ssl_keyfile=str(key))
    server = uvicorn.Server(config)
    thread = threading.Thread(target=server.run, name="excel-addin", daemon=True)
    thread.start()
    return thread


# --- Manifest ---


def manifest(origin: str = ORIGIN) -> str:
    """Die Beschreibung für Excel: Name, Knopf im Menüband „Start“ und die Seite des Seitenfensters."""
    def icon(size: int) -> str:
        return f"{origin}/addin/icon-{size}.png"

    return f"""<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>{ADDIN_ID}</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Gegenprobe</ProviderName>
  <DefaultLocale>de-DE</DefaultLocale>
  <DisplayName DefaultValue="Gegenprobe Prüfer">
    <Override Locale="en-US" Value="Gegenprobe Checker"/>
  </DisplayName>
  <Description DefaultValue="Rechnet die Mappe nach und zeigt, was nicht aufgeht. Ohne KI, die Daten bleiben auf diesem Rechner.">
    <Override Locale="en-US" Value="Recalculates the workbook and shows what does not add up. No AI; the data stays on this computer."/>
  </Description>
  <IconUrl DefaultValue="{icon(32)}"/>
  <HighResolutionIconUrl DefaultValue="{icon(64)}"/>
  <SupportUrl DefaultValue="https://github.com/WenjaDebug/Crosscheck"/>
  <AppDomains>
    <AppDomain>{origin}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="{origin}/addin/pruefer.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="VA.GetStarted.Title"/>
            <Description resid="VA.GetStarted.Description"/>
            <LearnMoreUrl resid="VA.Help.Url"/>
          </GetStarted>
          <FunctionFile resid="VA.Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="VA.Group">
                <Label resid="VA.Group.Label"/>
                <Icon>
                  <bt:Image size="16" resid="VA.Icon.16"/>
                  <bt:Image size="32" resid="VA.Icon.32"/>
                  <bt:Image size="80" resid="VA.Icon.80"/>
                </Icon>
                <Control xsi:type="Button" id="VA.Pruefer">
                  <Label resid="VA.Button.Label"/>
                  <Supertip>
                    <Title resid="VA.Button.Label"/>
                    <Description resid="VA.Button.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="VA.Icon.16"/>
                    <bt:Image size="32" resid="VA.Icon.32"/>
                    <bt:Image size="80" resid="VA.Icon.80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>VoiceAgentPruefer</TaskpaneId>
                    <SourceLocation resid="VA.Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="VA.Icon.16" DefaultValue="{icon(16)}"/>
        <bt:Image id="VA.Icon.32" DefaultValue="{icon(32)}"/>
        <bt:Image id="VA.Icon.80" DefaultValue="{icon(80)}"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="VA.Commands.Url" DefaultValue="{origin}/addin/befehle.html"/>
        <bt:Url id="VA.Taskpane.Url" DefaultValue="{origin}/addin/pruefer.html"/>
        <bt:Url id="VA.Help.Url" DefaultValue="https://github.com/WenjaDebug/Crosscheck"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="VA.GetStarted.Title" DefaultValue="Gegenprobe Prüfer ist bereit">
          <bt:Override Locale="en-US" Value="Gegenprobe Checker is ready"/>
        </bt:String>
        <bt:String id="VA.Group.Label" DefaultValue="Gegenprobe"/>
        <bt:String id="VA.Button.Label" DefaultValue="Prüfer">
          <bt:Override Locale="en-US" Value="Checker"/>
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="VA.GetStarted.Description" DefaultValue="Klick unter Start auf Prüfer. Das Fenster rechnet die Mappe nach, ohne KI.">
          <bt:Override Locale="en-US" Value="Click Checker on the Home tab. The panel recalculates the workbook, without AI."/>
        </bt:String>
        <bt:String id="VA.Button.Tooltip" DefaultValue="Öffnet den Prüfer neben der Mappe. Er rechnet nach und zeigt, was nicht aufgeht.">
          <bt:Override Locale="en-US" Value="Opens the checker next to the workbook. It recalculates and shows what does not add up."/>
        </bt:String>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
"""


# --- Zellen aus Excel (Office.js) ---


def _cell_text(value, kind: str | None) -> str:
    """Ein Wert aus Office.js so, wie der Prüfer ihn auch aus der Datei kennt."""
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        return "WAHR" if value else "FALSCH"
    if isinstance(value, (int, float)):
        number = float(value)
        # Ein Datum vor März 1900 gibt es in keiner echten Liste: Das ist eine Zahl, die nur ein Uhrzeit- oder
        # Datumsformat trägt (8 Stunden als hh:mm zeigt Excel als 00:00). Gerechnet wird mit der Zahl selbst.
        if kind == "d" and not 1 <= abs(number) < 61:
            return _date_text(number, False)
        if kind == "p":
            return f"{number_text(round(number * 100, 10))} %"
        return number_text(number)
    return _clean(str(value))


class Resend(Exception):
    """Das Seitenfenster schickt ein Blatt als unverändert, das Gegenprobe nicht (mehr) kennt: alles neu schicken."""


MAX_PASST = 500


def passt_from(payload: dict) -> dict[str, frozenset[str]]:
    """Je Blatt die Hinweise, die jemand mit „Passt so“ ausgeblendet hat: {Blatt: [hint_id, …]}, so wie das Fenster
    sie in den Einstellungen der Mappe hält."""
    items = payload.get("passt")
    if not isinstance(items, dict):
        return {}
    out = {}
    for name, ids in list(items.items())[:MAX_SHEETS * 5]:
        if isinstance(ids, list):
            out[str(name)] = frozenset(i[:16] for i in ids[:MAX_PASST] if isinstance(i, str))
    return out


def sheets_from(payload: dict, known=None) -> tuple[str, str, list[Sheet]]:
    """Die Mappe, wie das Add-in sie schickt: je Blatt Startzelle, Werte, Formeln und Zahlenart.

    Formeln und Zahlenarten kommen als Liste [Zeile, Spalte, Text] nur für die Zellen, die eine haben (formulas_at,
    kinds_at), von älteren Fenstern als ganzes Raster (formulas, kinds). stamp: woran das Fenster diesen Stand des
    Blatts erkennt. Ein unverändertes Blatt kommt nur mit Namen und same (seinem stamp): known(name, stamp) liefert
    es, wie es zuletzt kam, sonst Resend. cut: Das Fenster hat nur den oberen Teil eines großen Blatts gelesen.
    Höchstens MAX_CELLS Zellen je Blatt und MAX_BOOK_CELLS in der Mappe, das aktive Blatt zuerst."""
    book = _clean(str(payload.get("book") or "Mappe"))
    active = str(payload.get("active") or "")
    items = [item for item in (payload.get("sheets") or [])[:MAX_SHEETS] if isinstance(item, dict)]
    # Das aktive Blatt bekommt seine Zellen zuerst, die übrigen in der Reihenfolge der Mappe, was dann noch geht.
    order = sorted(range(len(items)), key=lambda i: str(items[i].get("name") or "") != active)
    read: dict[int, Sheet] = {}
    total = 0
    for i in order:
        item = items[i]
        name = str(item.get("name") or "")
        if item.get("same"):
            sheet = known(name, str(item["same"])) if known is not None else None
            if sheet is None:
                raise Resend(name)
        else:
            sheet = _sheet_from(item, name, max(0, min(MAX_CELLS, MAX_BOOK_CELLS - total)))
        sheet.hidden = bool(item.get("hidden"))
        read[i] = sheet
        total += len(sheet.cells)
    sheets = [read[i] for i in range(len(items))]
    # Excel-Tabellen der Mappe (Quelle einer Pivot-Tabelle), bei ihrem Blatt.
    by_name = {s.name: s for s in sheets}
    for table in (payload.get("tables") or [])[:MAX_PARTS * 5] if isinstance(payload.get("tables"), list) else []:
        if not isinstance(table, dict) or str(table.get("sheet") or "") not in by_name:
            continue
        area = area_of(str(table.get("address") or ""))
        if area is not None and table.get("name"):
            columns = table.get("columns")
            columns = tuple(_clean(str(c), 60) for c in columns[:area[3] - area[1] + 1]) if isinstance(columns, list) else ()
            by_name[str(table["sheet"])].tables.append(
                ExcelTable(_clean(str(table["name"]), 60), area, columns, header=table.get("headers") is not False,
                           totals=1 if table.get("totals") else 0))
    # Formeln mit Tabellenbezügen und Namen (=[@Menge]*MwSt) rechnet Gegenprobe mit den Zellen, die sie meinen.
    names, local = _names(payload.get("names"))
    bezuege.resolve(sheets, names, local)
    return book, active, sheets


def _names(items) -> tuple[dict[str, str], dict[str, dict[str, str]]]:
    """Die Namen, wie das Fenster sie schickt: [{"name", "formula", "sheet" (nur bei Namen eines Blatts)}]."""
    names: dict[str, str] = {}
    local: dict[str, dict[str, str]] = {}
    for item in items[:MAX_NAMES] if isinstance(items, list) else []:
        if not isinstance(item, dict) or not isinstance(item.get("name"), str) or not isinstance(item.get("formula"), str):
            continue
        name, formula = item["name"].strip()[:60], item["formula"].strip()[:MAX_FORMULA]
        if not name or not formula or name.startswith("_xlnm."):
            continue
        if isinstance(item.get("sheet"), str) and item["sheet"]:
            local.setdefault(item["sheet"], {})[name] = formula
        else:
            names[name] = formula
    return names, local


def _sheet_from(item: dict, name: str, limit: int) -> Sheet:
    """Ein Blatt aus dem Seitenfenster mit höchstens limit Zellen, nur ganze Zeilen."""
    top, left = int(item.get("row") or 1), int(item.get("col") or 1)
    values = item.get("values") or []
    values = values if isinstance(values, list) else []
    formulas = _sparse(item.get("formulas_at")) if "formulas_at" in item else _grid(item.get("formulas"))
    kinds = _sparse(item.get("kinds_at")) if "kinds_at" in item else _grid(item.get("kinds"))
    cells = Cells()
    for i, row in enumerate(values):
        if not isinstance(row, list):
            continue
        found = []
        for j, value in enumerate(row):
            formula = formulas.get((i, j))
            text = _cell_text(value, kinds.get((i, j)))
            # Eine Formel als Text (die Zelle ist als Text formatiert): Excel gibt dann bei „Formel“ den Text zurück.
            if isinstance(formula, str) and isinstance(value, str) and formula == value:
                formula = None
            formula = formula[1:].strip()[:MAX_FORMULA] if isinstance(formula, str) and formula.startswith("=") and len(formula) > 1 else None
            if kinds.get((i, j)) == "d" and isinstance(value, (int, float)) and not isinstance(value, bool) \
                    and 1 <= abs(value) < 61:
                cells.clipped[(top + i, left + j)] = float(value)
            if kinds.get((i, j)) == "i" and isinstance(value, (int, float)) and not isinstance(value, bool) \
                    and 0 < abs(value) < 0.5:
                cells.rounded[(top + i, left + j)] = float(value)
            if text or formula:
                found.append(((top + i, left + j), Cell(text, formula, kept_text(value) if isinstance(value, str) else None)))
        if len(cells) + len(found) > limit:
            # Nur ganze Zeilen: Gelesen ist bis zur Zeile davor, 0 heißt gar nichts.
            cells.cut = top + i - 1 if cells else 0
            break
        cells.update(found)
    if cells.cut is None and item.get("cut"):
        cells.cut = top + len(values) - 1
    sheet = Sheet(name, cells, filters=_filters(item.get("filtered")), charts=_charts(item.get("charts")),
                  pivots=_pivots(item.get("pivots")), merged=_merged(item.get("merged")))
    sheet.stamp = str(item.get("stamp") or "")
    rows = item.get("hiddenRows")
    # Ein älteres Fenster schickt keine: Dann sind sie unbekannt, nicht „keine“.
    if isinstance(rows, list):
        hide_rows(sheet, (int(r) for r in rows[:MAX_HIDDEN] if isinstance(r, int) and not isinstance(r, bool) and r > 0))
    return sheet


def _merged(items) -> list[tuple[int, int, int, int]]:
    """Verbundene Bereiche, wie das Add-in sie liest („A2:A6“), ohne einzelne Zellen."""
    found = []
    for item in (items if isinstance(items, list) else [])[:MAX_MERGED]:
        area = area_of(item) if isinstance(item, str) and len(item) < 40 else None
        if area is not None and area[:2] != area[2:]:
            found.append(area)
    return found


def _texts(items, limit: int = MAX_PARTS) -> tuple[str, ...]:
    return tuple(_clean(str(i), 60) for i in (items if isinstance(items, list) else [])[:limit] if i not in (None, ""))


def _charts(items) -> list[Chart]:
    """Diagramme, wie das Add-in sie liest: Name, Titel und je Reihe die Bereiche der Werte und Beschriftungen."""
    charts = []
    for item in (items if isinstance(items, list) else [])[:MAX_PARTS]:
        if not isinstance(item, dict):
            continue
        series = tuple(
            Series(_clean(str(s.get("name") or ""), 60), "", _clean(str(s.get("values") or ""), 200),
                   _clean(str(s.get("categories") or ""), 200))
            for s in (item.get("series") if isinstance(item.get("series"), list) else [])[:MAX_PARTS] if isinstance(s, dict))
        charts.append(Chart(_clean(str(item.get("name") or "Diagramm"), 60), chart_kind([]), _clean(str(item.get("title") or ""), 80),
                            series))
    return charts


# Wie Office.js die Zusammenfassung eines Wertfelds nennt → wie die Datei sie nennt.
_SUMMARIES = {"Sum": "sum", "Count": "count", "Average": "average", "Max": "max", "Min": "min", "Product": "product",
              "CountNumbers": "countNums", "StandardDeviation": "stdDev", "StandardDeviationP": "stdDevp",
              "Variance": "var", "VarianceP": "varp"}
# Sagt Office.js „Automatic“ (so bei einer Summe, die in der Datei ohne Angabe steht), sagt es die Beschriftung, die
# Excel dem Wertfeld gibt: „Summe von Umsatz“, „Sum of Sales“. Eine eigene Beschriftung sagt nichts, dann rechnet der
# Prüfer die Pivot-Tabelle nicht nach.
_CAPTIONS = [(re.compile(rf"^\s*(?:{words})\s+(?:von|of)\s", re.I), name) for words, name in (
    ("anzahl zahlen|count numbers|countnums|count nums", "countNums"), ("summe|sum", "sum"), ("anzahl|count", "count"),
    ("mittelwert|average", "average"), (r"max\.?", "max"), (r"min\.?", "min"), ("produkt|product", "product"),
    ("stdabwn|stdevp|stddevp", "stdDevp"), ("stdabw|stdev|stddev", "stdDev"), ("varianzn|varp", "varp"),
    ("varianz|var", "var"))]


def _summary(function: str, caption: str) -> str:
    if function in _SUMMARIES:
        return _SUMMARIES[function]
    return next((name for pattern, name in _CAPTIONS if pattern.match(caption)), "")


def _pivots(items) -> list[Pivot]:
    """Pivot-Tabellen, wie das Add-in sie liest (pruefer.js, readPivot)."""
    pivots = []
    for item in (items if isinstance(items, list) else [])[:MAX_PARTS]:
        if not isinstance(item, dict):
            continue
        values = []
        for v in (item.get("values") if isinstance(item.get("values"), list) else [])[:MAX_PARTS]:
            if not isinstance(v, dict):
                continue
            shown = v.get("showAs")
            caption = _clean(str(v.get("caption") or ""), 60)
            values.append(PivotValue(caption, _clean(str(v.get("field") or ""), 60),
                                     _summary(str(v.get("function") or ""), caption), shown not in (None, "", "None")))
        source = str(item.get("source") or "").strip().lstrip("=")
        table = _clean(str(item.get("table") or ""), 60)
        skipped = item.get("skipped") if isinstance(item.get("skipped"), dict) else {}
        grand = item.get("grand") if isinstance(item.get("grand"), list) and len(item["grand"]) == 2 else [True, True]
        filtered = item.get("filtered")
        pivots.append(Pivot(
            _clean(str(item.get("name") or "PivotTable"), 60),
            area_of(str(item.get("area") or "")),
            _texts(item.get("rows")),
            _texts(item.get("columns")),
            tuple(v.caption for v in values),
            _texts(item.get("filters")),
            (),
            _clean(source, 80),
            tuple(values),
            area_of(str(item.get("body") or "")),
            _clean(sheet_of(source), 40) if not table else "",
            area_of(source) if source and not table else None,
            table,
            tuple((_clean(str(k), 60), _texts(v, 500)) for k, v in list(skipped.items())[:MAX_PARTS]),
            (bool(grand[0]), bool(grand[1])),
            None if filtered is None else bool(filtered),
        ))
    return pivots


def _filters(addresses) -> list[AutoFilter]:
    """Bereiche, in denen gerade ein Filter Zeilen ausblendet („Tabelle1!A1:F200“), wie aus der Datei."""
    found = []
    for text in (addresses if isinstance(addresses, list) else [])[:20]:
        ref = str(text).rpartition("!")[2].replace("$", "")
        try:
            area = formel.area(ref) if ":" in ref else None
        except (AttributeError, ValueError):
            area = None
        if area is not None:
            found.append(AutoFilter(area, (FilterColumn(area[1], Both("gefiltert", "filtered")),)))
    return found


def _grid(grid) -> dict[tuple[int, int], object]:
    """Ein ganzes Raster (ältere Fenster) als {(Zeile, Spalte): Wert}, ohne leere Einträge."""
    out = {}
    for i, row in enumerate(grid if isinstance(grid, list) else []):
        if isinstance(row, list):
            for j, value in enumerate(row):
                if value is not None:
                    out[(i, j)] = value
    return out


def _sparse(items) -> dict[tuple[int, int], object]:
    """[[Zeile, Spalte, Wert], …] als {(Zeile, Spalte): Wert}."""
    out = {}
    for item in items if isinstance(items, list) else []:
        if isinstance(item, list) and len(item) == 3 and isinstance(item[0], int) and isinstance(item[1], int):
            out[(item[0], item[1])] = item[2]
    return out


# --- Einrichten ---


FULL_DISK_ACCESS = "x-apple.systempreferences:com.apple.preference.security?Privacy_AllFiles"
NO_ACCESS = """macOS lässt das Terminal nicht in den Ordner von Excel schreiben. Ich öffne dir die Einstellung dafür:
  1. In „Festplattenvollzugriff“ Terminal einschalten. Fehlt es: unten auf +, dann Programme, Dienstprogramme, Terminal.
  2. Terminal ganz beenden (Befehl-Q), neu öffnen und ExcelAddin.command noch einmal starten.
Danach kannst du den Festplattenvollzugriff für Terminal wieder ausschalten, der Prüfer braucht ihn nicht."""


LOGIN_ITEMS = """    Damit er auch nach einem Neustart von selbst startet, muss macOS ihn im Hintergrund erlauben.
    Ich öffne dafür Systemeinstellungen → Allgemein → Anmeldeobjekte. Schalte dort unter
    „Im Hintergrund erlauben“ python3.13 (oder Python) ein. Das ist der Prüfer."""


def _trusted(cert: Path) -> bool:
    """Vertraut der Mac dem Zertifikat für localhost, so wie Excel es prüft?"""
    done = subprocess.run(["security", "verify-cert", "-c", str(cert), "-p", "ssl", "-s", "localhost", "-q"],
                          capture_output=True, timeout=60)
    return done.returncode == 0


def install(remove: bool = False) -> int:
    if sys.platform == "win32":
        return install_windows(remove)
    if sys.platform != "darwin":
        print("Das Excel-Add-in richte ich nur am Mac und unter Windows ein. Hier nimm den Excel-Prüfer im Browser.")
        return 1
    target = wef_folder() / "voiceagent-pruefer.xml"
    if remove:
        target.unlink(missing_ok=True)
        (folder() / "manifest.sha256").unlink(missing_ok=True)
        dienst.remove()
        print("Das Add-in ist aus Excel entfernt und läuft nicht mehr im Hintergrund. Starte Excel neu.")
        return 0
    if not wef_folder().parent.parent.parent.exists():
        print("Excel finde ich nicht. Öffne Microsoft Excel einmal, schließe es wieder und starte das hier noch einmal.")
        return 1
    print("1/4 Ich lege ein Zertifikat nur für diesen Mac an …")
    ca, cert, _ = certificates()
    if _trusted(cert):
        print("2/4 Der Mac vertraut dem Zertifikat schon.")
    else:
        print("2/4 Jetzt fragt der Mac nach deinem Passwort, damit er dem Zertifikat vertraut. Das ist nur dieses eine Mal.")
        keychain = Path.home() / "Library" / "Keychains" / "login.keychain-db"
        done = subprocess.run(["security", "add-trusted-cert", "-r", "trustRoot", "-k", str(keychain), str(ca)], timeout=300)
        if done.returncode != 0 or not _trusted(cert):
            print("Das hat nicht geklappt. Starte es noch einmal und gib im Fenster dein Mac-Passwort ein.")
            return 1
    text = manifest()
    # Was zuletzt bei Excel angemeldet wurde, steht im eigenen Ordner: Den Ordner von Excel auch nur zu lesen,
    # braucht schon die Erlaubnis von macOS.
    record = folder() / "manifest.sha256"
    digest = hashlib.sha256(text.encode("utf-8")).hexdigest()
    if record.exists() and record.read_text(encoding="utf-8").strip() == digest:
        print("3/4 Das Add-in ist bei Excel schon angemeldet.")
    else:
        print("3/4 Ich melde das Add-in bei Excel an …")
        try:
            wef_folder().mkdir(parents=True, exist_ok=True)
            with tempfile.NamedTemporaryFile("w", encoding="utf-8", dir=wef_folder(), delete=False, suffix=".tmp") as file:
                file.write(text)
            Path(file.name).replace(target)
        except OSError as err:
            if isinstance(err, PermissionError) or err.errno == errno.EPERM:
                print(NO_ACCESS)
                subprocess.run(["open", FULL_DISK_ACCESS], capture_output=True, timeout=30)
            else:
                print(f"In den Excel-Ordner darf ich nicht schreiben ({err}).")
            return 1
        record.write_text(digest, encoding="utf-8")
    print("4/4 Ich richte den Prüfer im Hintergrund ein: macOS startet ihn, sobald Excel ihn braucht, auch nach einem Neustart.")
    seen, missed = dienst.before()
    for line in seen:
        print(f"    {line}")
    refused = dienst.install(start=False)
    background = not refused
    if refused:
        print(f"macOS startet ihn nicht im Hintergrund: {refused}.")
    else:
        # Geladen wie bei der Anmeldung nach einem Neustart, ohne nachzuhelfen. Dann anklopfen wie Excel: Startet
        # macOS ihn dafür von selbst, tut es das auch nach einem Neustart.
        print("    Ich frage ihn an wie Excel und schaue, ob macOS ihn dafür von selbst startet …")
        problem = dienst.answers(timeout=30)
        nudged = ""
        if problem:
            info = dienst.state()
            if not info["pid"]:
                nudged = dienst.explain(info) if not dienst.ready(info) else "macOS hat ihn auf die Anfrage hin nicht gestartet"
            # Für später: was macOS dazu sagt, steht im Protokoll (data/pruefer-dienst.log).
            said = f"nicht von selbst gestartet ({nudged})" if nudged else f"antwortet nicht ({problem})"
            dienst.note(f"Probe: {said}. launchctl print sagt:\n{dienst.report() or '(nichts)'}")
            dienst.repair()
            problem = dienst.answers(timeout=30)
        else:
            dienst.note("Probe: auf die Anfrage hin von selbst gestartet und geantwortet. launchctl print sagt:\n"
                        f"{dienst.report() or '(nichts)'}")
        fallback = False
        if problem and dienst.state()["sockets"]:
            # Rückfall wie früher: bei der Anmeldung starten und den Port selbst öffnen. Angestoßen lief er so.
            dienst.note(f"Rückfall auf den Start bei der Anmeldung ({problem}).")
            problem = dienst.install(on_demand=False) or dienst.answers(timeout=30)
            fallback = not problem
        if problem:
            print(f"Der Prüfer im Hintergrund antwortet nicht: {problem}.")
            tail = dienst.log_tail()
            if tail:
                print("Das schreibt er ins Protokoll (data/pruefer-dienst.log):")
                for line in tail:
                    print(f"    {line}")
            print("Schick davon ein Bildschirmfoto, dann weiß ich, was fehlt.")
            return 1
        if fallback:
            print("    Auf Anfrage startete macOS ihn nicht. Er läuft jetzt wie bisher, angestoßen, nach einem Neustart")
            print("    aber vielleicht nicht. Was macOS dazu sagt, steht in data/pruefer-dienst.log.")
        elif nudged or missed:
            if nudged:
                print(f"    Von selbst startete er nicht ({nudged}). Angestoßen antwortet er jetzt.")
            else:
                print("    Jetzt startet er von selbst und antwortet.")
            print(LOGIN_ITEMS)
            subprocess.run(["open", dienst.LOGIN_ITEMS], capture_output=True, timeout=30)
        else:
            print("    Er startet von selbst, sobald Excel ihn braucht, also auch nach einem Neustart.")
    print()
    print("Fertig. So öffnest du ihn:")
    print("  1. Excel ganz beenden (Befehl-Q) und neu öffnen, eine Mappe öffnen.")
    print("  2. Im Menüband Start ganz rechts auf Add-ins klicken, dann Gegenprobe Prüfer.")
    print("     Danach gibt es unter Start den Knopf Prüfer. Ein Klick, und er steht rechts in Excel.")
    if background:
        print("Terminal und Browser brauchst du dafür nicht. Meldet der Mac ein neues Hintergrundobjekt (python): Das ist der Prüfer.")
    else:
        print("Im Hintergrund starten ließ er sich nicht. Dann starte VoiceAgent.command, solange du den Prüfer brauchst.")
    return 0


WINDOWS_TRUST = f"""2/4 Jetzt fragt Windows einmal, ob es dem Zertifikat „{CA_NAME}“ vertrauen soll.
    Klick in der Sicherheitswarnung auf Ja. Das ist nur dieses eine Mal, Adminrechte brauchst du nicht."""

WINDOWS_CERT_STAYS = f"""Das Zertifikat bleibt in Windows, damit ein neues Einrichten nicht noch einmal fragt. Herausnehmen:
Windows-Taste, „Benutzerzertifikate verwalten“, Vertrauenswürdige Stammzertifizierungsstellen, Zertifikate,
„{CA_NAME}“ löschen."""

WINDOWS_DONE = """Fertig. So öffnest du ihn:
  1. Excel ganz schließen (alle Excel-Fenster) und neu öffnen, eine Mappe öffnen.
  2. Im Menüband Start steht ganz rechts der Knopf Prüfer (Gruppe Gegenprobe). Ein Klick, und er steht rechts in Excel.
     Fehlt er: unter Start auf Add-ins (oder Einfügen, Meine Add-ins), dort bei Entwickler-Add-ins Gegenprobe Prüfer.
Excel 2016 und 2019 zeigen Add-ins noch mit dem alten Internet Explorer, dort bleibt das Fenster leer.
Dann nimm den Prüfer im Browser (VoiceAgent.bat)."""


def install_windows(remove: bool = False, reg=None, run=subprocess.run, popen=subprocess.Popen, enum=None,
                    root: Path | None = None) -> int:
    """Wie am Mac, mit den Mitteln von Windows (windows.py): Zertifikat ohne openssl, Vertrauen mit certutil,
    Anmelden bei Excel in der Registry, im Hintergrund über HKCU …\\Run."""
    target = folder() / "voiceagent-pruefer.xml"
    if remove:
        windows.unregister(ADDIN_ID, reg)
        target.unlink(missing_ok=True)
        windows.remove(reg, run)
        print("Das Add-in ist aus Excel entfernt und läuft nicht mehr im Hintergrund. Schließ Excel ganz und öffne es neu.")
        print(WINDOWS_CERT_STAYS)
        return 0
    if not windows.office_seen(reg):
        # Anmelden geht trotzdem; Excel liest es beim ersten Start.
        print("Excel 2016 oder neuer habe ich unter diesem Benutzer noch nicht gefunden. Ich richte es trotzdem ein.")
    print("1/4 Ich lege ein Zertifikat nur für diesen PC an …")
    try:
        ca, _, _ = certificates()
    except (OSError, ValueError, RuntimeError) as err:
        print(f"Das Zertifikat ließ sich nicht anlegen ({err}).")
        return 1
    if windows.trusted(ca, enum):
        print("2/4 Windows vertraut dem Zertifikat schon.")
    else:
        print(WINDOWS_TRUST)
        ok, said = windows.trust(ca, run)
        if not ok and not windows.trusted(ca, enum):
            print("Das hat nicht geklappt. Starte ExcelAddin.bat noch einmal und klick in der Sicherheitswarnung auf Ja.")
            if said:
                print(f"    Windows sagt: {said}")
            return 1
    text = manifest()
    try:
        current = target.read_text(encoding="utf-8")
    except OSError:
        current = ""
    if current == text and windows.registered(ADDIN_ID, reg) == str(target):
        print("3/4 Das Add-in ist bei Excel schon angemeldet.")
    else:
        print("3/4 Ich melde das Add-in bei Excel an …")
        try:
            with tempfile.NamedTemporaryFile("w", encoding="utf-8", dir=target.parent, delete=False, suffix=".tmp") as file:
                file.write(text)
            Path(file.name).replace(target)
            windows.register(target, ADDIN_ID, reg)
        except OSError as err:
            print(f"Bei Excel anmelden ging nicht ({err}).")
            return 1
    print("4/4 Ich richte den Prüfer im Hintergrund ein: Windows startet ihn bei jeder Anmeldung, ohne Fenster.")
    refused = windows.install(root or Path.cwd(), reg=reg, run=run, popen=popen)
    if refused:
        print(f"Windows startet ihn nicht im Hintergrund: {refused}.")
    else:
        print("    Ich frage ihn an wie Excel …")
        # Beim ersten Start schaut Windows sich Python oft erst genau an; das dauert.
        problem = dienst.answers(timeout=60)
        if problem:
            print(f"Der Prüfer im Hintergrund antwortet nicht: {problem}.")
            tail = dienst.log_tail(root)
            if tail:
                print("Das schreibt er ins Protokoll (data\\pruefer-dienst.log):")
                for line in tail:
                    print(f"    {line}")
            print("Schick davon ein Bildschirmfoto, dann weiß ich, was fehlt.")
            return 1
        print("    Er antwortet. Windows startet ihn bei jeder Anmeldung von selbst, also auch nach einem Neustart.")
    print()
    print(WINDOWS_DONE)
    if refused:
        print("Im Hintergrund starten ließ er sich nicht. Dann starte VoiceAgent.bat, solange du den Prüfer brauchst.")
    else:
        print("Dieses Fenster und den Browser brauchst du dafür nicht.")
    return 0


def main() -> None:
    parser = argparse.ArgumentParser(description="Das Excel-Add-in (Prüfer in Excel) einrichten")
    parser.add_argument("--entfernen", action="store_true", help="Add-in wieder aus Excel nehmen")
    args = parser.parse_args()
    raise SystemExit(install(remove=args.entfernen))


if __name__ == "__main__":
    main()
