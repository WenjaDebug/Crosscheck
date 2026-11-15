"""Lokaler Server: liefert die Startseite, den Prüfer, den Excel-Helfer und das Seitenfenster in Excel aus."""

from __future__ import annotations

import asyncio
import json
import hashlib
import logging
import os
import sys
from contextlib import asynccontextmanager
from datetime import datetime
from pathlib import Path
from urllib.parse import urlsplit

from fastapi import Body, FastAPI, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse, PlainTextResponse, Response
from fastapi.staticfiles import StaticFiles

from . import __version__, addin, bericht, diagnose, dienst, sprache, uebung
from .analyse import failure
from .config import Settings
from .excel import ExcelLink
from .helfer import Helper
from .pruefer import Checker, report_hint
from .sprache import T

WEB_DIR = Path(__file__).parent / "web"


class Fresh(StaticFiles):
    """Jedes Mal nachfragen, ob sich eine Datei geändert hat (meist kommt nur „unverändert“ zurück): Sonst zeigt Excel
    nach einem Update noch das alte Seitenfenster."""

    async def get_response(self, path: str, scope) -> Response:
        response = await super().get_response(path, scope)
        response.headers["Cache-Control"] = "no-cache"
        return response


def pane_version() -> str:
    """Woran das Seitenfenster merkt, dass im Hintergrund eine neuere Fassung von ihm läuft."""
    digest = hashlib.sha256()
    for path in [*sorted((WEB_DIR / "addin").iterdir()), WEB_DIR / "style.css"]:
        if path.is_file():
            digest.update(path.name.encode("utf-8") + b"\0" + path.read_bytes())
    return digest.hexdigest()[:12]


_log = logging.getLogger("voiceagent.app")
# Gegenprobe antwortet nur unter diesen Namen. Eine fremde Webseite, deren Name plötzlich auf diesen Rechner zeigt
# (DNS-Rebinding), kommt mit ihrem eigenen Namen an und wird abgewiesen, ebenso eine fremde Seite, die etwas schicken
# oder die Live-Verbindung öffnen will: Für WebSockets fragt der Browser nicht vorher nach (CORS).
LOCAL_HOSTS = frozenset({"127.0.0.1", "localhost", "::1"})
_READING = frozenset({"GET", "HEAD", "OPTIONS"})


def foreign_text() -> str:
    return T("Gegenprobe antwortet nur Seiten auf diesem Rechner. Öffne es über http://127.0.0.1:8765 oder in Excel "
             "über Start → Prüfer.", "Gegenprobe only answers pages on this computer. Open it at http://127.0.0.1:8765 "
             "or in Excel under Home → Checker.")


def unexpected() -> str:
    return T("In Gegenprobe ist dabei ein Fehler passiert, deine Mappe ist unverändert. Versuch es noch einmal. ",
             "Something went wrong in Gegenprobe; your workbook is unchanged. Try again. ") + report_hint()


def _language(headers) -> tuple[str, str]:
    """Die Sprache der Hinweise und der Formeln, wie das Seitenfenster sie in der Anfrage nennt (sprache.py)."""
    return headers.get("x-gegenprobe-sprache", ""), headers.get("x-gegenprobe-formeln", "")


def create_app(settings: Settings | None = None, excel: ExcelLink | None = None) -> FastAPI:
    settings = settings or Settings()
    if excel is None and os.environ.get("VOICEAGENT_NO_EXCEL", "").strip() not in ("1", "true"):
        # Das laufende Excel direkt lesen (Mac und Windows), für den Helfer und den Prüfer im Browser.
        excel = ExcelLink()
    excel = excel if excel is not None and excel.supported else None

    @asynccontextmanager
    async def lifespan(_app: FastAPI):
        # Erst jetzt: uvicorn richtet seine Logger beim Start neu ein.
        diagnose.install()
        yield

    app = FastAPI(title="Gegenprobe", lifespan=lifespan)
    app.add_middleware(LocalOnly)

    @app.exception_handler(Exception)
    async def unexpected_error(request: Request, err: Exception) -> JSONResponse:
        """Ein unerwarteter Fehler: Das Fenster bekommt einen Satz, was los ist und was man tun kann, statt eines
        Stacktraces. Der Fehlerbericht bekommt, wo es war, ohne die Meldung (darin könnten Werte der Mappe stehen).
        Hier gilt die Sprache der Anfrage nicht mehr (LocalOnly ist schon verlassen): Sie steht noch in den Headern."""
        _log.error("Fehler bei %s: %s", request.url.path, failure(err))
        with sprache.using(*_language(request.headers)):
            text = unexpected()
        return JSONResponse({"state": "error", "text": text}, status_code=500)

    app.state.excel = excel

    @app.get("/")
    def index() -> FileResponse:
        return FileResponse(WEB_DIR / "index.html")

    @app.get("/api/status")
    def status() -> dict:
        return {"ok": True}

    def current_state() -> dict:
        return diagnose.status(settings, excel=excel)

    checker = Checker(excel, settings.data_dir)
    bericht.tidy(settings.data_dir / "berichte")
    pane = pane_version()
    helper = Helper(excel)

    @app.get("/api/pruefen")
    async def check_excel() -> dict:
        """Excel-Prüfer: die offene Mappe nachrechnen, ohne KI."""
        return await asyncio.to_thread(checker.check)

    @app.post("/api/addin/pruefen")
    async def check_addin(request: Request, body: dict = Body(...)) -> dict:
        """Das Seitenfenster in Excel schickt die Mappe (Office.js): nachrechnen wie im Prüfer."""
        require_local_client(request)
        result = await asyncio.to_thread(checker.check_payload, body)
        return {**result, "version": pane}

    @app.post("/api/addin/andere")
    async def others_addin(request: Request) -> dict:
        """Wie weit die übrigen Blätter geprüft sind (große prüft der Prüfer im Hintergrund), ohne neu zu rechnen."""
        require_local_client(request)
        return await asyncio.to_thread(checker.others)

    @app.post("/api/addin/alle")
    async def fix_all(request: Request, body: dict = Body(...)) -> dict:
        """„Alle einsetzen“: die Korrekturen des Blatts zusammen, nachgerechnet, wie sie nach dem Einsetzen rechnen."""
        require_local_client(request)
        return await asyncio.to_thread(checker.joint, str(body.get("sheet") or ""))

    @app.post("/api/addin/uebung")
    async def practice_addin(request: Request) -> dict:
        """Die Übungsmappe in Excel öffnen: Das Seitenfenster kann sie nicht herunterladen."""
        require_local_client(request)
        return await asyncio.to_thread(uebung.open_copy, settings.data_dir / "uebungsmappen")

    @app.get("/api/start")
    async def start_state() -> dict:
        """Für die Startseite: ob das Excel-Fenster eingerichtet ist und der Prüfer im Hintergrund läuft, auch nach
        dem letzten Neustart. Nur im eigenen Ordner nachsehen; den Ordner von Excel zu lesen, ließe macOS jedes Mal
        nach Erlaubnis fragen."""
        found = {"mac": sys.platform == "darwin", "addin": (addin.folder() / "ca.pem").exists(), "dienst": dienst.installed()}
        if sys.platform == "win32":
            from . import windows

            # Unter Windows startet er bei der Anmeldung (windows.py); läuft er, steht er unter den Prozessen.
            found |= {"windows": True, "dienst": windows.installed()}
            if found["dienst"]:
                ready = bool(await asyncio.to_thread(windows.processes))
                why = "Windows hat ihn nicht gestartet, oder er wurde beendet"
                found |= {"laeuft": ready, "abruf": False, "grund": "" if ready else why, "neustart": None}
            return found
        if found["dienst"]:
            info = await asyncio.to_thread(dienst.state)
            ready = dienst.ready(info)
            # Startet er erst, wenn Excel ihn braucht, sagt ein Start nach dem Neustart nichts darüber, ob er geht.
            boot = None if info.get("sockets") else await asyncio.to_thread(dienst.boot_time)
            found |= {"laeuft": ready, "abruf": bool(info.get("sockets")), "grund": "" if ready else dienst.explain(info),
                      "neustart": dienst.since_boot(None, boot)}
        return found

    @app.post("/api/addin/bericht")
    async def report_addin(request: Request, body: dict = Body(...)) -> dict:
        """Prüfbericht über die ganze Mappe als Seite im Datenordner; das Seitenfenster öffnet sie im Browser."""
        require_local_client(request)
        with_data = bool(body.get("daten"))
        html = await asyncio.to_thread(checker.report, with_data, bool(body.get("datei")))
        if html is None:
            return {"state": "again"}
        name = bericht.save(settings.data_dir / "berichte", html, with_data)
        return {"state": "ok", "url": f"/berichte/{name}"}

    @app.get("/berichte/{name}")
    async def report_file(name: str) -> FileResponse:
        path = settings.data_dir / "berichte" / name
        if not bericht.NAME.fullmatch(name) or not path.is_file():
            raise HTTPException(404, T("Diesen Bericht gibt es nicht (mehr)", "This report does not exist (any more)"))
        return FileResponse(path, media_type="text/html; charset=utf-8")

    @app.post("/api/addin/vergleich")
    async def compare_addin(request: Request, name: str = "") -> dict:
        """Die offene Mappe mit einer früheren Version (.xlsx im Rumpf) vergleichen. Die Datei bleibt im Arbeitsspeicher."""
        require_local_client(request)
        data = await request.body()
        return await asyncio.to_thread(checker.compare, name.strip()[:200] or T("Datei", "File"), data)

    @app.post("/api/addin/datei")
    async def check_file_addin(request: Request, name: str = "") -> dict:
        """Eine .xlsx-Datei (im Rumpf) prüfen, ohne sie in Excel zu öffnen. Sie bleibt im Arbeitsspeicher."""
        require_local_client(request)
        data = await request.body()
        return await asyncio.to_thread(checker.check_file, name.strip()[:200] or T("Datei", "File"), data)

    @app.post("/api/addin/mappe")
    async def describe_addin(request: Request) -> dict:
        """Mappe in Worten: was die ganze zuletzt geprüfte Mappe rechnet, von den Daten bis zu den Ergebnissen."""
        require_local_client(request)
        return await asyncio.to_thread(checker.describe_book)

    @app.post("/api/addin/regeln")
    async def rules_addin(request: Request) -> dict:
        """Die eigenen Regeln je Spalte (spaltenregeln.py), mit ihrer Beschreibung."""
        require_local_client(request)
        return await asyncio.to_thread(checker.rule_list)

    @app.post("/api/addin/regeln/spalte")
    async def rule_column_addin(request: Request, body: dict = Body(...)) -> dict:
        """Für das Formular: die Spalte der markierten Zelle, was darin steht, und ihre Regel, wenn es eine gibt."""
        require_local_client(request)
        only = body.get("blatt")
        return await asyncio.to_thread(checker.rule_column, str(body.get("sheet") or ""), str(body.get("cell") or ""),
                                       str(body.get("spalte") or ""), None if only is None else str(only))

    @app.post("/api/addin/regeln/speichern")
    async def rule_save_addin(request: Request, body: dict = Body(...)) -> dict:
        require_local_client(request)
        rule, before = body.get("regel"), body.get("vorher")
        return await asyncio.to_thread(checker.rule_save, rule if isinstance(rule, dict) else {},
                                       before if isinstance(before, dict) else None)

    @app.post("/api/addin/regeln/loeschen")
    async def rule_remove_addin(request: Request, body: dict = Body(...)) -> dict:
        require_local_client(request)
        return await asyncio.to_thread(checker.rule_remove, str(body.get("spalte") or ""), str(body.get("blatt") or ""))

    @app.post("/api/addin/herleitung")
    async def explain_addin(request: Request, body: dict = Body(...)) -> dict:
        """Woraus die markierte Zelle entsteht und wohin sie fließt, in der zuletzt geprüften Mappe."""
        require_local_client(request)
        return await asyncio.to_thread(checker.explain, str(body.get("sheet") or ""), str(body.get("cell") or ""))

    @app.get("/addin/pruefer.html")
    async def addin_pane() -> Response:
        """Das Seitenfenster, Skript und Stil mit ihrer Fassung in der Adresse: Eine neue Fassung hat neue Adressen,
        die Excel nicht aus seinem Zwischenspeicher nehmen kann."""
        page = (WEB_DIR / "addin" / "pruefer.html").read_text(encoding="utf-8")
        for asset in ("/style.css", "/addin/addin.css", "/addin/sprache.js", "/addin/pruefer.js"):
            page = page.replace(f'"{asset}"', f'"{asset}?v={pane}"')
        page = page.replace("__VERSION__", f"Version {__version__}")
        return Response(page, media_type="text/html; charset=utf-8", headers={"Cache-Control": "no-cache"})

    @app.get("/addin/manifest.xml")
    async def addin_manifest() -> Response:
        """Die Beschreibung des Add-ins für Excel (die Einrichtung legt sie selbst in den Excel-Ordner)."""
        return Response(addin.manifest(), media_type="application/xml")

    @app.post("/api/helfer/vorschau")
    async def helper_preview(request: Request, body: dict = Body(...)) -> dict:
        """Excel-Helfer: den Auftrag verstehen und zeigen, was herauskommt. Geschrieben wird noch nichts."""
        require_local_client(request)
        return await asyncio.to_thread(helper.preview, str(body.get("text") or "")[:500])

    @app.post("/api/helfer/ausfuehren")
    async def helper_run(request: Request, body: dict = Body(...)) -> dict:
        """Die Vorschau in ein neues Blatt schreiben und danach jede Zahl nachprüfen."""
        require_local_client(request)
        return await asyncio.to_thread(helper.run, str(body.get("token") or ""), bool(body.get("pivot")))

    @app.post("/api/helfer/datei")
    async def helper_file(request: Request, name: str = "") -> dict:
        """Eine Datei mit Tabellen (CSV, Excel, Word, PDF): nur im Arbeitsspeicher lesen und eine Vorschau zeigen."""
        require_local_client(request)
        data = await request.body()
        return await asyncio.to_thread(helper.load, name.strip() or T("Datei", "File"), data)

    @app.post("/api/helfer/tabelle")
    async def helper_table(request: Request, body: dict = Body(...)) -> dict:
        """Eine andere Tabelle aus derselben Datei zeigen."""
        require_local_client(request)
        try:
            index = int(body.get("index") or 0)
        except (TypeError, ValueError):
            index = 0
        return await asyncio.to_thread(helper.table, str(body.get("file") or ""), index)

    @app.get("/api/zustand")
    async def health() -> dict:
        """Jeder Baustein mit einem Satz: läuft, fehlt, oder geht auch ohne."""
        return await asyncio.to_thread(current_state)

    @app.get("/api/zustand/bericht")
    async def health_report() -> PlainTextResponse:
        """Der Fehlerbericht zum Weitergeben: ohne Zellinhalte."""
        state = await asyncio.to_thread(current_state)
        text = diagnose.report(state, settings)
        name = f"voiceagent-fehlerbericht-{datetime.now().strftime('%Y-%m-%d')}.txt"
        headers = {"Content-Disposition": f'attachment; filename="{name}"'}
        return PlainTextResponse(text, headers=headers)

    app.mount("/", Fresh(directory=WEB_DIR), name="web")
    return app


def _host_name(value: str) -> str:
    """„localhost:8766“ → localhost, „[::1]:8765“ → ::1."""
    value = value.strip().lower()
    if value.startswith("["):
        return value[1:value.find("]")] if "]" in value else value
    return value.rsplit(":", 1)[0] if value.count(":") == 1 else value


def local_origin(origin: str) -> bool:
    """Eine Seite dieses Rechners."""
    try:
        parts = urlsplit(origin)
        return parts.scheme in ("http", "https") and (parts.hostname or "") in LOCAL_HOSTS
    except ValueError:
        return False


class LocalOnly:
    """Vor jeder Anfrage: unter einem Namen dieses Rechners, und was schreibt, nur von einer Seite dieses Rechners.
    Lesen darf jede Seite anfragen; die Antwort gibt der Browser einer fremden Seite nicht heraus."""

    def __init__(self, app) -> None:
        self.app = app

    async def __call__(self, scope, receive, send) -> None:
        if scope["type"] == "http":
            headers = {k.decode("latin-1").lower(): v.decode("latin-1") for k, v in scope.get("headers") or ()}
            origin = headers.get("origin")
            writes = scope.get("method", "GET").upper() not in _READING
            # „null“ schicken eingebettete Seiten. Dann genügt der eigene Header (require_local_client), den eine
            # fremde Seite nicht ohne Nachfrage setzen kann.
            foreign = origin is not None and origin != "null" and not local_origin(origin)
            if _host_name(headers.get("host", "")) not in LOCAL_HOSTS or writes and foreign:
                _log.warning("Fremde Anfrage abgewiesen: %s", scope.get("path", ""))
                with sprache.using(*_language(headers)):
                    said = foreign_text()
                await JSONResponse({"detail": said}, status_code=403)(scope, receive, send)
                return
            # Die Sprache, in der das Seitenfenster die Hinweise will (sprache.py), für die ganze Anfrage.
            with sprache.using(*_language(headers)):
                if sprache.key() == ("de", "de"):
                    await self.app(scope, receive, send)
                else:
                    await self.app(scope, receive, _speaking(send))
            return
        await self.app(scope, receive, send)


def _speaking(send):
    """Eine JSON-Antwort mit Zahlen, Daten und Formeln in der Sprache der Anfrage (sprache.output)."""
    start: dict = {}
    chunks: list[bytes] = []

    async def wrapped(message) -> None:
        nonlocal start
        if message["type"] == "http.response.start":
            kind = dict(message.get("headers") or ()).get(b"content-type", b"")
            if not kind.startswith(b"application/json"):
                await send(message)
                start = {"pass": True}
                return
            start = message
            return
        if message["type"] != "http.response.body" or start.get("pass"):
            await send(message)
            return
        chunks.append(message.get("body", b""))
        if message.get("more_body"):
            return
        body = b"".join(chunks)
        try:
            body = json.dumps(sprache.output(json.loads(body)), ensure_ascii=False).encode("utf-8")
        except ValueError:
            pass
        headers = [(k, v) for k, v in start.get("headers") or () if k.lower() != b"content-length"]
        await send({**start, "headers": [*headers, (b"content-length", str(len(body)).encode())]})
        await send({"type": "http.response.body", "body": body})

    return wrapped


def require_local_client(request: Request) -> None:
    """Nur die eigene Oberfläche und das Seitenfenster in Excel dürfen Daten schicken.

    Eine fremde Webseite kann den eigenen Header nicht setzen, ohne dass der Browser
    vorher nachfragt (CORS), und diese Nachfrage beantwortet der Server nicht."""
    if request.headers.get("x-voiceagent") != "1":
        raise HTTPException(403, T("Nur von Gegenprobe selbst", "Only from Gegenprobe itself"))
