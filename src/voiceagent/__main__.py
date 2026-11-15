"""Startet Gegenprobe: `python -m voiceagent` oder `voiceagent`."""

from __future__ import annotations

import argparse
import json
import socket
import threading
import urllib.request
import webbrowser

import uvicorn

from . import addin, dienst
from .app import create_app
from .config import Settings


def _open_when_ready(url: str, attempts: int = 50) -> None:
    """Öffnet den Browser erst, wenn der Server antwortet."""
    for _ in range(attempts):
        try:
            with urllib.request.urlopen(f"{url}/api/status", timeout=1):
                webbrowser.open(url)
                return
        except OSError:
            threading.Event().wait(0.2)


def _ours(url: str) -> bool:
    """Antwortet dort schon eine Gegenprobe (etwa aus einem zweiten Doppelklick)?"""
    try:
        with urllib.request.urlopen(f"{url}/api/status", timeout=2) as response:
            return json.loads(response.read(1000)) == {"ok": True}
    except (OSError, ValueError):
        return False


def _free(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        try:
            probe.bind(("127.0.0.1", port))
        except OSError:
            return False
    return True


def choose_port(port: int, busy=dienst.busy, ours=_ours, free=_free) -> tuple[int | None, str]:
    """Welcher Port: der gewünschte; läuft dort schon Gegenprobe, keiner (None); hat ihn ein anderes Programm, der
    nächste freie (nie der Port des Seitenfensters). Dazu ein Satz fürs Fenster oder ""."""
    if not busy(port):
        return port, ""
    if ours(f"http://127.0.0.1:{port}"):
        return None, "Gegenprobe läuft schon, in einem anderen Fenster. Ich öffne sie dir im Browser."
    for other in range(port + 1, port + 40):
        if other != addin.PORT and free(other):
            return other, f"Port {port} hat schon ein anderes Programm. Gegenprobe nimmt diesmal Port {other}."
    return port, f"Port {port} hat schon ein anderes Programm, und ich finde keinen freien. Schließ andere Programme und starte noch einmal."


def main() -> None:
    parser = argparse.ArgumentParser(description="Gegenprobe lokal starten")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true", help="Browser nicht öffnen")
    args = parser.parse_args()

    port, said = choose_port(args.port)
    if said:
        print(said)
    if port is None:
        if not args.no_browser:
            webbrowser.open(f"http://127.0.0.1:{args.port}")
        return
    url = f"http://127.0.0.1:{port}"
    print(f"Gegenprobe läuft auf {url}")
    if not args.no_browser:
        threading.Thread(target=_open_when_ready, args=(url,), daemon=True).start()
    settings = Settings()
    app = create_app(settings)
    found = addin.ready()
    # Hält macOS den Port für den Prüfer im Hintergrund offen, nicht anklopfen: Das würde ihn unnötig starten.
    if found is not None and (dienst.waits() or dienst.busy(addin.PORT)):
        print("Der Prüfer für Excel läuft schon im Hintergrund: in Excel unter Start auf Prüfer klicken.")
    elif found is not None:
        # Eingerichtet mit `python -m voiceagent.addin`: dasselbe noch einmal mit https für das Seitenfenster in Excel.
        addin.serve(app, *found)
        print(f"Excel-Add-in bereit: in Excel unter Start auf Prüfer klicken ({addin.ORIGIN}).")
    # Nur auf 127.0.0.1: Die App ist von außen nicht erreichbar.
    uvicorn.run(app, host="127.0.0.1", port=port, log_level="warning")


if __name__ == "__main__":
    main()
