import os

# Die Tests fragen nie das echte Excel auf dem Rechner (am Mac fragt sonst macOS um Erlaubnis).
os.environ.setdefault("VOICEAGENT_NO_EXCEL", "1")

import pytest


@pytest.fixture(autouse=True)
def _testserver(monkeypatch):
    """Der TestClient fragt unter dem Namen „testserver“ an; Gegenprobe selbst antwortet nur unter 127.0.0.1 und
    localhost (test_nur_lokal.py prüft das ohne diese Ausnahme)."""
    # Immer, nicht nur wenn schon geladen: Ein Test, der den Dienst erst selbst lädt (client_with), braucht es auch,
    # sonst hängt er davon ab, welche Tests vor ihm liefen.
    import voiceagent.app as app

    monkeypatch.setattr(app, "LOCAL_HOSTS", app.LOCAL_HOSTS | {"testserver"})
