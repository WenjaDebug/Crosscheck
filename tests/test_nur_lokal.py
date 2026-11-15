"""Gegenprobe antwortet nur Seiten auf diesem Rechner. Eine fremde Webseite darf nichts schicken, und ein fremder
Name, der auf einmal auf 127.0.0.1 zeigt (DNS-Rebinding), bekommt gar keine Antwort mit Daten."""

import pytest
from fastapi.testclient import TestClient
from test_excel_live import FakeExcel, client_with

from voiceagent import app as server


@pytest.fixture()
def api(tmp_path, monkeypatch):
    monkeypatch.setattr(server, "LOCAL_HOSTS", frozenset({"127.0.0.1", "localhost", "::1"}))
    app, _ = client_with(tmp_path, FakeExcel(None))
    return lambda base="http://127.0.0.1:8765": TestClient(app, base_url=base)


def test_pages_of_this_computer_work(api):
    for base in ("http://127.0.0.1:8765", "https://localhost:8766", "http://[::1]:8765"):
        assert api(base).get("/api/status").status_code == 200, base
    assert api().get("/api/zustand").status_code == 200


def test_a_foreign_name_gets_no_data(api):
    response = api("http://boese.example:8765").get("/api/zustand")
    assert response.status_code == 403 and "nur Seiten auf diesem Rechner" in response.json()["detail"]


def test_a_foreign_page_cannot_send(api):
    for origin in ("https://boese.example", "http://127.0.0.1.boese.example"):
        sent = api().post("/api/addin/pruefen", json={}, headers={"X-VoiceAgent": "1", "Origin": origin})
        assert sent.status_code == 403, origin
    # Lesen darf jede Seite anfragen; die Antwort gibt der Browser ihr nicht heraus (kein CORS).
    assert api().get("/api/status", headers={"Origin": "https://boese.example"}).status_code == 200


def test_saving_a_rule_needs_the_own_header(api):
    assert api().post("/api/addin/regeln/speichern", json={}).status_code == 403
