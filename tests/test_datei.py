"""Eine Excel-Datei ins Seitenfenster ziehen und prüfen lassen, ohne sie in Excel zu öffnen."""

from test_addin import payload
from test_excel_live import FakeExcel, client_with
from test_workbook import xlsx

HEADERS = {"X-VoiceAgent": "1"}


def befunde():
    return xlsx(("Patient Huber", {
        "A1": "Wert", "B1": "Menge", "C1": "Anteil",
        "A2": "Natrium", "B2": 140, "C2": ("=", "B2/B5", 0.56),
        "A3": "Kalium", "B3": 4, "C3": ("=", "B3/B5", 0.016),
        "A4": "Chlorid", "B4": 106, "C4": ("=", "B4/B6", "#DIV/0!"),
        "A5": "Summe", "B5": ("=", "SUM(B2:B4)", 250),
    }))


def test_a_dropped_file_is_checked_without_excel(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    assert client.post("/api/addin/bericht", json={"daten": True, "datei": True}, headers=HEADERS).json()["state"] == "again"
    found = client.post("/api/addin/datei?name=Befunde%20Huber.xlsx", content=befunde(), headers=HEADERS).json()
    assert found["state"] == "ok" and found["book"] == "Befunde Huber.xlsx"
    # C4 teilt durch die leere B6: ein Hinweis mit der Ursache (früher noch einer zur leeren Zelle).
    assert found["text"] == "1 Blatt geprüft: 1 Fehler."
    sheet = found["sheets"][0]
    assert sheet["name"] == "Patient Huber" and any(h["cell"] == "C4" and h["level"] == "fehler" for h in sheet["hints"])
    full = client.post("/api/addin/bericht", json={"daten": True, "datei": True}, headers=HEADERS).json()
    assert "Befunde Huber.xlsx" in client.get(full["url"]).text
    plain = client.get(client.post("/api/addin/bericht", json={"daten": False, "datei": True}, headers=HEADERS).json()["url"]).text
    assert "Huber" not in plain and "Natrium" not in plain and "C4" in plain
    # Der Bericht über die offene Mappe bleibt davon unberührt.
    assert client.post("/api/addin/bericht", json={"daten": True}, headers=HEADERS).json()["state"] == "again"
    bad = client.post("/api/addin/datei?name=x.xlsx", content=b"kein zip", headers=HEADERS).json()
    assert bad["state"] == "error"
    assert client.post("/api/addin/datei", content=befunde()).status_code == 403


def test_a_file_that_is_no_version_of_the_open_workbook(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    client.post("/api/addin/pruefen", json=payload(), headers=HEADERS)
    found = client.post("/api/addin/vergleich?name=Befunde.xlsx", content=befunde(), headers=HEADERS).json()
    assert found["state"] == "ok" and found["related"] is False
