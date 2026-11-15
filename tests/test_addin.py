"""Das Excel-Add-in: Zellen aus Office.js, Prüfung, Manifest, Zertifikat und https."""

import shutil
import socket
import ssl
import subprocess
import sys
import time
import urllib.request
from xml.etree import ElementTree

import pytest
from test_excel_live import FakeExcel, client_with

from voiceagent import addin
from voiceagent.pruefer import _cell_of


def payload(**extra):
    """Wie pruefer.js die Mappe schickt: je Blatt Startzelle, Werte, Formeln (nur mit „=“) und Zahlenart."""
    values = [
        ["Partei", "Anteil", "Sitze", "Stand"],
        ["RN", 0.332, 191.564, 45566],
        ["NFP", 0.28, 161.56, 45566],
        ["Ensemble", 0.208, 120.016, 45566],
        ["Rest", 0.18, 103.86, None],
        ["Summe", 1, 473.14, None],
        ["", "", "#DIV/0!", True],
    ]
    formulas = [
        [None, None, None, None],
        [None, None, "=B3*577", None],
        [None, None, "=B4*577", None],
        [None, None, "=B5*577", None],
        [None, None, "=B6*577", None],
        [None, "=SUM(B3:B6)", "=SUM(C3:C5)", None],
        [None, None, "=C3/0", None],
    ]
    kinds = [[None] * 4] + [[None, "p", None, "d"] for _ in range(4)] + [[None, "p", None, None], [None] * 4]
    sheet = {"name": "Wahl", "row": 2, "col": 1, "values": values, "formulas": formulas, "kinds": kinds}
    return {"book": "Wahl.xlsx", "active": "Wahl", "sheets": [sheet, {"name": "Leer", "row": 1, "col": 1, "values": []}], **extra}


def test_cells_from_office_js_look_like_from_the_file():
    book, active, sheets = addin.sheets_from(payload())
    assert (book, active, [s.name for s in sheets]) == ("Wahl.xlsx", "Wahl", ["Wahl", "Leer"])
    cells = sheets[0].cells
    assert cells[(2, 1)].value == "Partei" and cells[(2, 1)].formula is None
    assert cells[(3, 2)].value == "33,2 %" and cells[(3, 4)].value == "01.10.2024"
    assert cells[(3, 3)].value == "191,564" and cells[(3, 3)].formula == "B3*577"
    assert cells[(8, 3)].value == "#DIV/0!" and cells[(8, 4)].value == "WAHR"
    assert (6, 4) not in cells and (8, 1) not in cells
    assert sheets[1].cells == {}


def test_addin_check_finds_mistakes_and_names_the_cell(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    result = client.post("/api/addin/pruefen", json=payload(), headers={"X-VoiceAgent": "1"}).json()
    assert result["state"] == "ok" and result["book"] == "Wahl.xlsx" and result["sheet"] == "Wahl"
    assert result["sheets"] == ["Wahl", "Leer"]
    by_cell = {h["cell"]: h for h in result["hints"]}
    assert "C8" in by_cell and "#DIV/0!" in by_cell["C8"]["text"]
    assert any("C7" == h["cell"] and "Rest" in h["text"] for h in result["hints"])
    assert all(h["cell"] for h in result["hints"] if h["level"] == "fehler")
    # Ohne den eigenen Header darf keine fremde Seite Zellen schicken.
    assert client.post("/api/addin/pruefen", json=payload()).status_code == 403


def test_addin_check_of_an_empty_sheet(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    result = client.post("/api/addin/pruefen", json=payload(active="Leer"), headers={"X-VoiceAgent": "1"}).json()
    assert result["state"] == "empty" and result["sheet"] == "Leer"
    empty = {"book": "Neu.xlsx", "active": "Tabelle1", "sheets": [{"name": "Tabelle1", "values": []}]}
    result = client.post("/api/addin/pruefen", json=empty, headers={"X-VoiceAgent": "1"}).json()
    assert result["state"] == "empty"


def test_addin_checks_the_whole_book(tmp_path):
    """SVERWEIS auf ein anderes Blatt wird nachgerechnet, und die übrigen Blätter werden gezählt."""
    _, client = client_with(tmp_path, FakeExcel(None))
    tarife = {"name": "Tarife", "row": 1, "col": 1, "values": [["Code", "Satz"], ["K", 0.1], ["A", 0.2], ["M", 0.3]]}
    rechnung = {"name": "Rechnung", "row": 1, "col": 1,
                "values": [["Code", "Betrag", "Satz"], ["K", 100, 0.1], ["A", 100, 0.2], ["M", 100, 0.2], ["X", 100, 0]],
                "formulas": [[None] * 3] + [[None, None, f"=IFERROR(VLOOKUP(A{r},Tarife!$A$2:$B$4,2,FALSE),0)"] for r in range(2, 6)]}
    wahl = payload()["sheets"][0]
    body = {"book": "Mappe.xlsx", "active": "Rechnung", "sheets": [tarife, rechnung, wahl]}
    result = client.post("/api/addin/pruefen", json=body, headers={"X-VoiceAgent": "1"}).json()
    assert result["stats"]["nicht_nachrechenbar"] == 0
    texts = " ".join(h["text"] for h in result["hints"])
    assert [h["cell"] for h in result["hints"] if h["key"].startswith("recalc:")] == ["C4"]  # Excel zeigt 0,2, richtig ist 0,3
    assert "„X“" in texts and "WENNFEHLER" in texts
    others = {o["name"]: o for o in result["others"]}
    assert others["Wahl"]["fehler"] >= 1 and "Tarife" not in others or others["Tarife"]["fehler"] == 0
    # Nur das aktive Blatt hat sich geändert: Die übrigen kommen aus dem Gedächtnis.
    rechnung["values"][3][2] = 0.3
    again = client.post("/api/addin/pruefen", json=body, headers={"X-VoiceAgent": "1"}).json()
    assert not [h for h in again["hints"] if h["key"].startswith("recalc:")] and again["others"] == result["others"]


@pytest.mark.parametrize("key, text, cell", [
    ("pattern:D7", "D7 (Sitze) rechnet anders", "D7"),
    ("gap:Sitze", "Die Summe in D11 lässt D9 aus", "D11"),
    ("constant:577", "Die feste Zahl 577 steht in 4 Formeln", ""),
    ("dupe:rn", "„RN“ steht 2-mal", ""),
    ("round:Sitze", "Sitze (C3:C7): 5 Zahlen", "C3:C7"),
])
def test_hint_cell(key, text, cell):
    assert _cell_of(key, text) == cell


def test_manifest_is_valid_xml_for_excel():
    root = ElementTree.fromstring(addin.manifest().encode("utf-8"))
    ns = {"o": "http://schemas.microsoft.com/office/appforoffice/1.1", "bt": "http://schemas.microsoft.com/office/officeappbasictypes/1.0",
          "v": "http://schemas.microsoft.com/office/taskpaneappversionoverrides"}
    assert root.find("o:Id", ns).text == addin.ADDIN_ID
    assert root.find("o:DefaultSettings/o:SourceLocation", ns).get("DefaultValue") == "https://localhost:8766/addin/pruefer.html"
    assert root.find("o:Hosts/o:Host", ns).get("Name") == "Workbook"
    # Jede resid hat eine Ressource, kurze Texte höchstens 125 Zeichen, lange 250.
    ids = {e.get("id") for e in root.iter() if e.tag.endswith(("}Image", "}Url", "}String")) and e.get("id")}
    used = {e.get("resid") for e in root.iter() if e.get("resid")}
    assert used <= ids
    for e in root.find("v:VersionOverrides/v:Resources/bt:ShortStrings", ns):
        assert len(e.get("DefaultValue")) <= 125
    for e in root.find("v:VersionOverrides/v:Resources/bt:LongStrings", ns):
        assert len(e.get("DefaultValue")) <= 250
    assert len(root.find("o:Description", ns).get("DefaultValue")) <= 250


def test_addin_pages_and_manifest_are_served(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    for path in ("/addin/pruefer.html", "/addin/pruefer.js", "/addin/addin.css", "/addin/befehle.html", "/addin/icon-80.png"):
        assert client.get(path).status_code == 200, path
    assert client.get("/addin/manifest.xml").text.startswith("<?xml")


def test_excel_never_keeps_an_old_pane(tmp_path):
    """Nach einem Update zeigte Excel noch das alte Seitenfenster aus seinem Zwischenspeicher."""
    from voiceagent.app import pane_version

    _, client = client_with(tmp_path, FakeExcel(None))
    script = client.get("/addin/pruefer.js")
    assert script.headers["cache-control"] == "no-cache"
    # Nachfragen kostet wenig: Ist die Datei gleich, kommt nur „unverändert“ zurück.
    assert client.get("/addin/pruefer.js", headers={"if-none-match": script.headers["etag"]}).status_code == 304
    # Läuft im Hintergrund eine neue Fassung, lädt sich das offene Fenster neu.
    result = client.post("/api/addin/pruefen", json={"book": "A.xlsx", "active": "T", "sheets": []},
                         headers={"X-VoiceAgent": "1"}).json()
    assert result["version"] == pane_version() and len(result["version"]) == 12
    assert "location.reload()" in script.text and "data.version !== version" in script.text
    # Skript und Stil haben ihre Fassung in der Adresse: Eine neue Fassung nimmt Excel nie aus dem Zwischenspeicher.
    page = client.get("/addin/pruefer.html")
    assert page.headers["cache-control"] == "no-cache"
    for asset in ("/style.css", "/addin/addin.css", "/addin/pruefer.js"):
        assert f'"{asset}?v={pane_version()}"' in page.text, asset
        assert client.get(f"{asset}?v={pane_version()}").status_code == 200


def test_not_set_up_means_no_https(tmp_path, monkeypatch):
    monkeypatch.setenv("VOICEAGENT_ADDIN_DIR", str(tmp_path / "addin"))
    assert addin.ready() is None


def test_setup_outside_the_mac_says_so(monkeypatch, capsys):
    # Windows hat seinen eigenen Weg (test_windows_fenster.py); sonst gibt es das Seitenfenster nicht.
    monkeypatch.setattr(addin.sys, "platform", "linux")
    assert addin.install() == 1
    assert "nur am Mac und unter Windows" in capsys.readouterr().out


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


@pytest.mark.skipif(shutil.which("openssl") is None or sys.platform == "win32", reason="der Weg mit openssl ist der vom Mac")
def test_certificate_and_https_server(tmp_path, monkeypatch):
    monkeypatch.setenv("VOICEAGENT_ADDIN_DIR", str(tmp_path / "addin"))
    ca, cert, key = addin.certificates()
    first = cert.read_bytes()
    done = subprocess.run(["openssl", "verify", "-CAfile", str(ca), str(cert)], capture_output=True, text=True, encoding="utf-8")
    assert done.returncode == 0, done.stderr
    # Noch lange gültig: beim nächsten Start bleibt es, und die Stelle (der der Mac vertraut) sowieso.
    assert addin.ready() == (cert, key) and cert.read_bytes() == first
    assert oct(key.stat().st_mode & 0o777) == "0o600"

    app, _ = client_with(tmp_path, FakeExcel(None))
    port = _free_port()
    addin.serve(app, cert, key, port)
    context = ssl.create_default_context(cafile=str(ca))
    for _ in range(100):
        try:
            with urllib.request.urlopen(f"https://localhost:{port}/addin/pruefer.html", context=context, timeout=2) as response:
                assert response.status == 200 and b"Gegenprobe Pr" in response.read()
            break
        except OSError:
            time.sleep(0.1)
    else:
        pytest.fail("https antwortet nicht")


@pytest.mark.skipif(shutil.which("openssl") is None or sys.platform == "win32", reason="der Weg mit openssl ist der vom Mac")
def test_a_certificate_without_its_authority_is_renewed(tmp_path, monkeypatch):
    """Das openssl am Mac (LibreSSL) lässt den Authority Key Identifier weg; Python 3.13 lehnt das ab."""
    monkeypatch.setattr(addin, "_LEAF_EXTENSIONS", addin._LEAF_EXTENSIONS.replace("keyid, issuer", "none"))
    ca, cert, key = addin.certificates(tmp_path)
    assert not addin._linked(cert)
    monkeypatch.undo()
    authority = ca.read_bytes()
    assert addin.certificates(tmp_path) == (ca, cert, key) and addin._linked(cert)
    # Die Stelle, der der Mac vertraut, bleibt: Es fragt nicht noch einmal nach dem Passwort.
    assert ca.read_bytes() == authority
    context = ssl.create_default_context(cafile=str(ca))
    # So streng wie Python ab 3.13 von sich aus.
    context.verify_flags |= ssl.VERIFY_X509_STRICT
    app, _ = client_with(tmp_path, FakeExcel(None))
    port = _free_port()
    addin.serve(app, cert, key, port)
    for _ in range(100):
        try:
            with urllib.request.urlopen(f"https://localhost:{port}/addin/pruefer.html", context=context, timeout=2) as response:
                assert response.status == 200
            break
        except OSError:
            time.sleep(0.1)
    else:
        pytest.fail("https antwortet nicht")


def test_hours_with_a_time_format_are_still_hours():
    """8 Stunden in einer Zelle mit dem Format hh:mm (Excel zeigt 00:00): Gerechnet wird mit 8, die Summe darüber ist
    richtig, und die Zelle bekommt einen Hinweis mit dem Format der Nachbarn zum Einsetzen."""
    from voiceagent import pruefer

    shifts = {"Früh": (0.25, 14.5 / 24), "Spät": (14 / 24, 22.5 / 24), "Nacht": (22 / 24, 6.5 / 24)}
    values = [["Datum", "Schicht", "Beginn", "Ende", "Pause (h)", "Stunden", "", "Stunden gesamt", 240]]
    formulas, kinds = [[0, 8, "=SUM(F2:F31)"]], [[9, 5, "d"]]
    for i in range(30):
        name = list(shifts)[i % 3]
        values.append([46174 + i // 3, name, *shifts[name], 0.5, 8, "", "", ""])
        formulas.append([i + 1, 5, f"=MOD(D{i + 2}-C{i + 2},1)*24-E{i + 2}"])
        kinds += [[i + 1, 0, "d"], [i + 1, 2, "d"], [i + 1, 3, "d"]]
    book, active, sheets = addin.sheets_from({"book": "S.xlsx", "active": "S", "sheets": [
        {"name": "S", "row": 1, "col": 1, "values": values, "formulas_at": formulas, "kinds_at": kinds}]})
    hints = {h["key"]: h for h in pruefer.Checker(None).check_book(book, active, sheets)["hints"]}
    assert not any(k.startswith(("label:", "range:", "value:")) for k in hints)
    hint = hints["clipped:F10"]
    assert hint["text"].startswith("F10 (Stunden") and "enthält 8, ist aber als Uhrzeit oder Datum formatiert" in hint["text"]
    assert hint["fix"]["format"] == {"F10": "F9"} and hint["fix"]["cells"] == {"F10": "=MOD(D10-C10,1)*24-E10"}
