"""„Passt so“: Ein Hinweis, den jemand angeschaut und für richtig befunden hat, verschwindet für diese Mappe. Das
Fenster hält ihn in den Einstellungen der Mappe; ändert sich an der Stelle etwas, kommt er wieder."""

from test_addin import payload
from test_excel_live import FakeExcel, client_with

HEADERS = {"X-VoiceAgent": "1"}


def _check(client, body):
    return client.post("/api/addin/pruefen", json=body, headers=HEADERS).json()


def test_a_hint_marked_passt_so_is_gone_and_counted(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    first = _check(client, payload())
    division = next(h for h in first["hints"] if h["cell"] == "C8")
    assert len(division["id"]) == 12 and len({h["id"] for h in first["hints"]}) == len(first["hints"])
    # Dieselbe Mappe, derselbe Hinweis: dieselbe Kennung.
    assert [h["id"] for h in _check(client, payload())["hints"]] == [h["id"] for h in first["hints"]]
    then = _check(client, payload(passt={"Wahl": [division["id"]]}))
    assert "C8" not in [h["cell"] for h in then["hints"]] and then["passt"] == 1
    assert len(then["hints"]) == len(first["hints"]) - 1 and then["text"] != first["text"]
    # Eine Kennung, die zu nichts passt, und ein anderes Blatt ändern nichts.
    other = _check(client, payload(passt={"Wahl": ["000000000000"], "Leer": [division["id"]]}))
    assert len(other["hints"]) == len(first["hints"]) and "passt" not in other


def test_a_hint_comes_back_when_the_place_changes(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    first = _check(client, payload())
    division = next(h for h in first["hints"] if h["cell"] == "C8")
    body = payload(passt={"Wahl": [division["id"]]})
    body["sheets"][0]["formulas"][6][2] = "=C4/0"
    again = _check(client, body)
    assert "C8" in [h["cell"] for h in again["hints"]] and "passt" not in again


def test_with_everything_marked_nothing_further_is_said(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    first = _check(client, payload())
    done = _check(client, payload(passt={"Wahl": [h["id"] for h in first["hints"]]}))
    assert done["hints"] == [] and done["text"] == "Mir fällt nichts weiter auf." and done["passt"] == len(first["hints"])


def test_other_sheets_and_the_report_count_without_them(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    wahl = payload()["sheets"][0]
    leer = {"name": "Notiz", "row": 1, "col": 1, "values": [["Text"]]}
    body = {"book": "Wahl.xlsx", "active": "Notiz", "sheets": [leer, wahl]}
    before = {o["name"]: o for o in _check(client, body)["others"]}["Wahl"]
    hints = _check(client, {**body, "active": "Wahl"})["hints"]
    errors = [h["id"] for h in hints if h["level"] == "fehler"]
    after = {o["name"]: o for o in _check(client, {**body, "passt": {"Wahl": errors}})["others"]}["Wahl"]
    assert before["fehler"] == len(errors) and after["fehler"] == 0
    made = client.post("/api/addin/bericht", json={"daten": True}, headers=HEADERS).json()
    report = client.get(made["url"])
    assert report.status_code == 200 and f"{len(errors)} Hinweise mit „Passt so“ ausgeblendet." in report.text


def test_what_the_pane_sends_is_kept_small_and_clean(tmp_path):
    from voiceagent import addin

    assert addin.passt_from({"passt": "kaputt"}) == {} and addin.passt_from({}) == {}
    got = addin.passt_from({"passt": {"Wahl": ["a" * 40, 7, None, "b"], "Leer": "x"}})
    assert got == {"Wahl": frozenset({"a" * 16, "b"})}
