"""Was Gegenprobe über das ganze Blatt selbst nachrechnet (task.py) und wie WorkbookTracker Änderungen liest."""

from test_excel_structure import WAHL, mappe

from voiceagent import task
from voiceagent.workbook import WorkbookTracker, read_book


def book(cells=None):
    return read_book(mappe(("Tabelle1", dict(cells or WAHL, C11="Kontrolle:"), "", {})))


def tracker_with(cells=None):
    tracker = WorkbookTracker()
    tracker.update("Wahl.xlsx", book(cells).sheets, 0.5)
    return tracker


def test_grid_and_facts_of_practice_sheet():
    found = task.material(tracker_with())
    assert "Zeile 6: B6 RN | C6 33,2 | D6 =0,332 * 577 → 191,564" in found.grid
    assert "Zeile 11: C11 Kontrolle: | D11 =SUMME(D6:D10) → 577" in found.grid
    texts = "\n".join(f.text for f in found.facts)
    # Was VoiceAgent selbst nachrechnet: Rundung, die feste 577, abgetippte Anteile.
    assert "einzeln auf ganze gerundet ergibt zusammen 578, nicht 577" in texts
    assert "Die Zahl 577 steht fest in 5 Formeln (D6, D7, D8, D9, D10) und auf dem Blatt in D2 „Sitze(577)“" in texts
    assert "D6: 0,332 = Wähleranteil(%) 33,2 / 100" in texts
    assert "Wähleranteil(%) (C6:C10): 5 Zahlen von 6,6 bis 33,2, zusammen 100." in texts
    assert found.headers == ["Parteien", "Wähleranteil(%)", "Anzahl Sitze"]


def test_sum_that_leaves_out_a_row_errors_and_a_typed_number():
    cells = dict(WAHL, D11=("=", "SUM(D6:D9)", 511.222), D10=65.778, D12=("=", "D6/0", "#DIV/0!"))
    facts = {f.key: f for f in task.facts(*tracker_with(cells).active_state()[1:])}
    assert facts["gap:D11"].ask == "Die Summe in D11 zählt Rest nicht mit. Ist das Absicht?"
    assert facts["typed:Anzahl Sitze"].text == "In Anzahl Sitze rechnen 4 Zeilen mit Formel, bei D10 (Rest) steht nur eine Zahl."
    assert facts["error:D12"].ask == "In D12 steht #DIV/0!. Weißt du schon, woran das liegt?"


# --- Kleine Fehler aus einer Testsitzung ---


def test_a_number_typed_and_deleted_again_is_no_decision():
    tracker = tracker_with()
    scratch = dict(WAHL, F14=578)
    assert tracker.update("Wahl.xlsx", book(scratch).sheets, 1.0) == []
    # Wieder gelöscht: Das war eine Nebenrechnung, keine Entscheidung.
    assert tracker.update("Wahl.xlsx", book().sheets, 2.0) == []
    # Ein Wert, der schon da war, zählt weiter.
    events = tracker.update("Wahl.xlsx", book(dict(WAHL, C6=34)).sheets, 3.0)
    assert [(e.label, e.before, e.after) for e in events] == [("Wähleranteil(%)", "33,2", "34")]
    assert events[0].text.endswith("(Zeile Parteien RN)")
    outside = tracker.update("Wahl.xlsx", book(dict(WAHL, C6=34, D2="Sitze(578)")).sheets, 4.0)
    assert outside[0].text.endswith("(Blatt Tabelle1)") and "Zeile Zeile" not in outside[0].text


def test_onedrive_book_only_in_the_cloud_or_locked(monkeypatch):
    from test_excel_live import Run

    from voiceagent import excel
    from voiceagent.excel import ExcelLink, Snapshot

    link = ExcelLink("darwin", Run())
    onedrive = Snapshot("Wahl.xlsx", "https://d.docs.live.net/4f2a9c/Dokumente/Wahl.xlsx")
    monkeypatch.setattr(excel, "local_path", lambda path, home=None: None)
    monkeypatch.setattr(excel, "cloud_denied", lambda home=None: False)
    assert link.book(onedrive) is None and link.file_note == "cloud"
    monkeypatch.setattr(excel, "cloud_denied", lambda home=None: True)
    assert link.book(onedrive) is None and link.file_note == "denied"

    def locked(path, home=None):
        raise PermissionError(path)

    monkeypatch.setattr(excel, "local_path", locked)
    assert link.book(Snapshot("Wahl.xlsx", "/Users/b/Library/CloudStorage/OneDrive/Wahl.xlsx")) is None
    assert link.file_note == "denied"


def test_inputs_above_their_calculations_are_no_missing_formulas():
    """Umsatz und Kosten getippt, Gewinn = Umsatz - Kosten darunter: Was die Formeln der Spalte einzeln lesen, ist eine
    Eingabe. Rechnen die Formeln jede Zeile anders, fällt eine feste Zahl dazwischen auch nicht auf."""
    from test_workbook import xlsx

    from voiceagent import pruefer

    cells = {"A1": "Posten", "B1": "Betrag", "A2": "Umsatz", "B2": 5000.0, "A3": "Kosten", "B3": 3000.0,
             "A4": "Gewinn", "B4": ("=", "B2-B3", 2000.0), "A5": "Steuer", "B5": ("=", "B4*0.2", 400.0),
             "A6": "Netto", "B6": ("=", "B4-B5", 1600.0)}
    done = pruefer.Checker(None).check_file("g.xlsx", xlsx(("Rechnung", cells)))
    assert [h["key"] for s in done["sheets"] for h in s["hints"]] == []
    costs = {"A1": "Posten", "B1": "Betrag", "A2": "Miete", "B2": 1200.0, "A3": "Strom", "B3": ("=", "45*12", 540.0),
             "A4": "Versicherung", "B4": 300.0, "A5": "Summe", "B5": ("=", "SUM(B2:B4)", 2040.0)}
    done = pruefer.Checker(None).check_file("k.xlsx", xlsx(("Kosten", costs)))
    assert not any(h["key"].startswith("typed:") for s in done["sheets"] for h in s["hints"])
