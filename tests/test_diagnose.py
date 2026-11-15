"""Zustand und Fehlerbericht: jeder Baustein in einem Satz, und nie ein Schlüssel oder eine IBAN im Bericht."""

import logging

from voiceagent import diagnose
from voiceagent.config import Settings
from voiceagent.update import last_run, record


def checks(state) -> dict:
    return {c["name"]: (c["ok"], c["text"]) for c in state["checks"]}


def test_every_part_says_how_it_is(tmp_path):
    state = diagnose.status(Settings(data_dir=tmp_path), root=tmp_path)
    found = checks(state)
    assert set(found) >= {"Version", "Daten"}
    assert "Spracherkennung" not in found and "Texterkennung" not in found


def test_report_keeps_warnings_but_nothing_private(tmp_path):
    diagnose.install()
    logging.getLogger("voiceagent.test").warning("Konto DE89 3704 0044 0532 0130 00 und gsk_%s", "a" * 30)
    text = diagnose.report(diagnose.status(Settings(data_dir=tmp_path), root=tmp_path), Settings(data_dir=tmp_path))
    assert "## Letzte Warnungen und Fehler" in text and "voiceagent.test" in text
    assert "DE89 3704" not in text and "gsk_" not in text


def test_scrub_finds_keys_without_being_told():
    text = diagnose.scrub(f"gsk_{'a' * 30} und sk-{'b' * 20} und AIza{'c' * 30} und ?key=geheimgeheim")
    assert text.count("[Schlüssel]") == 4


def test_last_update_is_remembered(tmp_path):
    assert last_run(tmp_path) == {}
    record("Neue Version geladen (3 Änderungen).", True, tmp_path)
    assert last_run(tmp_path)["message"] == "Neue Version geladen (3 Änderungen)." and last_run(tmp_path)["packages"]


def test_a_zip_install_shows_its_stand_without_branch(tmp_path):
    """Aus einer ZIP von GitHub: Stand und Commit aus stand.txt, ohne „Zweig None“."""
    (tmp_path / "src" / "voiceagent").mkdir(parents=True)
    (tmp_path / "src" / "voiceagent" / "stand.txt").write_text("9f062b75817394dd66163baf11c299896136a3c0 1794334526", encoding="utf-8")
    version = diagnose.status(Settings(data_dir=tmp_path), root=tmp_path)["checks"][0]
    assert version["ok"] is True and version["text"].startswith("Stand 10.11.2026 (9f062b7).")
