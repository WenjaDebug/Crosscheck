"""Listen mit Köpfen auf Französisch, Italienisch oder Englisch: „Date de sortie“ vor „Date d'entrée“, „Somma“ unter
der Spalte und „Part femmes“ über 100 % sind derselbe Fehler wie auf Deutsch. Mehrdeutige Wörter („Part no“,
„Share price“, „Media“) machen keinen falschen Alarm."""

import pytest
from test_auswertung import faelle, kinds, report, texts

from voiceagent.xlsx import Cell

WORDS = {
    "fr": {"Fall-Nr": "N° de cas", "Geburtsdatum": "Date de naissance", "Eintritt": "Date d'entrée",
           "Austritt": "Date de sortie", "Alter": "Âge", "Gewicht kg": "Poids kg", "Größe": "Taille cm",
           "Liegedauer": "Durée de séjour", "Station": "Unité"},
    "it": {"Fall-Nr": "N. caso", "Geburtsdatum": "Data di nascita", "Eintritt": "Data di ammissione",
           "Austritt": "Data di dimissione", "Alter": "Età", "Gewicht kg": "Peso kg", "Größe": "Altezza cm",
           "Liegedauer": "Durata degenza", "Station": "Reparto"},
    "en": {"Fall-Nr": "Case no", "Geburtsdatum": "Date of birth", "Eintritt": "Admission", "Austritt": "Discharge",
           "Alter": "Age", "Gewicht kg": "Weight kg", "Größe": "Height cm", "Liegedauer": "Length of stay",
           "Station": "Ward"},
}


def _translated(lang: str, **planted) -> dict:
    cells = faelle(**planted)
    for c in range(1, 10):
        cells[(1, c)] = Cell(WORDS[lang][cells[(1, c)].value])
    return cells


@pytest.mark.parametrize("lang", ["fr", "it", "en"])
def test_a_clean_list_in_another_language_is_quiet(lang):
    found = kinds(report(_translated(lang)))
    assert not {"missing", "code", "impossible", "order", "future", "id", "same", "units", "outlier"} & set(found), found


@pytest.mark.parametrize("lang", ["fr", "it", "en"])
def test_planted_mistakes_are_found_in_another_language(lang):
    rep = report(_translated(lang, impossible=True, order=True, code=True))
    found = kinds(rep)
    assert found["order"].level == "fehler"
    assert found["order"].text.startswith(f"{WORDS[lang]['Austritt']} liegt vor {WORDS[lang]['Eintritt']} (Zeile 11")
    assert f"Darum zeigt H11 ({WORDS[lang]['Liegedauer']}" in found["order"].text
    assert "-4 in E10: Ein Alter unter 0 oder über 120 Jahren gibt es nicht" in texts(rep, "impossible")
    assert "steht 2-mal 999 (E5, E19)" in found["code"].text


def _summary(head: str, values: list[float], label: str, formula: str, shown: float) -> dict:
    cells = {(1, 1): Cell("Clinica"), (1, 2): Cell(head)}
    for r, v in enumerate(values, start=2):
        cells[(r, 1)] = Cell(f"Reparto {r - 1}")
        cells[(r, 2)] = Cell(str(v))
    if label:
        cells[(len(values) + 2, 1)] = Cell(label)
        cells[(len(values) + 2, 2)] = Cell(str(shown), formula)
    return cells


@pytest.mark.parametrize("label", ["Somme", "Somma", "Totale", "Total"])
def test_a_total_below_the_column_that_misses_a_row_is_found_in_every_language(label):
    values = [120.0, 95.0, 143.0, 88.0, 101.0, 77.0]
    rep = report(_summary("Costi CHF", values, label, "SUM(B2:B6)", sum(values[:5])))
    assert any(f.key == "parts:B8" or f.key == "label:B8" for f in rep.findings), [f.key for f in rep.findings]


@pytest.mark.parametrize("head", ["Part femmes", "Quota donne", "Share women", "Pourcentage"])
def test_a_share_above_one_is_impossible_in_every_language(head):
    values = [0.52, 0.48, 0.61, 3.43, 0.55, 0.47]
    found = {f.key: f for f in report(_summary(head, values, "", "", 0)).findings}
    assert "impossible:B5" in found and "Anteil unter 0 oder über 100 %" in found["impossible:B5"].text


@pytest.mark.parametrize("head", ["Part no", "Share price", "Sales quota"])
def test_ambiguous_share_words_with_other_numbers_stay_quiet(head):
    values = [12.0, 48.5, 61.0, 143.0, 55.0, 47.0]
    found = {f.key for f in report(_summary(head, values, "", "", 0)).findings}
    assert not any(k.startswith("impossible:") for k in found), found


def test_a_country_code_stays_the_group_a_figure_asks_for():
    """„Somme DE“ meint die Zeilen mit DE, nicht alle: „de“ ist kein Füllwort."""
    cells = {(1, 1): Cell("Pays"), (1, 2): Cell("Ventes")}
    for r, (land, v) in enumerate([("DE", 10), ("FR", 20), ("DE", 30), ("IT", 40), ("FR", 5), ("DE", 7)], start=2):
        cells[(r, 1)], cells[(r, 2)] = Cell(land), Cell(str(v))
    cells[(9, 1)], cells[(9, 2)] = Cell("Somme DE"), Cell("47", 'SUMIF(A2:A7,"DE",B2:B7)')
    found = {f.key for f in report(cells).findings}
    assert not any(k.endswith(":B9") for k in found), found


def test_media_in_an_english_budget_is_not_an_average():
    cells = {(1, 1): Cell("Item"), (1, 2): Cell("Budget")}
    for r, (item, v) in enumerate([("TV", 100), ("Radio", 40), ("Online", 80), ("Print", 30)], start=2):
        cells[(r, 1)], cells[(r, 2)] = Cell(item), Cell(str(v))
    cells[(6, 1)], cells[(6, 2)] = Cell("Media total"), Cell("250", "SUM(B2:B5)")
    found = {f.key for f in report(cells).findings}
    assert not any(k.endswith(":B6") for k in found), found


def test_a_row_named_like_the_start_of_a_foreign_word_is_still_a_data_row():
    """„Medizin“ beginnt wie „medi“ (italienisch Mittelwerte), „Sommer“ wie „somme“: Beides sind Zeilen der Liste."""
    from voiceagent import woerter

    assert not woerter.starts_with("Medizin", woerter.SUMMARY_ROWS)
    assert not woerter.starts_with("Sommerkurs", woerter.SUMMARY_ROWS)
    assert woerter.starts_with("Somme des coûts", woerter.SUMMARY_ROWS)
    cells = {(1, 1): Cell("Klinik"), (1, 2): Cell("Fälle")}
    for r, (name, n) in enumerate([("Chirurgie", 18), ("Kardiologie", 13), ("Medizin", 19), ("Neurologie", 8),
                                   ("Sommerkurs", 8)], start=2):
        cells[(r, 1)], cells[(r, 2)] = Cell(name), Cell(str(n))
    cells[(7, 1)], cells[(7, 2)] = Cell("Total"), Cell("66", "SUM(B2:B6)")
    assert not [f.key for f in report(cells).findings if f.key.endswith(":B7")]
