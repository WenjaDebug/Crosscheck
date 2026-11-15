"""Dritte Testrunde mit KI-erstellten Mappen: Tabelle2 der Finanzfallen, „Fallbeispiel Hausarztpraxis“, mit 5
eingebauten Fehlern. Gegenprobe zeigte 11 Stellen, davon 8 Fehlalarme, und fand 2 der 5 Fehler.

Die Ursachen waren allgemein:
- „Gesamterlös“ verlor als Summenwort sein Maß („Erlös“): Jede Spalte links davon galt als Teil (6 Hinweise).
- Der Schutz vor dem Teilen durch 0 (=WENN(H18=0;0;K18/H18)) ließ die Quote der Summenzeile wie einen Bruch im
  Muster aussehen, und die Zeile verlangte eine Summe der Margen.
- „Nur gerundet“ hieß jede Abweichung unter 0,005, auch 0,1155 statt 0,119.
- Das Mittel der Monatsmargen (B64) und 12 − Startmonat (B34) kannte keine Prüfung; C30 hatte kein Einsetzen.
- =PROGNOSE(13;…) über die Monate 1 bis 12 galt als versteckte Annahme."""

import pytest
from test_dritte_serie import PERCENT, _hints

P = PERCENT


def _wrap(cells: dict) -> dict:
    return {k: ((v[0], P) if isinstance(v, tuple) else v) for k, v in cells.items()}


PLANTED = _wrap({
    'A1': 'Fallbeispiel Hausarztpraxis – Jahresanalyse 2025',
    'A2': 'Beträge in €. Eingaben blau, Formeln schwarz. Alle Werte sind frei erfundene Beispieldaten.',
    'A4': '1. Annahmen', 'A5': 'Parameter', 'B5': 'Wert', 'C5': 'Einheit', 'A6': 'Punktwert EBM', 'B6': 0.119,
    'C6': '€ je Punkt', 'A7': 'Sachkostenquote', 'B7': (0.14, P), 'C7': '% vom Gesamterlös',
    'A8': 'Zinssatz Darlehen', 'B8': (0.042, P), 'C8': '% p.a.', 'A9': 'Nutzungsdauer Ultraschallgerät', 'B9': 8,
    'C9': 'Jahre', 'A10': 'Anschaffungskosten Ultraschallgerät', 'B10': 96000, 'C10': '€',
    'A11': 'Inbetriebnahme (Monat)', 'B11': 7, 'C11': 'Monat 1–12', 'A12': 'Darlehenssumme', 'B12': 150000,
    'C12': '€', 'A13': 'Jährliche Tilgung', 'B13': 15000, 'C13': '€ p.a.', 'A16': '2. Monatsauswertung 2025',
    'A17': 'Monat', 'B17': 'Nr.', 'C17': 'Behandlungsfälle', 'D17': 'Punkte je Fall', 'E17': 'Punktwert (€)',
    'F17': 'KV-Umsatz (€)', 'G17': 'Privatumsatz (€)', 'H17': 'Gesamterlös (€)', 'I17': 'Personalkosten (€)',
    'J17': 'Sachkosten (€)', 'K17': 'Ergebnis (€)', 'L17': 'Marge (%)', 'A18': 'Jan', 'B18': 1, 'C18': 612,
    'D18': 455, 'E18': '=$B$6', 'F18': '=C18*D18*E18', 'G18': 5100, 'H18': '=F18+G18', 'I18': 26400,
    'J18': '=H18*$B$7', 'K18': '=H18-I18-J18', 'L18': ('=IF(H18=0,0,K18/H18)', P), 'A19': 'Feb', 'B19': 2,
    'C19': 588, 'D19': 462, 'E19': '=$B$6', 'F19': '=C19*D19*E19', 'G19': 4800, 'H19': '=F19+G19', 'I19': 26400,
    'J19': '=H19*$B$7', 'K19': '=H19-I19-J19', 'L19': ('=IF(H19=0,0,K19/H19)', P), 'A20': 'Mär', 'B20': 3,
    'C20': 641, 'D20': 449, 'E20': '=$B$6', 'F20': '=C20*D20*E20', 'G20': 5350, 'H20': '=F20+G20', 'I20': 26800,
    'J20': '=H20*$B$7', 'K20': '=H20-I20-J20', 'L20': ('=IF(H20=0,0,K20/H20)', P), 'A21': 'Apr', 'B21': 4,
    'C21': 602, 'D21': 470, 'E21': '=$B$6', 'F21': '=C21*D21*E21', 'G21': 4950, 'H21': '=F21+G21', 'I21': 26800,
    'J21': '=H21*$B$7', 'K21': '=H21-I21-J21', 'L21': ('=IF(H21=0,0,K21/H21)', P), 'A22': 'Mai', 'B22': 5,
    'C22': 655, 'D22': 458, 'E22': '=$B$6', 'F22': '=C22*D22*E22', 'G22': 5600, 'H22': '=F22+G22', 'I22': 27100,
    'J22': '=H22*$B$7', 'K22': '=H22-I22-J22', 'L22': ('=IF(H22=0,0,K22/H22)', P), 'A23': 'Jun', 'B23': 6,
    'C23': 630, 'D23': 466, 'E23': 0.1155, 'F23': '=C23*D23*E23', 'G23': 5200, 'H23': '=F23+G23', 'I23': 27100,
    'J23': '=H23*$B$7', 'K23': '=H23-I23-J23', 'L23': ('=IF(H23=0,0,K23/H23)', P), 'A24': 'Jul', 'B24': 7,
    'C24': 578, 'D24': 441, 'E24': '=$B$6', 'F24': '=C24*D24*E24', 'G24': 4700, 'H24': '=F24+G24', 'I24': 27100,
    'J24': '=H24*$B$7', 'K24': '=H24-I24-J24', 'L24': ('=IF(H24=0,0,K24/H24)', P), 'A25': 'Aug', 'B25': 8,
    'C25': 541, 'D25': 438, 'E25': '=$B$6', 'F25': '=C25*D25*E25', 'G25': 4300, 'H25': '=F25+G25', 'I25': 27300,
    'J25': '=H25*$B$7', 'K25': '=H25-I25-J25', 'L25': ('=IF(H25=0,0,K25/H25)', P), 'A26': 'Sep', 'B26': 9,
    'C26': 637, 'D26': 472, 'E26': '=$B$6', 'F26': '=C26*D26*E26', 'G26': 5400, 'H26': '=F26+G26', 'I26': 27300,
    'J26': '=H26*$B$7', 'K26': '=H26-I26-J26', 'L26': ('=IF(H26=0,0,K26/H26)', P), 'A27': 'Okt', 'B27': 10,
    'C27': 668, 'D27': 480, 'E27': '=$B$6', 'F27': '=C27*D27*E27', 'G27': 5700, 'H27': '=F27+G27', 'I27': 27600,
    'J27': '=H27*$B$7', 'K27': '=H27-I27-J27', 'L27': ('=IF(H27=0,0,K27/H27)', P), 'A28': 'Nov', 'B28': 11,
    'C28': 649, 'D28': 463, 'E28': '=$B$6', 'F28': '=C28*D28*E28', 'G28': 5450, 'H28': '=F28+G28', 'I28': 27600,
    'J28': '=H28*$B$7', 'K28': '=H28-I28-J28', 'L28': ('=IF(H28=0,0,K28/H28)', P), 'A29': 'Dez', 'B29': 12,
    'C29': 597, 'D29': 452, 'E29': '=$B$6', 'F29': '=C29*D29*E29', 'G29': 4900, 'H29': '=F29+G29', 'I29': 27600,
    'J29': '=H29*$B$7', 'K29': '=H29-I29-J29', 'L29': ('=IF(H29=0,0,K29/H29)', P), 'A30': 'Jahressumme',
    'C30': '=SUM(C18:C28)', 'F30': '=SUM(F18:F29)', 'G30': '=SUM(G18:G29)', 'H30': '=SUM(H18:H29)',
    'I30': '=SUM(I18:I29)', 'J30': '=SUM(J18:J29)', 'K30': '=SUM(K18:K29)', 'L30': ('=K30/H30', P),
    'A32': '3. Abschreibung Ultraschallgerät (linear, pro rata temporis)', 'A33': 'Jahr', 'B33': 'AfA-Monate',
    'C33': 'AfA (€)', 'D33': 'Kumulierte AfA (€)', 'E33': 'Restbuchwert (€)', 'A34': 2025, 'B34': '=12-$B$11',
    'C34': '=$B$10/($B$9*12)*B34', 'D34': '=C34', 'E34': '=$B$10-D34', 'A35': '=A34+1',
    'B35': '=MIN(12,$B$9*12-SUM($B$34:B34))', 'C35': '=$B$10/($B$9*12)*B35', 'D35': '=D34+C35', 'E35': '=$B$10-D35',
    'A36': '=A35+1', 'B36': '=MIN(12,$B$9*12-SUM($B$34:B35))', 'C36': '=$B$10/($B$9*12)*B36', 'D36': '=D35+C36',
    'E36': '=$B$10-D36', 'A37': '=A36+1', 'B37': '=MIN(12,$B$9*12-SUM($B$34:B36))', 'C37': '=$B$10/($B$9*12)*B37',
    'D37': '=D36+C37', 'E37': '=$B$10-D37', 'A38': '=A37+1', 'B38': '=MIN(12,$B$9*12-SUM($B$34:B37))',
    'C38': '=$B$10/($B$9*12)*B38', 'D38': '=D37+C38', 'E38': '=$B$10-D38', 'A39': '=A38+1',
    'B39': '=MIN(12,$B$9*12-SUM($B$34:B38))', 'C39': '=$B$10/($B$9*12)*B39', 'D39': '=D38+C39', 'E39': '=$B$10-D39',
    'A40': '=A39+1', 'B40': '=MIN(12,$B$9*12-SUM($B$34:B39))', 'C40': '=$B$10/($B$9*12)*B40', 'D40': '=D39+C40',
    'E40': '=$B$10-D40', 'A41': '=A40+1', 'B41': '=MIN(12,$B$9*12-SUM($B$34:B40))', 'C41': '=$B$10/($B$9*12)*B41',
    'D41': '=D40+C41', 'E41': '=$B$10-D41', 'A42': '=A41+1', 'B42': '=MIN(12,$B$9*12-SUM($B$34:B41))',
    'C42': '=$B$10/($B$9*12)*B42', 'D42': '=D41+C42', 'E42': '=$B$10-D42', 'A43': 'Summe', 'B43': '=SUM(B34:B42)',
    'C43': '=SUM(C34:C42)', 'A45': '4. Darlehen – feste Tilgung', 'A46': 'Jahr',
    'B46': 'Restschuld Jahresanfang (€)', 'C46': 'Zinsen (€)', 'D46': 'Tilgung (€)',
    'E46': 'Restschuld Jahresende (€)', 'F46': 'Zahlungen gesamt (€)', 'A47': 1, 'B47': '=$B$12', 'C47': '=B47*$B$8',
    'D47': '=MIN($B$13,B47)', 'E47': '=B47-D47', 'F47': '=C47+D47', 'A48': '=A47+1', 'B48': '=E47',
    'C48': '=B48*$B$8', 'D48': '=MIN($B$13,B48)', 'E48': '=B48-D48', 'F48': '=C48+D48', 'A49': '=A48+1',
    'B49': '=E48', 'C49': '=E49*$B$8', 'D49': '=MIN($B$13,B49)', 'E49': '=B49-D49', 'F49': '=C49+D49',
    'A50': '=A49+1', 'B50': '=E49', 'C50': '=B50*$B$8', 'D50': '=MIN($B$13,B50)', 'E50': '=B50-D50',
    'F50': '=C50+D50', 'A51': '=A50+1', 'B51': '=E50', 'C51': '=B51*$B$8', 'D51': '=MIN($B$13,B51)',
    'E51': '=B51-D51', 'F51': '=C51+D51', 'A52': '=A51+1', 'B52': '=E51', 'C52': '=B52*$B$8',
    'D52': '=MIN($B$13,B52)', 'E52': '=B52-D52', 'F52': '=C52+D52', 'A53': '=A52+1', 'B53': '=E52',
    'C53': '=B53*$B$8', 'D53': '=MIN($B$13,B53)', 'E53': '=B53-D53', 'F53': '=C53+D53', 'A54': '=A53+1',
    'B54': '=E53', 'C54': '=B54*$B$8', 'D54': '=MIN($B$13,B54)', 'E54': '=B54-D54', 'F54': '=C54+D54',
    'A55': '=A54+1', 'B55': '=E54', 'C55': '=B55*$B$8', 'D55': '=MIN($B$13,B55)', 'E55': '=B55-D55',
    'F55': '=C55+D55', 'A56': '=A55+1', 'B56': '=E55', 'C56': '=B56*$B$8', 'D56': '=MIN($B$13,B56)',
    'E56': '=B56-D56', 'F56': '=C56+D56', 'A57': 'Summe', 'C57': '=SUM(C47:C56)', 'D57': '=SUM(D47:D56)',
    'F57': '=SUM(F47:F56)', 'A59': '5. Kennzahlen & Prognose', 'A60': 'Kennzahl', 'B60': 'Wert', 'C60': 'Berechnung',
    'A61': 'Gesamterlös 2025 (€)', 'B61': '=H30', 'C61': 'Summe Monatswerte', 'A62': 'Betriebsergebnis 2025 (€)',
    'B62': '=K30', 'C62': 'vor AfA und Zins', 'A63': 'Erlös je Behandlungsfall (€)', 'B63': '=H30/C30',
    'C63': 'Gesamterlös / Fälle', 'A64': 'Jahresmarge (%)', 'B64': ('=AVERAGE(L18:L29)', P),
    'C64': 'Marge über das Jahr', 'A65': 'Zinsaufwand 2025 (€)', 'B65': '=C47', 'C65': 'Jahr 1 aus Tilgungsplan',
    'A66': 'AfA 2025 (€)', 'B66': '=C34', 'C66': 'Jahr 1 aus AfA-Plan',
    'A67': 'Jahresergebnis nach AfA und Zins (€)', 'B67': '=B62-B65-B66', 'C67': 'Betriebsergebnis − Zins − AfA',
    'A68': 'Prognose Erlös Januar 2026 (€)', 'B68': '=FORECAST(13,H18:H29,B18:B29)',
    'C68': 'Lineare Trendfortschreibung', 'A69': 'Prognose Ergebnis Januar 2026 (€)',
    'B69': '=FORECAST(13,K18:K29,B18:B29)', 'C69': 'Lineare Trendfortschreibung',
    'A70': 'Bestimmtheitsmaß R² (Erlös)', 'B70': '=RSQ(H18:H29,B18:B29)', 'C70': 'Güte des Trends',
    'A72': '6. Plausibilitätsprüfungen', 'A73': 'Prüfung', 'B73': 'Status',
    'A74': 'Erlös = KV-Umsatz + Privatumsatz', 'B74': '=IF(ROUND(H30-(F30+G30),2)=0,"OK","Fehler")',
    'A75': 'AfA-Summe = Anschaffungskosten', 'B75': '=IF(ROUND(C43-B10,2)=0,"OK","Fehler")',
    'A76': 'Darlehen vollständig getilgt', 'B76': '=IF(ROUND(E56,2)=0,"OK","Fehler")',
    'A77': 'Ergebnis = Erlös − Personal − Sachkosten', 'B77': '=IF(ROUND(K30-(H30-I30-J30),2)=0,"OK","Fehler")',
})
RIGHT = {**PLANTED, "C30": "=SUM(C18:C29)", "B64": ("=K30/H30", P), "B34": "=13-$B$11", "C49": "=B49*$B$8",
         "E23": "=$B$6"}
WRONG = {"typedcell:E23", "gap:C30", "ratio:B64", "pattern:C49", "startmonat:B34"}


def _found(cells: dict) -> dict:
    return {k: h for (s, k), h in _hints(("Tabelle2", cells)).items() if h["level"] != "info"}


def test_alle_fuenf_fehler_ohne_fehlalarm():
    found = _found(PLANTED)
    assert set(found) == WRONG, sorted(found)
    for key in WRONG:
        assert found[key]["level"] == "fehler" and found[key].get("fix"), key


def test_einsetzen_formeln():
    found = _found(PLANTED)
    assert found["gap:C30"]["fix"]["cells"] == {"C30": "=SUM(C18:C29)"}
    assert found["ratio:B64"]["fix"]["cells"] == {"B64": "=SUM(K18:K29)/SUM(H18:H29)"}
    assert found["startmonat:B34"]["fix"]["cells"] == {"B34": "=13-$B$11"}
    assert "Juli bis Dezember sind 6 Monate" in found["startmonat:B34"]["text"]
    assert "gerundet" not in found["typedcell:E23"]["text"]


def test_richtige_tabelle_ist_still():
    assert _found(RIGHT) == {}


@pytest.mark.parametrize("change", [
    {"B34": "=12-$B$11+1"},                       # zählt den Startmonat schon mit
    {"A11": "Monate vor Inbetriebnahme", "B11": 6},  # verstrichene Monate: 12 − 6 stimmt
    {"B64": ("=AVERAGE(L18:L29)", P), "A64": "Marge im Monatsmittel (ungewichtet)"},
])
def test_aehnliche_richtige_faelle_still(change):
    found = _found({**RIGHT, **change})
    assert not {k for k in found if k.split(":")[0] in ("startmonat", "ratio")}, sorted(found)


def test_prognose_mit_anderer_zahl_bleibt_hinweis():
    # 12,5 liegt nicht auf dem Raster der Monate 1 bis 12: eine echte Annahme, die in eine eigene Zelle gehört.
    found = _found({**RIGHT, "B68": "=FORECAST(12.5,H18:H29,B18:B29)", "B69": "=FORECAST(12.5,K18:K29,B18:B29)"})
    assert "constant:12,5" in found, sorted(found)
