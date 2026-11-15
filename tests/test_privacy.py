

from voiceagent.privacy import Privacy

privacy = Privacy()


def test_bank_mail_and_phone_are_masked():
    assert privacy.line("IBAN: DE89 3704 0044 0532 0130 00") == "IBAN: [IBAN …3000]"
    assert privacy.line("Kontakt: rechnung@keller.example") == "Kontakt: [E-Mail]"
    assert privacy.line("Telefon: 0171 2345678") == "Telefon: [Telefon]"
    assert privacy.line("Tel.: +49 (0)30 123456-78") == "Tel.: [Telefon]"
    assert privacy.line("Karte: 4111 1111 1111 1111") == "Karte: [Karte …1111]"
    assert privacy.line("HU42 1177 3016 1111 1018 0000 0000") == "[IBAN …0000]"
    assert privacy.text("Ruf an unter 0171-2345678, bitte.") == "Ruf an unter [Telefon], bitte."


def test_sensitive_fields_are_masked_whatever_they_contain():
    assert privacy.line("Geburtsdatum: 03.05.1980") == "Geburtsdatum: [geschwärzt]"
    assert privacy.line("Steuer-ID: 12 345 678 901") == "Steuer-ID: [geschwärzt]"
    assert privacy.line("Telefon privat: 01712345678") == "Telefon privat: [geschwärzt]"
    # Ohne Doppelpunkt, wie die Texterkennung es manchmal liest.
    assert privacy.line("Telefon 01712345678") == "Telefon [geschwärzt]"
    assert Privacy(["Anschrift"]).line("Anschrift: Hauptstr. 1") == "Anschrift: [geschwärzt]"


def test_values_rules_depend_on_stay():
    lines = [
        "Kostenstelle: 0400",
        "Kostenstelle: 0000400100",
        "Sachkonto: 4400",
        "Betrag: 12.500,00 €",
        "Summe: 12.500,00 € +1.180,00 €",
        "Datum: 18.10.2026",
        "Zeitraum: 01/2026 - 03/2026",
        "PLZ: 04109",
        "Land: Ungarn",
        "USt-IdNr: DE123456789",
        "Bestellnummer: 4500012345",
        "Rechnung 4471 4472 4473 4474",
        "Kostenstelle 0400 4711 Gebucht",
        "Mobilität: hoch",
        "Telefonat am 3.5. mit Anna",
    ]
    assert privacy.lines(lines) == lines


def test_two_columns_read_as_one_line():
    # Nicht den Status schwärzen und die IBAN stehen lassen.
    line = "IBAN DE89370400440532013000 Status: Offen"
    assert privacy.line(line) == "IBAN [IBAN …3000] Status: Offen"


