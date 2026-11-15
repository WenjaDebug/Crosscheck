"""Datenschutz: IBAN, E-Mail, Telefon und Ähnliches schwärzen, bevor etwas im Fehlerbericht landet.

Von IBAN und Kartennummer bleiben wie auf Kontoauszügen die letzten vier Zeichen stehen: Man sieht noch, dass
sich etwas geändert hat, aber nicht die Nummer.

Zweierlei wird geschwärzt:
- Werte, die wie eine IBAN, E-Mail-Adresse, Telefon- oder Kartennummer aussehen, egal wo.
- Der ganze Wert von Feldern wie „Geburtsdatum“, „Steuer-ID“ oder „Passwort“
  (LABELS).
Namen und Anschriften bleiben stehen.
"""

from __future__ import annotations

import re


HIDDEN = "[geschwärzt]"

# Feldnamen, deren Wert immer geschwärzt wird. „Konto“ allein fehlt mit Absicht:
# In der Buchhaltung ist das meist ein Sachkonto, und das hilft beim Lesen eines Berichts.
LABELS = (
    r"iban", r"bic", r"swift(?:-?code)?", r"bankverbindung", r"bankkonto", r"kontoinhaber\w*",
    r"e-?mail(?:-?adresse)?", r"telefon\w*", r"tel", r"mobil(?:nummer|telefon|funk|nr)?", r"handy\w*",
    r"fax\w*", r"phone", r"mobile",
    r"geburts\w*", r"date of birth", r"birthday",
    r"steuer-?id", r"steueridentifikationsnummer", r"sozialversicherungs\w*", r"sv-?nummer", r"ssn",
    r"personalausweis\w*", r"ausweis\w*", r"reisepass\w*", r"passnummer",
    r"passwort", r"kennwort", r"password", r"pin", r"tan", r"puk",
    r"kreditkart\w*", r"karten-?nummer", r"kartennr", r"card number", r"credit card", r"cvc", r"cvv",
    r"account number",
)

_LABEL = re.compile(r"^(\s*([^:]{2,40}?)\s*:\s*)(.*)$")  # wie events._LABEL
_PLACEHOLDER = re.compile(r"\[(?:IBAN …\w{1,4}|Karte …\d{1,4}|E-Mail|Telefon|geschwärzt)\]")
_ONLY_PLACEHOLDERS = re.compile(rf"^(?:{_PLACEHOLDER.pattern}[\s,;/]*)+$")

_EMAIL = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
_IBAN = re.compile(r"\b([A-Z]{2})(\d{2})((?: ?[A-Z0-9]){11,30})\b")
# Länge der IBAN in den häufigsten Ländern. Andere Länder zählen nur mit gültiger Prüfziffer.
_IBAN_LENGTH = {
    "DE": 22, "AT": 20, "CH": 21, "LI": 21, "FR": 27, "IT": 27, "ES": 24, "PT": 25, "NL": 18,
    "BE": 16, "LU": 20, "PL": 28, "CZ": 24, "SK": 24, "HU": 28, "SI": 19, "HR": 21, "RO": 24,
    "BG": 22, "GB": 22, "IE": 22, "DK": 18, "SE": 24, "NO": 15, "FI": 18, "EE": 20, "LV": 21,
    "LT": 20, "GR": 27, "TR": 26,
}
_CARD = re.compile(
    r"(?<![\d.,])(?:(?:\d{4}[ -]){3}\d{1,7}|\d{4}[ -]\d{6}[ -]\d{5}|\d{13,19})(?![\d]|[.,]\d)"
)
# International mit +49 oder 0049, sonst nur mit Vorwahl und Leerzeichen, Schrägstrich
# oder Bindestrich (0171 2345678): Eine Kostenstelle wie 0400 oder ein Datum bleibt stehen.
_PHONE_WORLD = re.compile(r"(?<![\w+])(?:\+|00)[1-9](?:[ ./()-]{0,2}\d){6,14}(?!\d)")
_PHONE_HOME = re.compile(r"(?<![\w.,/+-])0[1-9]\d{1,4}[ /-]+\d(?:[ -]?\d){4,10}(?![\d]|[.,]\d)")


class Privacy:
    def __init__(self, extra_labels: tuple[str, ...] | list[str] = ()) -> None:
        words = list(LABELS) + [re.escape(label.strip().lower()) for label in extra_labels if label.strip()]
        self._sensitive = re.compile(r"\b(?:" + "|".join(words) + r")\b", re.IGNORECASE)
        self._leading = re.compile(r"^\W*(?:" + "|".join(words) + r")\b[\s.:#-]*", re.IGNORECASE)

    # --- Text ---

    def text(self, text: str) -> str:
        """Schwärzt IBAN, E-Mail, Telefon- und Kartennummern in einem beliebigen Text."""
        return _apply(text, _pattern_spans(text))

    def line(self, line: str) -> str:
        """Eine Zeile, z. B. „IBAN: DE89 …“ → „IBAN: [IBAN …3000]“."""
        return _apply(line, self._line_spans(line))

    def lines(self, lines: list[str]) -> list[str]:
        return [self.line(line) for line in lines]

    def _value(self, value: str) -> str:
        """Der Wert eines Feldes, das immer geschwärzt wird."""
        found = [placeholder for _, _, placeholder in _pattern_spans(value)]
        return " ".join(found) if found else HIDDEN

    def _line_spans(self, line: str) -> list[tuple[int, int, str]]:
        match = _LABEL.match(line)
        if match and match.group(3).strip():
            if self._is_label(match.group(2)):
                return [(match.end(1), len(line.rstrip()), self._value(match.group(3)))]
            return _pattern_spans(line)
        # Ohne Doppelpunkt („Telefon 0171 2345678“): Rest schwärzen, wenn er nach Nummer aussieht.
        lead = self._leading.match(line)
        if lead and not match:
            rest = line[lead.end():]
            if sum(ch.isdigit() for ch in rest) >= 3 or "@" in rest:
                return [(lead.end(), len(line.rstrip()), self._value(rest))]
        return _pattern_spans(line)

    def _is_label(self, label: str) -> bool:
        """„E-Mail“ oder „Telefon privat“, aber nicht zwei zusammengelesene Spalten wie
        „IBAN DE89 … Status“: Sonst würde der Status geschwärzt und die IBAN nicht."""
        return len(label.split()) <= 4 and bool(self._sensitive.search(label)) and not _pattern_spans(label)

    def _label_only(self, text: str) -> bool:
        """Nur ein Feldname wie „E-Mail“ oder „IBAN:“; der Wert steht rechts daneben."""
        label = text.strip().rstrip(":*").strip()
        lead = self._leading.match(label)
        return bool(label) and len(label) <= 40 and lead is not None and lead.end() >= len(label)


def _apply(text: str, spans: list[tuple[int, int, str]]) -> str:
    for start, end, placeholder in sorted(spans, reverse=True):
        text = text[:start] + placeholder + text[end:]
    return text


def _pattern_spans(text: str) -> list[tuple[int, int, str]]:
    """Fundstellen (Anfang, Ende, Ersatz), ohne Überschneidungen."""
    spans: list[tuple[int, int, str]] = []

    def add(start: int, end: int, placeholder: str) -> None:
        if not any(start < e and s < end for s, e, _ in spans):
            spans.append((start, end, placeholder))

    for match in _EMAIL.finditer(text):
        add(match.start(), match.end(), "[E-Mail]")
    for match in _IBAN.finditer(text):
        found = _iban(match)
        if found:
            add(match.start(), match.start() + found[0], f"[IBAN …{found[1][-4:]}]")
    for match in _CARD.finditer(text):
        digits = re.sub(r"\D", "", match.group(0))
        if 13 <= len(digits) <= 19 and digits[0] in "3456" and _luhn(digits):
            add(match.start(), match.end(), f"[Karte …{digits[-4:]}]")
    for pattern in (_PHONE_WORLD, _PHONE_HOME):
        for match in pattern.finditer(text):
            add(match.start(), match.end(), "[Telefon]")
    return sorted(spans)


def _iban(match: re.Match) -> tuple[int, str] | None:
    """Länge im Text und die IBAN ohne Leerzeichen, oder None, wenn es keine ist."""
    raw = match.group(0)
    compact = raw.replace(" ", "")
    country = match.group(1)
    if country in _IBAN_LENGTH:
        compact = compact[: _IBAN_LENGTH[country]]
        if sum(ch.isdigit() for ch in compact[4:]) < 8:
            return None
    else:
        compact = next((compact[:n] for n in range(len(compact), 14, -1) if _iban_valid(compact[:n])), "")
        if not compact:
            return None
    # Wie weit reicht die IBAN im Text (mit Leerzeichen)?
    count = 0
    for index, ch in enumerate(raw):
        if ch != " ":
            count += 1
            if count == len(compact):
                return index + 1, compact
    return len(raw), compact


def _iban_valid(compact: str) -> bool:
    moved = compact[4:] + compact[:4]
    digits = "".join(str(int(ch, 36)) for ch in moved)
    return int(digits) % 97 == 1


def _luhn(digits: str) -> bool:
    total = 0
    for index, ch in enumerate(reversed(digits)):
        value = int(ch) * (2 if index % 2 else 1)
        total += value - 9 if value > 9 else value
    return total % 10 == 0
