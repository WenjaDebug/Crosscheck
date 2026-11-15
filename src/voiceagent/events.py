"""Ereignisse, die WorkbookTracker aus zwei Ständen einer Mappe macht (eine Formel neu, ein Wert geändert)."""

from __future__ import annotations

from dataclasses import asdict, dataclass

# Ändern sich so viele Felder auf einmal, wurde ein anderer Datensatz geöffnet
# (z. B. eine andere Rechnung). Das ist keine Entscheidung.
RECORD_SWITCH = 3


@dataclass(frozen=True)
class Event:
    t: float
    # value_changed | record_changed | text_added | text_removed
    kind: str
    text: str
    label: str | None = None
    before: str | None = None
    after: str | None = None
    rule: str | None = None
    confirmed: bool | None = None
    # Die "Label: Wert"-Zeilen der Zeile nach der Änderung.
    context: tuple[str, ...] | None = None
    # Zwischenwerte, wenn mehrere schnelle Änderungen zusammengefasst wurden.
    via: tuple[str, ...] | None = None
    # Wer die Frage formuliert hat: „Google Gemini (gemini-2.5-flash)“ oder „Vorlage“.
    by: str | None = None

    def to_dict(self) -> dict:
        return {k: v for k, v in asdict(self).items() if v is not None}
