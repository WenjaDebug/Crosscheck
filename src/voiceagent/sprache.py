"""Sprache der Hinweise: Deutsch oder Englisch.

Jeder Satz, den Gegenprobe sagt, steht im Code in beiden Sprachen: T(f"Die Summe in {cell} …", f"The sum in {cell} …").
Welche gilt, sagt die Anfrage des Seitenfensters (X-Gegenprobe-Sprache); ohne Angabe Deutsch. Die Sprache gilt für die
ganze Anfrage, auch in den Fäden, die sie startet (contextvars), und gehört zum Schlüssel, unter dem ein Ergebnis
gemerkt wird.

Zahlen, Daten und Formeln in den Sätzen schreibt Gegenprobe beim Rechnen deutsch („103,86“, „01.10.2024“,
„=SUMME(C3:C6)“). Auf Englisch macht output() daraus zum Schluss „103.86“, „2024-10-01“ und, wenn Excel Formeln
englisch zeigt (X-Gegenprobe-Formeln), „=SUM(C3:C6)“: So muss kein Satz die Zahlen selbst umschreiben, und die Prüfungen
rechnen weiter mit denselben Werten.
"""

from __future__ import annotations

import re
from contextlib import contextmanager
from contextvars import ContextVar

LANGS = ("de", "en")
LANG: ContextVar[str] = ContextVar("gegenprobe_sprache", default="de")
FORMULAS: ContextVar[str] = ContextVar("gegenprobe_formeln", default="de")


def lang() -> str:
    return LANG.get()


def english() -> bool:
    return LANG.get() == "en"


def T(german: str, english_text: str) -> str:
    """Der Satz in der Sprache der Anfrage."""
    return english_text if LANG.get() == "en" else german


def chosen(value: str | None) -> str:
    value = (value or "").strip().lower()[:2]
    return value if value in LANGS else "de"


@contextmanager
def using(language: str, formulas: str | None = None):
    """Für eine Anfrage oder einen Test: Hinweise in dieser Sprache, Formeln wie Excel sie zeigt."""
    a = LANG.set(chosen(language))
    b = FORMULAS.set(chosen(formulas or language))
    try:
        yield
    finally:
        FORMULAS.reset(b)
        LANG.reset(a)


def key() -> tuple[str, str]:
    """Gehört zum Schlüssel jedes gemerkten Ergebnisses."""
    return LANG.get(), FORMULAS.get()


# --- Zahlen, Daten und Formeln in fertigen Sätzen ---

# Eine Formel im Satz: „=SUMME(C3:C6)“, „=B5/12“, bis zum Leerzeichen; ein Punkt oder Komma am Ende gehört zum Satz.
_FORMULA = re.compile(r"(?<![\w\"'])=(?=[A-ZÄÖÜa-zäöü$(_\[\d])(?:\"[^\"]*\"|[^\s\"„“])+?(?=[.,;:]?(?:\s|$|“|\)?$))")
_DATE = re.compile(r"(?<![\d.])(\d{2})\.(\d{2})\.(\d{4})(?!\d|\.\d|,\d)")
_THOUSANDS = re.compile(r"(?<![\d.,])(\d{1,3}(?:\.\d{3})+)(?:,(\d+))?(?!\d|\.\d|,\d)")
_DECIMAL = re.compile(r"(?<![\d.,A-Za-z$])(\d+),(\d+)(?!\d|,\d|\.\d)")
_SPACED = re.compile(r"(?<=\d) (?=\d{3}(?!\d))")
_PERCENT = re.compile(r"(?<=\d) (?=%)")


def _english_formula(text: str) -> str:
    from .einsetzen import english as convert
    return convert(text)


def _numbers(text: str) -> str:
    """103,86 → 103.86, 1.234,5 → 1,234.5, 01.10.2024 → 2024-10-01, 12,5 % → 12.5% (nur außerhalb von Formeln)."""
    text = _DATE.sub(lambda m: f"{m.group(3)}-{m.group(2)}-{m.group(1)}", text)
    text = _THOUSANDS.sub(lambda m: m.group(1).replace(".", ",") + (f".{m.group(2)}" if m.group(2) else ""), text)
    text = _DECIMAL.sub(r"\1.\2", text)
    text = _PERCENT.sub("", text)
    return text


# Text aus der Mappe steht im Satz in Anführungszeichen: Er bleibt, wie er in der Zelle steht (nur eine Formel darin
# zeigt Excel in seiner Sprache). Auf Englisch mit englischen Anführungszeichen.
_QUOTED = re.compile(r"„([^“„]*)“|“([^”“]*)”")
# Fehlerwerte zeigt Excel in seiner Sprache.
_ERRORS = {"#WERT!": "#VALUE!", "#NV": "#N/A", "#BEZUG!": "#REF!", "#ZAHL!": "#NUM!", "#ÜBERLAUF!": "#SPILL!",
           "#KALK!": "#CALC!", "#FELD!": "#FIELD!"}
_ERROR = re.compile("|".join(re.escape(code) for code in sorted(_ERRORS, key=len, reverse=True)) + r"(?![\w/])")


def _plain(text: str, numbers: bool) -> str:
    """Satzteil ohne Anführungszeichen: Formeln wie Excel sie zeigt, sonst Zahlen und Daten in der Sprache."""
    english_formulas = FORMULAS.get() == "en"
    out, last = [], 0
    for m in _FORMULA.finditer(text):
        before = text[last:m.start()]
        out.append(_numbers(before) if numbers else before)
        out.append(_english_formula(m.group(0)) if english_formulas else m.group(0))
        last = m.end()
    rest = text[last:]
    out.append(_numbers(rest) if numbers else rest)
    done = "".join(out)
    return _ERROR.sub(lambda m: _ERRORS[m.group(0)], done) if english_formulas else done


def sentence(text: str) -> str:
    """Ein fertiger Satz, wie er in der Sprache der Anfrage aussieht."""
    if LANG.get() != "en" and FORMULAS.get() != "en":
        return text
    english = LANG.get() == "en"
    out, last = [], 0
    for m in _QUOTED.finditer(text):
        out.append(_plain(text[last:m.start()], english))
        inner = _plain(m.group(1) if m.group(1) is not None else m.group(2), False)
        german = m.group(1) is not None
        out.append(f"“{inner}”" if english or not german else f"„{inner}“")
        last = m.end()
    out.append(_plain(text[last:], english))
    return "".join(out)


def formula(text: str) -> str:
    """Eine ganze Formel (deutsch geschrieben), wie Excel sie in seiner Sprache zeigt."""
    if FORMULAS.get() == "en" and str(text).startswith("="):
        return _english_formula(text)
    return text


# Schlüssel, unter denen in einer Antwort Sätze stehen. Alles andere (Zellen, Formeln zum Einsetzen, Schlüssel,
# Namen aus der Mappe) bleibt, wie es ist.
SENTENCES = {"text", "tip", "title", "label", "words", "said", "verdict", "note", "rows", "detail", "why",
             "line", "summary", "head", "lead", "jetzt", "beschreibung", "grund", "satz", "wrong"}
FORMULA_KEYS = {"formula", "want", "now"}


def output(value, parent: str = ""):
    """Eine Antwort an das Seitenfenster: Sätze in der Sprache der Anfrage (sentence), Formeln wie Excel sie zeigt.
    Korrekturen (fix) behalten ihre Zellen; nur was das Fenster anzeigt, wird umgeschrieben."""
    if LANG.get() != "en" and FORMULAS.get() != "en":
        return value
    if isinstance(value, dict):
        out = {}
        for k, v in value.items():
            if k in ("cells", "was", "german", "english", "id", "key", "parts", "range", "general", "decimals"):
                out[k] = v
            elif isinstance(v, str) and k in FORMULA_KEYS and v.startswith("="):
                out[k] = formula(v)
            elif isinstance(v, str) and (k in SENTENCES or k in FORMULA_KEYS):
                out[k] = sentence(v)
            elif isinstance(v, list) and k in SENTENCES:
                out[k] = [sentence(x) if isinstance(x, str) else output(x, k) for x in v]
            else:
                out[k] = output(v, k)
        return out
    if isinstance(value, list):
        return [output(v, parent) for v in value]
    return value
