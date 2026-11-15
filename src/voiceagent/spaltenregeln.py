"""Eigene Regeln je Spalte: einmal festlegen, was in einer Spalte stehen darf, und jede Mappe hält sich daran.

Eine Regel gehört zu einem Spaltennamen („Alter“) und gilt in jeder Tabelle mit einer Spalte dieses Namens, auf
jedem Blatt oder nur auf einem. Groß- und Kleinschreibung und doppelte Leerzeichen im Namen zählen nicht. Was eine
Regel verlangen kann:

- Pflichtfeld: keine leere Zelle, kein „k.A.“ oder „-“ statt eines Werts.
- Art: Zahl, ganze Zahl, Datum oder Text. Eine Zahl als Text, ein Datum ohne Datumsformat (45000) und ein Fehlerwert
  sind Verstöße, mit dem Grund.
- Grenzen: Zahlen von … bis … (Prozentzellen in Prozent, wie Excel sie zeigt), Daten nicht vor oder nach einem Tag,
  nicht in der Zukunft.
- Erlaubte Werte („m, w, d“ oder „1; 2; 3“): Was nur anders geschrieben ist („M“, „m “ mit Leerzeichen), heißt so.
- Jeder Wert nur einmal (Fallnummern), ein Muster („PAT-####“: # eine Ziffer, @ ein Buchstabe, ? ein Zeichen,
  * beliebig viele) und ein Vergleich mit einer anderen Spalte derselben Zeile (Austritt nicht vor Eintritt).

Jeder Verstoß kommt mit Zelle und Grund, in einem Hinweis je Spalte. Was eine eigene Regel genauer sagt, ersetzt die
allgemeinen Hinweise dazu (unmöglicher Wert, Nummer mehrmals, Datum in der Zukunft, Ende vor dem Anfang): Die Regel
ist die eigene Entscheidung, was erlaubt ist. Die Regeln stehen in data/spaltenregeln.json, nur auf diesem Rechner.
Alles ohne KI.
"""

from __future__ import annotations

import json
import math
import os
import re
from collections import Counter
from datetime import date
from pathlib import Path

from .analyse import Finding, Sheet, _fmt, address
from .auswertung import _PLACEHOLDERS, _data_rows, codes_in, order_start
from .formel import Err, serial_of
from .sprache import T, english
from .workbook import layout
from .xlsx import Sheet as Book
from .xlsx import column_letters, number_text

Pos = tuple[int, int]

FILE = "spaltenregeln.json"
KINDS = {"": "beliebig", "zahl": "Zahl", "ganz": "ganze Zahl", "datum": "Datum", "text": "Text"}
_KINDS_EN = {"": "any", "zahl": "number", "ganz": "whole number", "datum": "date", "text": "text"}
RELATIONS = ("ge", "le")
MAX_RULES = 200
MAX_VALUES = 200
MAX_TEXT = 200
# So viele Verstöße nennt ein Hinweis mit Zelle und Grund, so viele Zellen markiert ein Klick.
MAX_LISTED = 5
MAX_AREAS = 25
_EPOCH = date(1899, 12, 30)


# --- Regeln lesen, prüfen, speichern ---


def norm(name) -> str:
    return " ".join(str(name or "").split()).casefold()


def path(data_dir: Path) -> Path:
    return Path(data_dir) / FILE


def load(data_dir: Path | None) -> list[dict]:
    """Die gespeicherten Regeln; eine kaputte oder fremde Datei gibt keine Regeln (und wird beim Speichern beiseitegelegt)."""
    if data_dir is None or not path(data_dir).exists():
        return []
    try:
        saved = json.loads(path(data_dir).read_text(encoding="utf-8"))
        return [clean(r) for r in saved.get("regeln", [])][:MAX_RULES]
    except (ValueError, TypeError, AttributeError, OSError):
        return []


def _write(data_dir: Path, rules: list[dict]) -> None:
    target = path(data_dir)
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        try:
            json.loads(target.read_text(encoding="utf-8"))["regeln"]
        except (ValueError, TypeError, KeyError, OSError):
            target.replace(target.with_name(FILE + ".kaputt"))
    temporary = target.with_name(FILE + ".neu")
    temporary.write_text(json.dumps({"regeln": rules}, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, target)


def _same(a: dict, b: dict) -> bool:
    return norm(a["spalte"]) == norm(b["spalte"]) and norm(a["blatt"]) == norm(b["blatt"])


def save(data_dir: Path, rule: dict, before: dict | None = None) -> list[dict]:
    """Eine Regel speichern; eine für dieselbe Spalte (und dasselbe Blatt) ersetzt sie, und before (Spalte und Blatt
    der Regel, wie sie vor dem Ändern hieß) auch. ValueError: was nicht stimmt."""
    rule = clean(rule)
    old = {"spalte": str(before.get("spalte") or ""), "blatt": str(before.get("blatt") or "")} if isinstance(before, dict) else None
    rules = [r for r in load(data_dir) if not _same(r, rule) and not (old and _same(r, old))]
    if len(rules) >= MAX_RULES:
        raise ValueError(T(f"Mehr als {MAX_RULES} Regeln gehen nicht. Lösch zuerst eine.",
                           f"More than {MAX_RULES} rules are not possible. Delete one first."))
    rules.append(rule)
    rules.sort(key=lambda r: (norm(r["spalte"]), norm(r["blatt"])))
    _write(data_dir, rules)
    return rules


def remove(data_dir: Path, column: str, sheet: str = "") -> list[dict]:
    gone = {"spalte": column, "blatt": sheet}
    rules = load(data_dir)
    kept = [r for r in rules if not _same(r, gone)]
    if len(kept) != len(rules):
        _write(data_dir, kept)
    return kept


def number(text) -> float | None:
    """„120“, „1,5“, „37.5“, „1'000“ oder „-3“ als Zahl. Komma und Punkt sind beide das Dezimalzeichen, mit beiden
    ist das erste das Tausenderzeichen (1.234,5 und 1,234.5)."""
    if isinstance(text, bool):
        return None
    if isinstance(text, (int, float)):
        return float(text) if math.isfinite(text) else None
    plain = re.sub(r"[\s'’ ]", "", str(text or "")).replace("−", "-")
    if "," in plain and "." in plain:
        thousands = "." if plain.rfind(",") > plain.rfind(".") else ","
        plain = plain.replace(thousands, "")
    plain = plain.replace(",", ".")
    if not re.fullmatch(r"[-+]?(\d+\.?\d*|\.\d+)", plain):
        return None
    return float(plain)


def day(text) -> str | None:
    """„2020-01-31“ (aus dem Datumsfeld) oder „31.1.2020“ als „2020-01-31“, sonst None."""
    plain = str(text or "").strip()
    match = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", plain) or re.fullmatch(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", plain)
    if not match:
        return None
    parts = [int(p) for p in match.groups()]
    year, month, mday = parts if len(match.group(1)) == 4 else parts[::-1]
    try:
        return date(year, month, mday).isoformat()
    except ValueError:
        return None


def _serial(iso: str) -> int:
    return (date.fromisoformat(iso) - _EPOCH).days


def _date_words(iso: str) -> str:
    return date.fromisoformat(iso).strftime("%d.%m.%Y")


def _items(values) -> list[str]:
    """Die erlaubten Werte: eine Liste, oder Text mit Strichpunkt („0,5; 1“), sonst mit Komma getrennt („m, w, d“)."""
    if isinstance(values, str):
        values = values.split(";") if ";" in values else values.split(",") if "," in values else values.splitlines()
    found = []
    for value in values or []:
        value = str(value).strip() if not isinstance(value, str) else value.strip()
        if value and value not in found:
            found.append(value[:MAX_TEXT])
    return found


def clean(rule: dict) -> dict:
    """Eine Regel aus dem Seitenfenster oder der Datei, geprüft und in fester Form. ValueError: was nicht stimmt."""
    if not isinstance(rule, dict):
        raise ValueError(T("Die Regel fehlt.", "The rule is missing."))
    column = " ".join(str(rule.get("spalte") or "").split())[:MAX_TEXT]
    if not column:
        raise ValueError(T("Zu welcher Spalte gehört die Regel? Markiere eine Zelle der Spalte.",
                           "Which column does the rule belong to? Select a cell of the column."))
    kind = str(rule.get("art") or "")
    if kind not in KINDS:
        raise ValueError(T("Diese Art gibt es nicht.", "This type does not exist."))
    found = {"spalte": column, "blatt": " ".join(str(rule.get("blatt") or "").split())[:MAX_TEXT],
             "pflicht": bool(rule.get("pflicht")), "art": kind}
    for key, word, label in (("min", "Von", "from"), ("max", "Bis", "to")):
        raw = rule.get(key)
        value = None if raw is None or str(raw).strip() == "" else number(raw)
        if value is None and raw is not None and str(raw).strip():
            raise ValueError(T(f"„{raw}“ unter {word} ist keine Zahl.", f"“{raw}” under “{label}” is not a number."))
        found[key] = value
    for key, word, label in (("nicht_vor", "nicht vor", "not before"), ("nicht_nach", "nicht nach", "not after")):
        raw = rule.get(key)
        value = day(raw) if raw else None
        if raw and value is None:
            raise ValueError(T(f"„{raw}“ unter {word} ist kein Datum (TT.MM.JJJJ).",
                               f"“{raw}” under “{label}” is not a date (DD.MM.YYYY)."))
        found[key] = value
    found["keine_zukunft"] = bool(rule.get("keine_zukunft"))
    found["werte"] = _items(rule.get("werte"))[:MAX_VALUES]
    found["einmal"] = bool(rule.get("einmal"))
    found["muster"] = str(rule.get("muster") or "").strip()[:60]
    other = " ".join(str(rule.get("spalte2") or "").split())[:MAX_TEXT]
    relation = str(rule.get("beziehung") or "ge")
    if other and relation not in RELATIONS:
        raise ValueError(T("Diesen Vergleich gibt es nicht.", "This comparison does not exist."))
    if other and norm(other) == norm(column):
        raise ValueError(T("Eine Spalte mit sich selbst zu vergleichen, sagt nichts.",
                           "Comparing a column with itself says nothing."))
    found["spalte2"], found["beziehung"] = other, relation if other else ""
    if found["min"] is not None and found["max"] is not None and found["min"] > found["max"]:
        raise ValueError(T(f"Von ({_fmt(found['min'])}) ist größer als bis ({_fmt(found['max'])}).",
                           f"From ({_fmt(found['min'])}) is greater than to ({_fmt(found['max'])})."))
    if found["nicht_vor"] and found["nicht_nach"] and found["nicht_vor"] > found["nicht_nach"]:
        before, after = _date_words(found["nicht_vor"]), _date_words(found["nicht_nach"])
        raise ValueError(T(f"„Nicht vor dem {before}“ liegt nach „nicht nach dem {after}“.",
                           f"“Not before {before}” is after “not after {after}”."))
    numbers = found["min"] is not None or found["max"] is not None
    dates = bool(found["nicht_vor"] or found["nicht_nach"] or found["keine_zukunft"])
    if numbers and kind in ("datum", "text"):
        raise ValueError(T(f"Von und bis für Zahlen passen nicht zur Art {KINDS[kind]}.",
                           f"From and to for numbers do not fit the type {_KINDS_EN[kind]}."))
    if dates and kind in ("zahl", "ganz", "text"):
        raise ValueError(T(f"Grenzen für Daten passen nicht zur Art {KINDS[kind]}.",
                           f"Limits for dates do not fit the type {_KINDS_EN[kind]}."))
    if not (found["pflicht"] or kind or numbers or dates or found["werte"] or found["einmal"] or found["muster"] or other):
        raise ValueError(T("Leg mindestens eine Bedingung fest, zum Beispiel Pflichtfeld oder von und bis.",
                           "Set at least one condition, for example required, or from and to."))
    return found


def _quote(text) -> str:
    return T(f"„{text}“", f"“{text}”")


def _listing(values: list[str], word: str = "oder") -> str:
    shown = [v if number(v) is not None else _quote(v) for v in values]
    word = T(word, {"oder": "or", "und": "and"}.get(word, word))
    return shown[0] if len(shown) == 1 else f"{', '.join(shown[:-1])} {word} {shown[-1]}"


def describe(rule: dict) -> str:
    """„Pflichtfeld, Zahl von 0 bis 120“ oder „Datum nicht vor dem 01.01.2020, nicht vor „Eintritt““."""
    if english():
        return _describe_en(rule)
    parts = ["Pflichtfeld"] if rule["pflicht"] else []
    low, high = rule["min"], rule["max"]
    base = KINDS[rule["art"]] if rule["art"] else ""
    if low is not None and high is not None:
        span = f"von {_fmt(low)} bis {_fmt(high)}"
    elif low is not None:
        span = f"ab {_fmt(low)}"
    elif high is not None:
        span = f"bis {_fmt(high)}"
    else:
        span = ""
    if base or span:
        parts.append(" ".join(p for p in (base or "Zahl", span) if p))
    if rule["nicht_vor"]:
        parts.append(f"nicht vor dem {_date_words(rule['nicht_vor'])}")
    if rule["nicht_nach"]:
        parts.append(f"nicht nach dem {_date_words(rule['nicht_nach'])}")
    if rule["keine_zukunft"]:
        parts.append("nicht in der Zukunft")
    if rule["werte"]:
        parts.append(f"nur {_listing(rule['werte'])}")
    if rule["einmal"]:
        parts.append("jeder Wert nur einmal")
    if rule["muster"]:
        parts.append(f"Muster „{rule['muster']}“")
    if rule["spalte2"]:
        words = ({"ge": "nicht vor", "le": "nicht nach"} if rule["art"] == "datum"
                 else {"ge": "nicht kleiner als", "le": "nicht größer als"})
        parts.append(f"{words[rule['beziehung']]} „{rule['spalte2']}“")
    return ", ".join(parts)


def _describe_en(rule: dict) -> str:
    """describe() auf Englisch: „Required, number from 0 to 120“."""
    parts = ["required"] if rule["pflicht"] else []
    low, high = rule["min"], rule["max"]
    base = _KINDS_EN[rule["art"]] if rule["art"] else ""
    if low is not None and high is not None:
        span = f"from {_fmt(low)} to {_fmt(high)}"
    elif low is not None:
        span = f"from {_fmt(low)}"
    elif high is not None:
        span = f"up to {_fmt(high)}"
    else:
        span = ""
    if base or span:
        parts.append(" ".join(p for p in (base or "number", span) if p))
    if rule["nicht_vor"]:
        parts.append(f"not before {_date_words(rule['nicht_vor'])}")
    if rule["nicht_nach"]:
        parts.append(f"not after {_date_words(rule['nicht_nach'])}")
    if rule["keine_zukunft"]:
        parts.append("not in the future")
    if rule["werte"]:
        parts.append(f"only {_listing(rule['werte'])}")
    if rule["einmal"]:
        parts.append("each value only once")
    if rule["muster"]:
        parts.append(f"pattern “{rule['muster']}”")
    if rule["spalte2"]:
        words = ({"ge": "not before", "le": "not after"} if rule["art"] == "datum"
                 else {"ge": "not less than", "le": "not greater than"})
        parts.append(f"{words[rule['beziehung']]} “{rule['spalte2']}”")
    text = ", ".join(parts)
    return text[:1].upper() + text[1:]


def public(rules: list[dict]) -> list[dict]:
    return [{**r, "text": describe(r)} for r in rules]


# --- Prüfen ---


class _Value:
    """Was in einer Zelle steht, wie eine Regel es braucht: Art, Zahl (bei Prozent in Prozent), angezeigter Text."""

    def __init__(self, sheet: Sheet, pos: Pos) -> None:
        cell = sheet.cells.get(pos)
        self.pos, self.cell = pos, cell
        self.shown = sheet.shown.get(pos)
        used = sheet.inputs.get(pos)
        self.raw = "" if cell is None else cell.text if cell.text is not None else cell.value
        self.display = "" if cell is None else cell.value.strip()
        self.percent = self.display.endswith("%")
        self.number: float | None = None
        if cell is None or cell.value.strip() == "":
            self.kind = "leer"
        elif isinstance(self.shown, Err):
            self.kind = "fehler"
        elif isinstance(self.shown, bool):
            self.kind = "wahr"
        elif isinstance(self.shown, float):
            self.kind, self.number = "zahl", self.shown * 100 if self.percent else self.shown
        elif isinstance(self.shown, str) and isinstance(used, float) and cell.text is None and serial_of(self.shown) is not None:
            self.kind, self.number = "datum", used
        else:
            self.kind = "text"

    @property
    def where(self) -> str:
        return address(self.pos)

    @property
    def quoted(self) -> str:
        """So, wie Excel es zeigt: Zahlen und Daten ohne, Text mit Anführungszeichen."""
        return self.display if self.kind in ("zahl", "datum") else _quote(self.raw)

    @property
    def as_text(self) -> str:
        """Womit Muster und erlaubte Werte vergleichen: Zahlen ohne Tausenderzeichen, Text genau, wie er drinsteht."""
        if self.kind == "zahl" and not self.percent:
            return number_text(self.shown)
        return self.raw if self.kind == "text" else self.display


def _pattern(text: str) -> re.Pattern:
    out = []
    for char in text:
        out.append({"#": r"\d", "@": r"[^\W\d_]", "?": ".", "*": ".*"}.get(char) or re.escape(char))
    return re.compile("".join(out), re.DOTALL)


_PLACE = {"#": ("eine Ziffer", "a digit"), "@": ("ein Buchstabe", "a letter"), "?": ("ein Zeichen", "a character")}


def _pattern_why(value: _Value, pattern: str) -> str:
    """Warum ein Wert nicht zum Muster passt: zu kurz, zu lang oder welches Zeichen. Mit * nur, dass er nicht passt."""
    text = value.as_text
    base = T(f"passt nicht zum Muster „{pattern}“", f"does not match the pattern “{pattern}”")
    if "*" in pattern:
        if _pattern(pattern).fullmatch(text.casefold()) or _pattern(pattern.casefold()).fullmatch(text.casefold()):
            return base + T(": nur anders geschrieben", ": only written differently")
        return base
    if _pattern(pattern.casefold()).fullmatch(text.casefold()):
        return base + T(": nur anders geschrieben (groß oder klein)", ": only written differently (upper or lower case)")
    if len(text) != len(pattern):
        short = len(text) < len(pattern)
        why = base + T(f": {'zu kurz' if short else 'zu lang'}, {len(text)} statt {len(pattern)} Zeichen",
                       f": too {'short' if short else 'long'}, {len(text)} instead of {len(pattern)} characters")
        if short and value.kind == "zahl" and set(pattern) == {"#"}:
            why += T(". Als Zahl hat Excel die Nullen vorne weggenommen; als Text gespeichert bleiben sie",
                     ". As a number, Excel dropped the leading zeros; stored as text, they stay")
        return why
    for i, (char, want) in enumerate(zip(text, pattern), 1):
        if not _pattern(want).fullmatch(char):
            needed = T(*_PLACE[want]) if want in _PLACE else _quote(want)
            return base + T(f": Zeichen {i} ist „{char}“, dort gehört {needed} hin",
                            f": character {i} is “{char}”, where {needed} belongs")
    return base


def _allowed(value: _Value, items: list[str], numeric: bool) -> tuple[bool, str]:
    """Ob der Wert in der Liste steht, und sonst, ob er nur anders geschrieben ist. numeric: Die Liste hat nur Zahlen,
    und in der Spalte stehen auch echte Zahlen: Dann ist „2“ als Text eine andere 2 (die Pivot-Tabelle zeigt sie zweimal)."""
    numbers = [n for n in (number(i) for i in items) if n is not None]
    text = value.as_text
    if value.kind == "zahl" and any(abs(value.number - n) <= 1e-9 * max(1.0, abs(n)) for n in numbers):
        return True, ""
    if text in items and not (numeric and value.kind == "text"):
        return True, ""
    if value.kind == "text":
        plain = text.strip()
        as_number = number(text)
        same = next((i for i in items if as_number is not None and number(i) is not None
                     and abs(as_number - number(i)) <= 1e-9 * max(1.0, abs(number(i)))), None)
        if same is not None:
            return False, (T("als Text gespeichert, nicht als Zahl", "stored as text, not as a number") if numeric
                           else T(f"anders geschrieben als „{same}“", f"written differently from “{same}”"))
        if text != plain and plain in items:
            return False, T("Leerzeichen am Rand", "spaces at the edge")
        near = next((i for i in items if i.casefold() == plain.casefold()), None)
        if near is not None:
            return False, T(f"anders geschrieben als „{near}“", f"written differently from “{near}”")
    return False, ""


class _Found:
    """Die Verstöße gegen eine Regel in einer Spalte, je Zelle: eine kurze Art (zum Zählen), der Wert und was mit ihm
    nicht stimmt („über 120“). Derselbe Wert mit demselben Grund steht später in einem Satz mit allen seinen Zellen."""

    def __init__(self) -> None:
        self.cells: dict[Pos, tuple[str, str, str]] = {}

    def add(self, pos: Pos, short: str, value: str, phrase: str) -> None:
        if pos not in self.cells:
            self.cells[pos] = (short, value, phrase)


def _cell_verdict(value: _Value, rule: dict, found: _Found, today: int, numeric: bool = False) -> None:  # noqa: C901 - eine Bedingung je Abschnitt
    kind, pos = value.kind, value.pos
    wanted = rule["art"]
    conditions = bool(wanted or rule["min"] is not None or rule["max"] is not None or rule["nicht_vor"] or rule["nicht_nach"]
                      or rule["keine_zukunft"] or rule["werte"] or rule["muster"])
    if kind == "leer":
        if rule["pflicht"]:
            formula = value.cell is not None and value.cell.formula is not None
            found.add(pos, T("leer", "empty"), "", T(f"leer{' (die Formel ergibt nichts)' if formula else ''}",
                                                    f"empty{' (the formula returns nothing)' if formula else ''}"))
        return
    if kind == "text" and value.raw.strip().casefold() in _PLACEHOLDERS and (rule["pflicht"] or wanted in ("zahl", "ganz", "datum")):
        found.add(pos, T("leer", "empty"), value.quoted, T("steht statt eines Werts da", "is there instead of a value"))
        return
    if kind == "fehler":
        if conditions:
            found.add(pos, T("Fehlerwert", "error value"), value.display, T("ist ein Fehlerwert", "is an error value"))
        return
    if wanted in ("zahl", "ganz") and kind != "zahl":
        if kind == "datum":
            found.add(pos, T("Datum statt Zahl", "date instead of number"), value.display,
                      T("ist ein Datum, keine Zahl", "is a date, not a number"))
        elif kind == "text" and number(value.raw) is not None:
            found.add(pos, T("Zahl als Text", "number as text"), value.quoted,
                      T("ist als Text gespeichert, nicht als Zahl", "is stored as text, not as a number"))
        else:
            found.add(pos, T("keine Zahl", "not a number"), value.quoted, T("ist keine Zahl", "is not a number"))
        return
    if wanted == "ganz" and abs(value.number - round(value.number)) > 1e-9:
        found.add(pos, T("keine ganze Zahl", "not a whole number"), value.quoted,
                  T("ist keine ganze Zahl", "is not a whole number"))
        return
    if wanted == "datum" and kind != "datum":
        if kind == "zahl" and not value.percent and 1 <= value.number <= 2958465:
            shown = date.fromordinal(_EPOCH.toordinal() + int(value.number)).strftime("%d.%m.%Y")
            found.add(pos, T("kein Datum", "not a date"), value.quoted,
                      T(f"ist eine Zahl, kein Datum (als Datum formatiert wäre es der {shown})",
                        f"is a number, not a date (formatted as a date, it would be {shown})"))
        elif kind == "text" and serial_of(value.raw) is not None:
            found.add(pos, T("Datum als Text", "date as text"), value.quoted,
                      T("ist als Text gespeichert, Excel rechnet damit nicht wie mit einem Datum",
                        "is stored as text; Excel does not calculate with it as with a date"))
        else:
            found.add(pos, T("kein Datum", "not a date"), value.quoted, T("ist kein Datum", "is not a date"))
        return
    if wanted == "text" and kind in ("zahl", "datum"):
        what = T("eine Zahl", "a number") if kind == "zahl" else T("ein Datum", "a date")
        found.add(pos, T("kein Text", "not text"), value.quoted, T(f"ist {what}, kein Text", f"is {what}, not text"))
        return
    if kind == "zahl":
        if rule["min"] is not None and value.number < rule["min"] - 1e-9:
            low = _fmt(rule["min"])
            found.add(pos, T(f"unter {low}", f"below {low}"), value.quoted, T(f"ist unter {low}", f"is below {low}"))
            return
        if rule["max"] is not None and value.number > rule["max"] + 1e-9:
            high = _fmt(rule["max"])
            found.add(pos, T(f"über {high}", f"above {high}"), value.quoted, T(f"ist über {high}", f"is above {high}"))
            return
    if kind == "datum":
        on = math.floor(value.number + 1e-9)
        if rule["nicht_vor"] and on < _serial(rule["nicht_vor"]):
            words = _date_words(rule["nicht_vor"])
            found.add(pos, T(f"vor dem {words}", f"before {words}"), value.display,
                      T(f"liegt vor dem {words}", f"is before {words}"))
            return
        if rule["nicht_nach"] and on > _serial(rule["nicht_nach"]):
            words = _date_words(rule["nicht_nach"])
            found.add(pos, T(f"nach dem {words}", f"after {words}"), value.display,
                      T(f"liegt nach dem {words}", f"is after {words}"))
            return
        if rule["keine_zukunft"] and on > today:
            found.add(pos, T("in der Zukunft", "in the future"), value.display,
                      T("liegt in der Zukunft", "is in the future"))
            return
    if rule["werte"]:
        ok, why = _allowed(value, rule["werte"], numeric)
        if not ok:
            found.add(pos, T("nicht in der Liste", "not in the list"), value.quoted,
                      T(f"steht nicht in der Liste{f' ({why})' if why else ''}", f"is not in the list{f' ({why})' if why else ''}"))
            return
    if rule["muster"] and kind != "wahr" and not _pattern(rule["muster"]).fullmatch(value.as_text):
        found.add(pos, T("passt nicht zum Muster", "does not match the pattern"), value.quoted,
                  _pattern_why(value, rule["muster"]))


def _compare(sheet: Sheet, rule: dict, col: int, other: int, rows: list[int], found: _Found) -> None:
    """Die Spalte mit einer anderen derselben Zeile: Austritt nicht vor Eintritt, Diastolisch nicht über Systolisch."""
    plan = sheet.plan
    for r in rows:
        a, b = _Value(sheet, (r, col)), _Value(sheet, (r, other))
        if a.kind not in ("zahl", "datum") or b.kind not in ("zahl", "datum") or a.percent != b.percent:
            continue
        if a.pos in found.cells:
            continue
        dates = a.kind == b.kind == "datum"
        low = rule["beziehung"] == "ge" and a.number < b.number - 1e-9
        high = rule["beziehung"] == "le" and a.number > b.number + 1e-9
        if not (low or high):
            continue
        head = plan.headers[other]
        if dates:
            short = T(f"{'vor' if low else 'nach'} {head}", f"{'before' if low else 'after'} {head}")
            phrase = T(f"liegt {'vor' if low else 'nach'} {head} {b.display} ({b.where})",
                       f"is {'before' if low else 'after'} {head} {b.display} ({b.where})")
        else:
            short = T(f"{'kleiner' if low else 'größer'} als {head}", f"{'less' if low else 'greater'} than {head}")
            phrase = T(f"ist {'kleiner' if low else 'größer'} als {head} {b.display} ({b.where})",
                       f"is {'less' if low else 'greater'} than {head} {b.display} ({b.where})")
        found.add(a.pos, short, f"{plan.headers[col]} {a.display}", phrase)


def _key(value: _Value):
    if value.kind in ("zahl", "datum"):
        return value.kind, round(value.shown if value.kind == "zahl" else value.number, 9)
    if value.kind == "text" and value.raw.strip().casefold() not in _PLACEHOLDERS:
        return "text", value.raw.strip().casefold()
    return None


def _twice(sheet: Sheet, col: int, rows: list[int]) -> list[list[_Value]]:
    """Werte, die mehrmals vorkommen (Text ohne Unterschied in Groß und Klein und Leerzeichen am Rand)."""
    groups: dict = {}
    for r in rows:
        value = _Value(sheet, (r, col))
        key = _key(value)
        if key is not None:
            groups.setdefault(key, []).append(value)
    return [g for g in groups.values() if len(g) > 1]


def _areas(cells: list[Pos]) -> str:
    """Zum Markieren: C5,C14:C16,F3 (aufeinanderfolgende Zeilen zusammen), höchstens MAX_AREAS Bereiche."""
    areas = []
    for pos in sorted(cells, key=lambda p: (p[1], p[0])):
        if areas and areas[-1][1] == pos[1] and areas[-1][2] == pos[0] - 1:
            areas[-1][2] = pos[0]
        else:
            areas.append([pos[0], pos[1], pos[0]])
    out = []
    for top, col, bottom in areas[:MAX_AREAS]:
        letters = column_letters(col)
        out.append(f"{letters}{top}" if top == bottom else f"{letters}{top}:{letters}{bottom}")
    return ",".join(out)


def _rows(sheet: Sheet) -> list[int]:
    """Die Datenzeilen, ohne ganz leere: Die sind kein Fall, dem etwas fehlt."""
    plan = sheet.plan
    return [r for r in _data_rows(sheet) if any(sheet.cells.get((r, c)) is not None and sheet.cells[(r, c)].value.strip()
                                                 for c in plan.headers)]


def _rules_here(sheet: Sheet, rules: list[dict]) -> list[tuple[int, dict]]:
    return _matching(sheet.name or "", sheet.plan.headers, rules)


def _matching(name: str, headers: dict[int, str], rules: list[dict]) -> list[tuple[int, dict]]:
    """Je Spalte der Tabelle die Regel, die hier gilt: eine für dieses Blatt vor einer für alle Blätter."""
    here = norm(name)
    found = []
    for col, header in sorted(headers.items()):
        wanted = norm(header)
        fitting = [r for r in rules if norm(r["spalte"]) == wanted and norm(r["blatt"]) in ("", here)]
        fitting.sort(key=lambda r: r["blatt"] == "")
        if fitting:
            found.append((col, fitting[0]))
    return found


def used(cells_of: dict[str, dict], rules: list[dict]) -> list[dict]:
    """Für den Prüfbericht: die Regeln, die in dieser Mappe gelten, mit den Blättern, auf denen ihre Spalte steht."""
    if not rules:
        return []
    sheets: dict[int, list[str]] = {}
    for name, cells in cells_of.items():
        plan = layout(Book(name, cells))
        if plan.header_row is None:
            continue
        for _, rule in _matching(name, plan.headers, rules):
            sheets.setdefault(id(rule), []).append(name)
    return [{**r, "text": describe(r), "blaetter": sheets[id(r)]} for r in rules if id(r) in sheets]


def checks(sheet: Sheet, rules: list[dict], earlier: list[Finding], today: date | None = None) -> tuple[list[Finding], set[str]]:
    """Die Verstöße gegen die eigenen Regeln, ein Hinweis je Spalte, und die Schlüssel der allgemeinen Hinweise, die
    eine Regel genauer sagt (earlier: die Hinweise, die es zu diesem Blatt schon gibt)."""
    plan = sheet.plan
    if not rules or plan.header_row is None:
        return [], set()
    matched = _rules_here(sheet, rules)
    if not matched:
        return [], set()
    rows = _rows(sheet)
    names = {norm(h): c for c, h in plan.headers.items()}
    now = ((today or date.today()) - _EPOCH).days
    found, covered = [], set()
    for col, rule in matched:
        result = _Found()
        values = [_Value(sheet, (r, col)) for r in rows]
        numeric = bool(rule["werte"]) and all(number(i) is not None for i in rule["werte"]) and any(v.kind == "zahl" for v in values)
        for value in values:
            _cell_verdict(value, rule, result, now, numeric)
        other = names.get(norm(rule["spalte2"])) if rule["spalte2"] else None
        if other is not None and other != col:
            _compare(sheet, rule, col, other, rows, result)
        doubles = _twice(sheet, col, rows) if rule["einmal"] else []
        codes = _code_cells(sheet, col, result, rows)
        here = _covers(sheet, rule, col, other, result, earlier, codes)
        covered |= here
        if result.cells or doubles:
            item = _finding(sheet, rule, col, result, doubles, codes)
            # Was ein Ende vor dem Anfang in seiner Zeile anrichtet (die negative Liegedauer), sagt der Hinweis zur Regel
            # weiter, statt es mit dem allgemeinen Hinweis wegzulassen.
            effects = sheet.memo.get("auswertung.effect") or {}
            for before in earlier:
                effect = effects.get(before.key)
                if before.key in here and before.key.startswith("order:") and effect and effect in before.text:
                    item.text += " " + before.text[before.text.index(effect):]
            found.append(item)
    return found, covered


def _covers(sheet: Sheet, rule: dict, col: int, other: int | None, result: _Found, earlier: list[Finding],
            codes: dict[float, list[Pos]]) -> set[str]:
    """Was die Regel dieser Spalte genauer sagt als die allgemeinen Hinweise (auswertung.py, analyse.py)."""
    plan = sheet.plan
    numbers = rule["min"] is not None or rule["max"] is not None
    dates = bool(rule["nicht_nach"] or rule["keine_zukunft"])
    placeholders = rule["pflicht"] or rule["art"] in ("zahl", "ganz", "datum")
    covered = set()
    for item in earlier:
        kind, _, where = item.key.partition(":")
        if not re.fullmatch(r"[A-Z]{1,3}\d+", where):
            continue
        letters = re.match(r"[A-Z]+", where).group(0)
        here = letters == column_letters(col)
        if kind == "impossible" and here and numbers:
            covered.add(item.key)
        elif kind == "id" and here and rule["einmal"]:
            covered.add(item.key)
        elif kind == "future" and here and dates:
            covered.add(item.key)
        elif kind == "placeholder" and here and placeholders:
            covered.add(item.key)
        elif kind == "outlier" and any(address(p) == where for p in result.cells):
            covered.add(item.key)
        elif kind == "code" and here and any(address(p) == where for cells in codes.values() for p in cells):
            covered.add(item.key)
        elif kind == "order" and other is not None and rule["beziehung"] == "ge" and here:
            # „Austritt liegt vor Eintritt“: genau das Paar dieser Regel.
            if item.text.startswith(order_start(plan.headers[col], plan.headers[other])):
                covered.add(item.key)
        elif kind == "order" and other is not None and rule["beziehung"] == "le" and letters == column_letters(other):
            if item.text.startswith(order_start(plan.headers[other], plan.headers[col])):
                covered.add(item.key)
    return covered


def _code_cells(sheet: Sheet, col: int, result: _Found, rows: list[int]) -> dict[float, list[Pos]]:
    """Codes wie 999 für „fehlt“ (auswertung.codes_in), wenn jede Zelle mit dem Code gegen die Regel verstößt: Zahl →
    Zellen. Dann sagt der Hinweis der Regel es mit; sonst bleibt der allgemeine Hinweis zum Code."""
    numbers = {(r, col): sheet.shown[(r, col)] for r in rows if isinstance(sheet.shown.get((r, col)), float)}
    out: dict[float, list[Pos]] = {}
    for code in sorted(codes_in(list(numbers.values()))):
        where = sorted(p for p, v in numbers.items() if v == code)
        if all(p in result.cells for p in where):
            out[code] = where
    return out


def _cells_text(cells: list[Pos]) -> str:
    """„E6“, „E6 und E20“, „I9, I10, I12 und I13“ oder „I9, I10, I12 und 9 weiteren Zellen“."""
    names = [address(p) for p in cells]
    if len(names) == 1:
        return names[0]
    if len(names) <= 4:
        return T(f"{', '.join(names[:-1])} und {names[-1]}", f"{', '.join(names[:-1])} and {names[-1]}")
    return T(f"{', '.join(names[:3])} und {len(names) - 3} weiteren Zellen",
             f"{', '.join(names[:3])} and {len(names) - 3} more cells")


def _sentence(value: str, phrase: str, cells: list[Pos]) -> str:
    """„999 in E6 und E20 ist über 110“, „E5 ist leer“ oder „E5 und E9 sind leer“."""
    if not value:
        return f"{_cells_text(cells)} {T('ist', 'is') if len(cells) == 1 else T('sind', 'are')} {phrase}"
    return f"{value} in {_cells_text(cells)} {phrase}"


def _finding(sheet: Sheet, rule: dict, col: int, result: _Found, doubles: list[list[_Value]],
             codes: dict[float, list[Pos]]) -> Finding:
    header = sheet.plan.headers[col]
    where = _quote(header) + (T(f" auf „{rule['blatt']}“", f" on “{rule['blatt']}”") if rule["blatt"] else "")
    said = describe(rule)
    parts = [T(f"Deine Regel für {where} ({said}):", f"Your rule for {where} ({said[:1].lower()}{said[1:]}):")]
    cells = sorted(result.cells, key=lambda p: (p[0], p[1]))
    if cells:
        counts = Counter(short for short, _, _ in result.cells.values())
        count = T(f"{len(cells)} {'Zelle passt' if len(cells) == 1 else 'Zellen passen'} nicht",
                  f"{len(cells)} {'cell does' if len(cells) == 1 else 'cells do'} not fit")
        if len(counts) > 1:
            count += " (" + ", ".join(f"{n} {short}" for short, n in counts.most_common()) + ")"
        # Derselbe Wert mit demselben Grund in einem Satz; von jeder Art zuerst ein Satz, dann der Reihe nach.
        groups: dict[tuple, list[Pos]] = {}
        for p in cells:
            groups.setdefault(result.cells[p], []).append(p)
        order = []
        for short, _ in counts.most_common():
            order.append(next(g for g in groups if g[0] == short))
        order += [g for g in groups if g not in order]
        listed = sorted(order[:MAX_LISTED], key=lambda g: groups[g][0])
        texts = [_sentence(value, phrase, groups[(short, value, phrase)]) for short, value, phrase in listed]
        more = sum(len(groups[g]) for g in order[MAX_LISTED:])
        parts.append(f"{count}: {'; '.join(texts)}." + (T(f" Dazu {more} {'weitere Zelle' if more == 1 else 'weitere Zellen'}.",
                                                            f" Plus {more} more {'cell' if more == 1 else 'cells'}.")
                                                          if more else ""))
        for code, where_ in codes.items():
            parts.append(T(f"{_fmt(code)} steht {len(where_)}-mal da, weit weg von den übrigen Werten: wohl ein Code für "
                           f"„fehlt“, und MITTELWERT und SUMME rechnen damit.",
                           f"{_fmt(code)} appears {len(where_)} times, far from the other values: probably a code for "
                           f"“missing”, and AVERAGE and SUM calculate with it."))
    if doubles:
        shown = []
        for group in doubles[:3]:
            places = [v.where for v in group]
            listed = (", ".join(places[:-1][:3]) + T(f" und {places[-1]}", f" and {places[-1]}") if len(places) <= 4
                      else ", ".join(places[:4]) + T(" und weitere", " and more"))
            shown.append(f"{group[0].quoted} in {listed}")
        more = T(f" und {len(doubles) - 3} weitere", f" and {len(doubles) - 3} more") if len(doubles) > 3 else ""
        if english():
            word = "One value appears" if len(doubles) == 1 else f"{len(doubles)} values appear"
            parts.append(f"{word} more than once: {'; '.join(shown)}{more}.")
        else:
            word = "Ein Wert kommt" if len(doubles) == 1 else f"{len(doubles)} Werte kommen"
            parts.append(f"{word} mehrmals vor: {'; '.join(shown)}{more}.")
    marked = list(result.cells) + [v.pos for g in doubles for v in g]
    tip = T("Korrigier die Zellen. Stimmt ein Wert doch, ändere die Regel (unter „Eigene Regeln je Spalte“).",
            "Correct the cells. If a value is right after all, change the rule (under “Your own rules per column”).")
    if doubles:
        tip = T("Ist derselbe Fall doppelt erfasst? Dann zählt er doppelt. ",
                "Is the same case entered twice? Then it counts twice. ") + tip
    # Der Schlüssel ist die Überschrift: Er bleibt, wenn „Einsetzen“ einen der Verstöße behebt und andere bleiben.
    return Finding(f"regel:{address((sheet.plan.header_row, col))}", " ".join(parts), "fehler", tip,
                   action={"kind": "rule", "spalte": rule["spalte"], "blatt": rule["blatt"]}, cell=_areas(marked))


# --- Für das Formular: was jetzt in der Spalte steht ---


def find(rules: list[dict], column: str, sheet: str = "") -> dict | None:
    return next((r for r in rules if _same(r, {"spalte": column, "blatt": sheet})), None)


def column(sheet: Sheet, pos: Pos, rules: list[dict], only: str | None = None) -> dict:
    """Die Spalte der markierten Zelle: ihr Name, was darin steht (damit man die Grenzen sieht), die übrigen
    Spalten (zum Vergleichen) und die Regel, die schon gilt (only: die Regel für genau dieses Blatt, "" für alle)."""
    plan = sheet.plan
    r, c = pos
    if plan.header_row is None or c not in plan.headers or not plan.header_row <= r <= plan.last_row + 1:
        return {"state": "none", "text": T("Markiere in Excel eine Zelle in einer Spalte der Tabelle (auf oder unter ihrer "
                                           "Überschrift), dann lege ich die Regel für diese Spalte an.",
                                           "Select a cell in a column of the table in Excel (on or below its heading), "
                                           "then I set up the rule for this column.")}
    header = plan.headers[c]
    rows = _rows(sheet)
    values = [_Value(sheet, (row, c)) for row in rows]
    kinds = Counter(v.kind for v in values)
    numbers = [v for v in values if v.kind == "zahl"]
    dates = [v for v in values if v.kind == "datum"]
    texts = [v for v in values if v.kind == "text"]
    said = [T(f"{len(rows)} {'Zeile' if len(rows) == 1 else 'Zeilen'}", f"{len(rows)} {'row' if len(rows) == 1 else 'rows'}")]
    if numbers:
        low, high = min(numbers, key=lambda v: v.number), max(numbers, key=lambda v: v.number)
        said.append(T(f"{len(numbers)} {'Zahl' if len(numbers) == 1 else 'Zahlen'}",
                      f"{len(numbers)} {'number' if len(numbers) == 1 else 'numbers'}") +
                    (T(f" von {low.display} bis {high.display}", f" from {low.display} to {high.display}")
                     if len(numbers) > 1 else f" ({low.display})"))
    if dates:
        low, high = min(dates, key=lambda v: v.number), max(dates, key=lambda v: v.number)
        said.append(T(f"{len(dates)} Datum", f"{len(dates)} date") if len(dates) == 1 else
                    T(f"{len(dates)} Daten vom {low.display} bis {high.display}",
                      f"{len(dates)} dates from {low.display} to {high.display}"))
    if texts:
        counts = Counter(v.raw for v in texts)
        listed = ", ".join(_quote(t) + (T(f" {n}-mal", f" {n} times") if n > 1 else "") for t, n in counts.most_common(8))
        said.append(T(f"{len(texts)} Text ({len(counts)} verschieden: {listed}{' und weitere' if len(counts) > 8 else ''})",
                      f"{len(texts)} text ({len(counts)} different: {listed}{' and more' if len(counts) > 8 else ''})"))
    if kinds["leer"]:
        said.append(T(f"{kinds['leer']} leer", f"{kinds['leer']} empty"))
    if kinds["fehler"]:
        said.append(T(f"{kinds['fehler']} Fehlerwert{'e' if kinds['fehler'] > 1 else ''}",
                      f"{kinds['fehler']} error value{'s' if kinds['fehler'] > 1 else ''}"))
    text = said[0] + (": " + ", ".join(said[1:]) if len(said) > 1 else "") + "."
    doubles = _twice(sheet, c, rows)
    filled = len(numbers) + len(dates) + len(texts)
    if filled > 1 and not doubles:
        text += T(" Jeder Wert kommt nur einmal vor.", " Each value appears only once.")
    elif filled > 1 and filled - sum(len(g) - 1 for g in doubles) >= 0.9 * filled:
        # Fast jeder Wert nur einmal, wie bei Nummern: die wenigen doppelten sind wichtig. Bei Gruppen (m, w) nicht.
        text += T(f" {len(doubles)} {'Wert kommt' if len(doubles) == 1 else 'Werte kommen'} mehrmals vor.",
                  f" {len(doubles)} {'value appears' if len(doubles) == 1 else 'values appear'} more than once.")
    if any(v.percent for v in numbers):
        text += T(" Die Zahlen sind als Prozent formatiert: Gib von und bis in Prozent an, wie Excel sie zeigt (0 bis 100).",
                  " The numbers are formatted as percent: give from and to in percent, as Excel shows them (0 to 100).")
    guess = ""
    if filled and len(dates) >= 0.8 * filled:
        guess = "datum"
    elif filled and len(numbers) >= 0.8 * filled:
        guess = "ganz" if all(abs(v.number - round(v.number)) <= 1e-9 for v in numbers) else "zahl"
    if only is not None:
        rule = find(rules, header, only)
    else:
        rule = next((x for col, x in _rules_here(sheet, rules) if col == c), None)
    return {"state": "ok", "spalte": header, "blatt": sheet.name or "", "jetzt": text, "art": guess,
            "spalten": [h for col, h in sorted(plan.headers.items()) if col != c and h.strip()],
            "regel": {**rule, "text": describe(rule)} if rule else None}
