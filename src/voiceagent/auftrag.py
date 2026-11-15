"""Aufträge für Excel verstehen: aus „Summe der Sitze pro Partei“ wird eine feste Form.

Einfache Sätze versteht Gegenprobe selbst, ohne KI: Rechenwörter (Summe, Anzahl, Durchschnitt,
Minimum, Maximum), „pro“, „je“ oder „nach“ vor den Spalten, nach denen aufgeteilt wird, und
„in den Spalten“ für eine Kreuztabelle. Die Spaltennamen dürfen ungefähr stimmen („Parteien“
für die Spalte „Partei“). Rechnen und Schreiben macht danach Gegenprobe selbst, alles ohne KI und offline.
"""

from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field

from .sprache import T


# Rechenarten, wie sie in Excel heißen (Pivot: „Summe von Sitze“).
FUNCTIONS = {"summe": "Summe", "anzahl": "Anzahl", "mittelwert": "Mittelwert", "min": "Minimum", "max": "Maximum"}
# So heißen sie in der englischen Excel („Sum of Sitze“).
ENGLISH = {"summe": "Sum", "anzahl": "Count", "mittelwert": "Average", "min": "Min", "max": "Max"}


def function_name(funktion: str) -> str:
    """Die Rechenart, wie Excel sie in der Sprache der Anfrage nennt: „Summe“ oder „Sum“."""
    return T(FUNCTIONS[funktion], ENGLISH[funktion])

_AGG = {
    "summe": "summe", "summen": "summe", "summiere": "summe", "summieren": "summe", "summiert": "summe",
    "gesamtsumme": "summe", "addiere": "summe", "addiert": "summe", "sum": "summe",
    "anzahl": "anzahl", "anzahlen": "anzahl", "zaehle": "anzahl", "zaehlen": "anzahl", "zaehl": "anzahl",
    "gezaehlt": "anzahl", "count": "anzahl", "haeufigkeit": "anzahl", "haeufigkeiten": "anzahl",
    "mittelwert": "mittelwert", "mittelwerte": "mittelwert", "durchschnitt": "mittelwert",
    "durchschnitte": "mittelwert", "durchschnittlich": "mittelwert", "durchschnittliche": "mittelwert",
    "durchschnittlichen": "mittelwert", "durchschnittlicher": "mittelwert", "schnitt": "mittelwert",
    "mittel": "mittelwert", "average": "mittelwert", "mean": "mittelwert", "avg": "mittelwert",
    "minimum": "min", "minima": "min", "min": "min", "kleinste": "min", "kleinsten": "min", "kleinster": "min",
    "niedrigste": "min", "niedrigsten": "min", "tiefste": "min", "tiefsten": "min", "minimal": "min",
    "minimale": "min", "minimalen": "min",
    "maximum": "max", "maxima": "max", "max": "max", "groesste": "max", "groessten": "max", "groesster": "max",
    "hoechste": "max", "hoechsten": "max", "hoechster": "max", "maximal": "max", "maximale": "max",
    "maximalen": "max",
}
_COUNT_PHRASES = ("wie viele", "wieviele", "wie viel", "wieviel", "wie oft")
# Danach kommen die Spalten, nach denen aufgeteilt wird.
_GROUP = {
    "pro", "je", "nach", "per", "by", "jede", "jeden", "jedes", "jeder", "gruppiert", "aufgeteilt", "getrennt",
    "aufgeschluesselt", "unterteilt", "verteilt", "zeilen",
}
# „in den Spalten das Jahr“, „Jahre als Spalten“, „quer nach Jahr“: eine Kreuztabelle.
_ACROSS = {"spalten", "spaltenweise", "quer", "nebeneinander", "spaltenkoepfe"}
# „Partei gegen Kanton“: links die Zeilen, rechts die Spalten einer Kreuztabelle.
_VERSUS = {"gegen", "gegenueber", "versus", "vs", "mal", "x", "kreuz"}
_SORT = {"sortiert", "sortiere", "sortieren", "geordnet", "ordne", "ordnen", "rangliste", "ranking"}
_DESC = ("absteigend", "abwaerts", "groesste zuerst", "groessten zuerst", "hoechste zuerst", "hoechsten zuerst",
         "meiste zuerst", "meisten zuerst", "rangliste", "ranking", "von gross nach klein", "von oben nach unten")
_ASC = ("aufsteigend", "aufwaerts", "kleinste zuerst", "kleinsten zuerst", "wenigste zuerst", "wenigsten zuerst",
        "von klein nach gross")
# Wörter, die zwischen Rechenart und Spalte stehen dürfen („Summe der Sitze“).
_LINK = {"der", "des", "von", "vom", "aller", "an", "die", "das", "den", "dem"}
_UMLAUTS = str.maketrans({"ä": "ae", "ö": "oe", "ü": "ue", "ß": "ss"})


@dataclass
class Wert:
    # summe | anzahl | mittelwert | min | max, None: je nach Spalte (Zahlen: Summe, sonst Anzahl)
    funktion: str | None
    # Spalte mit den Werten, None: Anzahl der Zeilen
    spalte: int | None


@dataclass
class Auftrag:
    """Eine Pivot-Übersicht: aufteilen nach zeilen (und spalten), zusammenfassen werte."""

    zeilen: list[int] = field(default_factory=list)
    spalte: int | None = None
    werte: list[Wert] = field(default_factory=list)
    # "" (nach Namen) | "absteigend" | "aufsteigend" (nach dem ersten Wert)
    sortierung: str = ""


@dataclass
class Verstanden:
    auftrag: Auftrag | None
    # Warum nicht, wenn auftrag None ist.
    text: str = ""


def normal(text: str) -> str:
    """Klein, ohne Umlaute und Satzzeichen: „Pivot-Tabelle, größte zuerst!“ → „pivot tabelle groesste zuerst“."""
    text = text.casefold().translate(_UMLAUTS)
    text = re.sub(r"[^\w%]+", " ", text)
    return " ".join(text.split())


def _stem(word: str) -> str:
    for ending in ("innen", "en", "er", "es", "e", "n", "s"):
        if word.endswith(ending) and len(word) - len(ending) >= 3:
            return word[: -len(ending)]
    return word


def _same(word: str, other: str) -> bool:
    """Ungefähr dasselbe Wort: Mehrzahl, Fälle, kleine Tippfehler („Parteien“, „Patei“ → „Partei“)."""
    if word == other:
        return True
    if word.isdigit() or other.isdigit():
        return False
    if _stem(word) == _stem(other):
        return True
    short, long = sorted((word, other), key=len)
    if len(short) >= 4 and long.startswith(short) and len(long) - len(short) <= 3:
        return True
    return len(short) >= 5 and difflib.SequenceMatcher(None, word, other).ratio() >= 0.85


def _matches(words: list[str], headers: dict[int, str]) -> dict[int, tuple[int, int]]:
    """Wo im Satz welche Spalte genannt ist: Position → (Spalte, Anzahl Wörter). Längere Namen zuerst."""
    names = sorted(((col, normal(name).split()) for col, name in headers.items()), key=lambda item: -len(item[1]))
    taken: set[int] = set()
    found: dict[int, tuple[int, int]] = {}
    for col, parts in names:
        if not parts:
            continue
        for start in range(len(words) - len(parts) + 1):
            span = range(start, start + len(parts))
            if any(i in taken for i in span):
                continue
            if all(_same(words[start + k], parts[k]) for k in range(len(parts))):
                follower = words[start + len(parts)] if start + len(parts) < len(words) else ""
                if len(parts) == 1 and words[start] in _AGG and follower in _LINK:
                    # „Summe der Sitze“: Hier ist „Summe“ die Rechenart, auch wenn eine Spalte so heißt.
                    continue
                taken.update(span)
                found[start] = (col, len(parts))
    return found


def parse(text: str, headers: dict[int, str], numeric: set[int] | None = None) -> Verstanden:
    """Den Satz ohne KI verstehen. headers: Spalte → Name, wie Gegenprobe die Kopfzeile liest.

    numeric: Spalten mit Zahlen, nur für ein passendes Beispiel, wenn etwas fehlt."""
    plain = normal(text)
    if not plain:
        return Verstanden(None, T("Schreib, was ich machen soll, zum Beispiel: Summe der Sitze pro Partei.",
                                  "Tell me what to do, for example: sum of seats per party."))
    order = ""
    if any(f" {p} " in f" {plain} " for p in _DESC):
        order = "absteigend"
    elif any(f" {p} " in f" {plain} " for p in _ASC):
        order = "aufsteigend"
    for phrase in ("groesste zuerst", "groessten zuerst", "hoechste zuerst", "hoechsten zuerst", "kleinste zuerst",
                   "kleinsten zuerst", "meiste zuerst", "meisten zuerst", "wenigste zuerst", "wenigsten zuerst"):
        # „größte zuerst“ heißt sortieren, nicht Maximum.
        plain = plain.replace(phrase, "zuerst")
    for phrase in _COUNT_PHRASES:
        plain = plain.replace(phrase, "anzahl")
    words = plain.split()
    found = _matches(words, headers)
    auftrag = Auftrag(sortierung=order)
    mode, pending, last = "wert", [], []
    loose: list[Wert] = []
    i = 0
    while i < len(words):
        if i in found:
            col, length = found[i]
            if mode == "gruppe":
                if col not in auftrag.zeilen:
                    auftrag.zeilen.append(col)
            elif mode == "quer":
                auftrag.spalte = col
                mode = "gruppe"
            elif mode == "sortieren":
                auftrag.sortierung = auftrag.sortierung or "absteigend"
                mode = "wert"
                if not any(w.spalte == col for w in auftrag.werte):
                    auftrag.werte.append(Wert(None, col))
            else:
                functions = pending or last
                if functions:
                    auftrag.werte += [Wert(f, col) for f in functions]
                    last, pending = functions, []
                else:
                    item = Wert(None, col)
                    auftrag.werte.append(item)
                    loose.append(item)
            i += length
            continue
        word = words[i]
        if word in _AGG:
            function = _AGG[word]
            if loose:
                # „Sitze summiert pro Partei“: Die Rechenart kommt nach der Spalte.
                for item in loose:
                    item.funktion = function
                loose, last = [], [function]
            else:
                if mode != "wert":
                    mode, last = "wert", []
                if function not in pending:
                    pending.append(function)
        elif word in _SORT:
            mode = "sortieren"
            if i + 1 < len(words) and words[i + 1] == "nach":
                i += 1
        elif word in _ACROSS:
            before = i - 1
            while before >= 0 and words[before] in _LINK | {"als", "in", "ueber", "oben"}:
                before -= 1
            ahead = i + 1
            while ahead < len(words) and words[ahead] in _LINK | {"nach", "mit", "fuer"}:
                ahead += 1
            if ahead in found:
                mode = "quer"
                i = ahead
                continue
            if auftrag.zeilen and before >= 0 and any(p <= before < p + n for p, (_, n) in found.items()):
                # „Jahre in den Spalten“: Die zuletzt genannte Spalte wandert nach oben.
                auftrag.spalte = auftrag.zeilen.pop()
        elif word in _VERSUS and i + 1 in found and any(p + n == i for p, (_, n) in found.items()):
            left = next(found[p][0] for p, (_, n) in found.items() if p + n == i)
            if left not in auftrag.zeilen:
                for item in [w for w in auftrag.werte if w.spalte == left]:
                    auftrag.werte.remove(item)
                    if item.funktion in (None, "anzahl"):
                        auftrag.werte.append(Wert("anzahl", None))
                loose = [w for w in loose if w.spalte != left]
                auftrag.zeilen.append(left)
            auftrag.spalte = found[i + 1][0]
            i += 1 + found[i + 1][1]
            mode = "gruppe"
            continue
        elif word in _GROUP:
            mode = "gruppe"
        elif word in ("und", "sowie", "plus", "zuerst"):
            pass
        i += 1
    for function in pending:
        # „Anzahl pro Partei“: ohne Spalte nach der Rechenart.
        auftrag.werte.append(Wert(function, None))
    if not auftrag.zeilen and auftrag.spalte is not None:
        auftrag.zeilen, auftrag.spalte = [auftrag.spalte], None
    if not auftrag.zeilen:
        return Verstanden(None, _not_understood(auftrag, headers))
    if not auftrag.werte:
        auftrag.werte = [Wert("anzahl", None)]
    for item in auftrag.werte:
        if item.spalte is None and item.funktion not in (None, "anzahl"):
            name = function_name(item.funktion)
            others = [c for c in headers if c not in auftrag.zeilen]
            other = next((c for c in others if c in (numeric or ())), others[0] if others else None)
            example = T(f"{name} von {headers[other]}", f"{name} of {headers[other]}") if other is not None else name
            return Verstanden(None, T(f"{name} von welcher Spalte? Schreib sie dazu, zum Beispiel: "
                                      f"{example} pro {headers[auftrag.zeilen[0]]}.",
                                      f"{name} of which column? Add it, for example: "
                                      f"{example} per {headers[auftrag.zeilen[0]]}."))
    if len(auftrag.zeilen) > 2:
        return Verstanden(None, T("Mehr als zwei Spalten nebeneinander wird unübersichtlich. Nimm höchstens zwei, "
                                  "zum Beispiel: Summe der Sitze pro Partei und Jahr.",
                                  "More than two columns side by side gets hard to read. Use two at most, "
                                  "for example: sum of seats per party and year."))
    return Verstanden(auftrag)


def _example(headers: dict[int, str]) -> str:
    return next(iter(headers.values()), T("Werte", "values"))


def _not_understood(auftrag: Auftrag, headers: dict[int, str]) -> str:
    names = ", ".join(list(headers.values())[:8])
    if auftrag.werte:
        return T("Wonach soll ich aufteilen? Schreib „pro“ und eine Spalte dazu, zum Beispiel: "
                 f"… pro {_example(headers)}. Deine Spalten: {names}.",
                 "What should I split by? Add “per” and a column, for example: "
                 f"… per {_example(headers)}. Your columns: {names}.")
    return T(f"Ich finde in deinem Satz keine Spalte. Deine Spalten: {names}.",
             f"I cannot find a column in your sentence. Your columns: {names}.")


def understand(text: str, headers: dict[int, str], numeric: set[int] | None = None) -> Verstanden:
    """Den Satz ohne KI verstehen."""
    return parse(text, headers, numeric)
