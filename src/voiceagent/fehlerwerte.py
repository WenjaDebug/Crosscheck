"""Fehlerwerte in Zellen (#DIV/0!, #NAME?, #NV, #WERT! …): wo sie entstehen und warum.

Excel zeigt nur den Code, und jede Zelle, die eine Zelle mit Fehler liest, zeigt ihn auch. Der Prüfer sucht die
Zelle, in der der Fehler entsteht, und darin den Teil der Formel, der ihn erzeugt: den innersten, der den Fehler
ergibt, während alles, was er liest, noch ohne Fehler ist. Daraus folgt die Ursache: welcher Nenner leer oder 0 ist,
welcher Wert beim Nachschlagen fehlt, wo Text statt einer Zahl steht, welcher Name in der Formel Excel unbekannt ist.
Die Zellen, die den Fehler nur weiterreichen, nennt derselbe Hinweis mit.

Bei #NAME? kann auch alles richtig geschrieben sein. Dann hat Excel die Formel beim Öffnen nicht erkannt: Die Datei
kam aus einem anderen Programm, das neue Funktionen ohne den Zusatz speichert, den Excel dafür braucht, oder diese
Excel-Version kennt die Funktion noch nicht. Neu eingegeben rechnet sie; das setzt der Knopf „Einsetzen“ ein.

Behebt die Korrektur eines anderen Hinweises den Fehler schon (der Nenner, der beim Kopieren mitrutscht), sagt jener
Hinweis den Fehler mit, statt dass ein zweiter dasselbe meldet. Alles ohne KI.
"""

from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field

from . import einsetzen, formel
from .analyse import Finding, Refs, Sheet, address, changed_by
from .bereiche import _german_formula, _text
from .formel import Err, Node
from .nachschlagen import _shown
from .sprache import FORMULAS, T
from .sprache import english as _in_english
from .workbook import FUNCTIONS, pattern
from .xlsx import _date_text

Pos = tuple[int, int]

MAX_FINDINGS = 10
# So viele Zellen mit Fehler sieht sich der Prüfer einzeln an (große Blätter mit tausend #NV).
MAX_CELLS = 2000
MAX_LISTED = 5

_ENGLISH = dict(FUNCTIONS)
_FROM_GERMAN = {german: english for english, german in _ENGLISH.items()}
# Seit wann Excel eine Funktion kennt: In einer älteren Version bleibt #NAME? auch nach dem Neueingeben.
_SINCE = {
    **dict.fromkeys(("IFS", "MAXIFS", "MINIFS", "SWITCH", "CONCAT", "TEXTJOIN"),
                    ("erst ab Excel 2019", "only from Excel 2019 on")),
    **dict.fromkeys(("XLOOKUP", "XMATCH", "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY", "LET"),
                    ("erst ab Excel 2021", "only from Excel 2021 on")),
    **dict.fromkeys(("LAMBDA", "TEXTSPLIT", "TEXTBEFORE", "TEXTAFTER", "VSTACK", "HSTACK", "TAKE", "DROP",
                     "CHOOSECOLS", "CHOOSEROWS", "TOCOL", "TOROW"),
                    ("nur in Microsoft 365 und Excel 2024", "only in Microsoft 365 and Excel 2024")),
}
# Wie analyse.value_of die Codes liest (englisch); die übrigen Namen, die Excel je nach Sprache zeigt.
_CODES = {"#ÜBERLAUF!": "#SPILL!", "#KALK!": "#CALC!"}
_LOOKUPS = {"VLOOKUP": "der ersten Spalte von", "HLOOKUP": "der ersten Zeile von", "MATCH": "", "XMATCH": "",
            "XLOOKUP": "", "LOOKUP": ""}
_LOOKUPS_EN = {"VLOOKUP": "the first column of", "HLOOKUP": "the first row of"}
_SPREAD = {"STDEV.S": 2, "STDEV": 2, "VAR.S": 2, "VAR": 2, "STDEVA": 2, "STDEV.P": 1, "STDEVP": 1, "VAR.P": 1, "VARP": 1}
_OPS = {"+": "addiert", "-": "zieht ab", "*": "multipliziert", "/": "teilt", "^": "potenziert"}
_FAILS = (formel.Unknown, ValueError, TypeError, ZeroDivisionError, IndexError, KeyError, RecursionError, OverflowError)
_UNKNOWN = object()

# Texte in Anführungszeichen und Blattnamen in Hochkommas: Darin sucht der Prüfer keine Namen.
_QUOTED = re.compile(r'("(?:[^"]|"")*"|\'(?:[^\']|\'\')*\')')
_CALL = re.compile(r"(?<![A-Za-z0-9_.$\\])(?:_xlfn\.|_xlws\.)?([A-Za-zÄÖÜäöüß_][A-Za-z0-9ÄÖÜäöüß_.]*)\s*(?=\()")
_BARE = re.compile(r"(?<![A-Za-z0-9ÄÖÜäöüß_.$!:\\])([A-Za-zÄÖÜäöüß_][A-Za-z0-9ÄÖÜäöüß_.]*)(?![A-Za-z0-9ÄÖÜäöüß_.$(!:\[])")
_CELL = re.compile(r"[A-Za-z]{1,3}[1-9]\d{0,6}")
_NO_COLON = re.compile(r"([A-Za-z]{1,3}[1-9]\d{0,6})([A-Za-z]{1,3}[1-9]\d{0,6})")
_CURLY = re.compile(r"[„“”«»]")


@dataclass
class _Why:
    kind: str  # Art der Ursache: Heruntergezogene Zellen mit derselben Art kommen in einen Hinweis.
    text: str  # Was genau, als Satz ohne Punkt: „Die Formel teilt durch D10, und D10 ist leer“.
    tip: str = ""
    fixed: str | None = None  # Die korrigierte Formel für diese Zelle, englisch und ohne „=“.
    empty: bool = False  # Die Ursache ist eine leere Zelle (dann ist der Hinweis zur leeren Zelle derselbe).
    there: tuple[str, Pos] | None = None  # Der Fehler kommt von einem anderen Blatt.
    # Für heruntergezogene Zellen: derselbe Satz für alle, mit {} für das, was je Zeile anders ist (item).
    many: str = ""
    item: str = ""
    many_tip: str = ""


@dataclass
class _Group:
    cells: list[Pos]
    whys: list[_Why]
    owner: Finding | None = None
    after: list[Pos] = field(default_factory=list)

    @property
    def why(self) -> _Why:
        return self.whys[0]


def checks(sheet: Sheet, findings: list[Finding], hidden: set[str]) -> tuple[list[Finding], set[str]]:
    """Hinweise zu den Fehlerwerten des Blatts. findings: was die anderen Prüfungen schon melden (mit Korrekturen),
    hidden: error:… von Zellen, deren Fehler nachschlagen.py genauer erklärt. Zurück: die neuen Hinweise und die
    Schlüssel der Hinweise, die sie ersetzen."""
    errors = {pos: _code(v) for pos, v in sheet.inputs.items() if isinstance(v, Err)}
    if not errors:
        return [], set()
    calc, _ = sheet.evaluator()
    # Für #NAME?: so, wie Excel rechnet, wenn es die Formeln erkennt (die Zellen mit #NAME? neu gerechnet).
    fresh, _ = sheet.evaluator(chain={p for p, code in errors.items() if code == "#NAME?"})
    feeds: dict[Pos, list[Pos]] = {}
    whys: dict[Pos, _Why] = {}
    roots: list[Pos] = []
    for pos in sorted(errors, key=lambda p: (p[1], p[0])):
        code, formula = errors[pos], sheet.cells[pos].formula
        # Ein „@“ vor dem Funktionsnamen: Excel kennt die Funktion in dieser Zelle selbst nicht, auch wenn sie dazu noch
        # eine Zelle mit #NAME? liest.
        if formula is not None and not (code == "#NAME?" and _unrecognized(formula)):
            reads = sheet.precedents.get(pos, ())
            if isinstance(reads, Refs) and len(errors) < len(reads.singles) + sum(sheet.columns.count(a) for a in reads.areas):
                # Ein großer Bereich: die Fehlerwerte durchgehen statt aller Zellen darin.
                before = sorted(p for p, c in errors.items() if c == code and p != pos and p in reads)
            else:
                before = [p for p in reads if p != pos and errors.get(p) == code]
            if before:
                feeds[pos] = before
                continue
        roots.append(pos)
    covered = {f"recalc:{address(p)}" for p in errors}
    # Wohin ein weitergereichter Fehler zurückführt (die erste Ursprungszelle; im Kreis: er selbst).
    origin: dict[Pos, Pos] = {}
    for pos in sorted(feeds):
        origin[pos] = _root(pos, feeds, origin)
    roots += [p for p in feeds if origin[p] == p]
    shown = {p: _code_shown(sheet, p) for p in errors}
    owners = _owners(sheet, calc, findings, errors)
    seen = 0
    for pos in sorted(roots, key=lambda p: (p[1], p[0])):
        if f"error:{address(pos)}" in hidden:
            continue
        if seen >= MAX_CELLS:
            break
        seen += 1
        whys[pos] = _why(sheet, fresh if errors[pos] == "#NAME?" else calc, pos, errors[pos], shown[pos])
    groups = _grouped(sheet, whys, errors, owners)
    passed: dict[Pos, list[Pos]] = {}
    for pos, root in sorted(origin.items(), key=lambda item: (item[0][0], item[0][1])):
        if root != pos:
            passed.setdefault(root, []).append(pos)
    for group in groups:
        group.after = sorted((p for root in group.cells for p in passed.get(root, ())), key=lambda p: (p[0], p[1]))
        if any(why.empty for why in group.whys):
            covered |= {f"emptyref:{address(p)}" for p in group.cells}
    # Weitergereichte Fehler aus Zellen, die nachschlagen.py erklärt: Der Hinweis dort nennt sie mit.
    for item in findings:
        tail = item.key.partition(":")[2]
        if f"error:{tail}" in hidden and passed.get(formel.position(tail)):
            item.text += _follows(passed[formel.position(tail)])
    found: list[Finding] = []
    rest: list[_Group] = []
    unnamed: dict[int, tuple[Finding, dict[str, list[Pos]]]] = {}
    for group in groups:
        if group.owner is not None:
            # Sagt der Hinweis den Fehler schon selbst („darum ergibt C6 #NV“), bleibt es dabei.
            code = errors[group.cells[0]]
            if not any(word in group.owner.text for word in {shown[group.cells[0]], code, _shown(Err(code))}):
                group.owner.text += " " + _along(sheet, group, shown)
            else:
                # Nennt er die Zelle selbst nicht („und 2 weitere Formeln“), dann kurz dazu.
                fixed = (group.owner.fix or {}).get("cells", {})
                missing = [p for p in group.cells + group.after if address(p) not in fixed
                           and not re.search(rf"(?<![A-Z]){address(p)}(?!\d)", group.owner.text)]
                if missing:
                    unnamed.setdefault(id(group.owner), (group.owner, {}))[1].setdefault(shown[group.cells[0]], []).extend(missing)
            continue
        if len(found) >= MAX_FINDINGS:
            rest.append(group)
            continue
        found.append(_finding(sheet, group, shown, fresh))
    for owner, by_code in unnamed.values():
        for code, cells in by_code.items():
            cells = sorted(set(cells), key=lambda p: (p[0], p[1]))
            verb = "zeigt" if len(cells) == 1 else "zeigen"
            owner.text += T(f" Deshalb {verb} auch {_listed(cells)} {code}.",
                            f" That is why {_listed(cells)} {'shows' if len(cells) == 1 else 'show'} {code} as well.")
    unexplained = [p for p in roots if p not in whys and f"error:{address(p)}" not in hidden]
    if rest or unexplained:
        cells = sorted({p for g in rest for p in g.cells} | set(unexplained), key=lambda p: (p[0], p[1]))
        listed = ", ".join(f"{address(p)} ({shown[p]})" for p in cells[:MAX_LISTED])
        more = (T(f" und {len(cells) - MAX_LISTED} weitere", f" and {len(cells) - MAX_LISTED} more")
                if len(cells) > MAX_LISTED else "")
        found.append(Finding(
            f"error:{address(cells[0])}" if not any(f.key == f"error:{address(cells[0])}" for f in found) else "error:weitere",
            T(f"Noch mehr Zellen zeigen Fehlerwerte: {listed}{more}.", f"More cells show error values: {listed}{more}."),
            "fehler",
            T("Klick die Zelle an und sieh nach, womit die Formel rechnet: Oft ist es eine leere Zelle oder Text.",
              "Click the cell and check what the formula calculates with: often it is an empty cell or text."),
            cell=address(cells[0]),
        ))
    return found, covered


# --- Woher der Fehler kommt ---


def _code(value: Err) -> str:
    return _CODES.get(value.code, value.code)


_GERMAN_CODES = {"#N/A": "#NV", "#VALUE!": "#WERT!", "#REF!": "#BEZUG!", "#NUM!": "#ZAHL!", "#SPILL!": "#ÜBERLAUF!",
                 "#CALC!": "#KALK!"}
_ENGLISH_CODES = {german: english for english, german in _GERMAN_CODES.items()}


def _code_shown(sheet: Sheet, pos: Pos) -> str:
    """Der Fehler so, wie das deutsche Excel ihn in der Zelle zeigt: In der Datei steht er immer englisch (#N/A),
    zu sehen ist #NV. Zeigt Excel Formeln englisch, ist es der englische Code."""
    text = sheet.cells[pos].value.strip()
    text = text if text.startswith("#") else _shown(sheet.inputs[pos])
    if FORMULAS.get() == "en":
        return _ENGLISH_CODES.get(text.upper(), text)
    return _GERMAN_CODES.get(text.upper(), text)


def _root(pos: Pos, feeds: dict[Pos, list[Pos]], origin: dict[Pos, Pos]) -> Pos:
    """Die erste Zelle hinter pos, die den Fehler selbst erzeugt (nicht nur weiterreicht)."""
    seen, stack = {pos}, [pos]
    while stack:
        here = stack.pop()
        for before in sorted(feeds.get(here, ()), key=lambda p: (p[1], p[0]), reverse=True):
            if before in seen:
                continue
            if before not in feeds:
                return before
            if before in origin and origin[before] != before:
                return origin[before]
            seen.add(before)
            stack.append(before)
    return pos


def _value(calc, node: Node):
    try:
        return calc.run(node)
    except _FAILS:
        return _UNKNOWN


def _is(value, code: str) -> bool:
    return isinstance(value, Err) and _code(value) == code


def _origin(calc, node: Node, code: str) -> Node:
    """Der innerste Teil von node, der den Fehler code ergibt, während alles, was er liest, ohne Fehler ist."""
    while True:
        child = _culprit(calc, node, code)
        if child is None:
            return node
        node = child


def _culprit(calc, node: Node, code: str) -> Node | None:
    args = [a for a in node.args if isinstance(a, Node)]
    if node.kind == "call" and node.value == "IF" and args:
        test = _value(calc, args[0])
        if _is(test, code):
            return args[0]
        try:
            taken = args[1] if formel._truth(test) else (args[2] if len(args) > 2 else None)  # noqa: SLF001
        except (formel._Fail, *_FAILS):  # noqa: SLF001
            return None
        return taken if taken is not None and _is(_value(calc, taken), code) else None
    if node.kind == "call" and node.value in ("IFERROR", "IFNA"):
        return args[1] if len(args) > 1 and _is(_value(calc, args[1]), code) else None
    for child in args:
        if child.kind not in ("range", "xrange", "empty") and _is(_value(calc, child), code):
            return child
    return None


def _there(sheet: Sheet, calc, node: Node, code: str) -> tuple[str, Pos] | None:
    """Eine Zelle auf einem anderen Blatt, die node liest und die schon den Fehler zeigt."""
    for item in _walk(node):
        if item.kind == "xref" and _is(_value(calc, item), code):
            return item.value
        if item.kind == "xrange":
            try:
                name, (r1, c1, r2, c2) = calc.area_of(item)
            except _FAILS:
                continue
            if (r2 - r1 + 1) * (c2 - c1 + 1) > 20_000:
                continue
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    try:
                        value = sheet.other(name, (r, c))
                    except _FAILS:
                        return None
                    if _is(value, code):
                        return name, (r, c)
    return None


def _walk(node: Node):
    yield node
    for child in node.args:
        if isinstance(child, Node):
            yield from _walk(child)


# --- Warum ---


def _say(node: Node) -> str:
    """Ein Teil der Formel so, wie er im deutschen Excel aussieht. Zeigt Excel Formeln englisch, so wie dort (ohne „=“
    davor schreibt sprache.output ihn nicht um), mit Leerzeichen nach dem Komma: So liest sprache „-3, 0“ nicht als
    Dezimalzahl."""
    text = _text(node, {})
    if text is None:
        return "…"
    if FORMULAS.get() == "en":
        return "".join(p if i % 2 else p.replace(",", ", ") for i, p in enumerate(_QUOTED.split(text)))
    return _german_formula(text)


def _german(name: str) -> str:
    return _ENGLISH.get(name, name)


def _fn(name: str) -> str:
    """Ein Funktionsname im Satz: deutsch, auf Englisch der englische."""
    return name if _in_english() else _german(name)


def _sep() -> str:
    """Das Trennzeichen zwischen Argumenten, wie Excel es zeigt."""
    return "," if FORMULAS.get() == "en" else ";"


def _why(sheet: Sheet, calc, pos: Pos, code: str, shown: str) -> _Why:
    cell = sheet.cells[pos]
    if cell.formula is None:
        return _Why("typed", T(f"{shown} steht dort als fester Wert, nicht aus einer Formel (oft beim Einfügen mitkopiert)",
                               f"{shown} is there as a fixed value, not from a formula (often copied along when pasting)"),
                    T("Trag den richtigen Wert ein oder lass die Zelle leer.", "Enter the right value or leave the cell empty."))
    node = sheet.nodes.get(pos)
    if code == "#NAME?":
        return _name(sheet, calc, pos, node)
    if code == "#REF!" and "#REF!" in cell.formula.upper():
        return _Why("ref", T(f"In der Formel steht {shown} statt einer Zelle: Sie zeigte auf Zeilen, Spalten oder ein Blatt, "
                             "die inzwischen gelöscht sind",
                             f"The formula contains {shown} instead of a cell: it pointed to rows, columns or a sheet "
                             "that have since been deleted"),
                    T("Setz den Bezug in der Formel neu. Hast du eben etwas gelöscht, holt Befehl-Z es zurück.",
                      "Set the reference in the formula again. If you just deleted something, Cmd+Z brings it back."))
    if code in _FIXED:
        return _Why(code, *_fixed_why(code))
    if node is None:
        return _Why("other", T("Die Formel kann ich nicht nachrechnen, also auch nicht sagen, woher der Fehler kommt",
                               "I cannot recalculate the formula, so I cannot say where the error comes from"), "")
    there = _there(sheet, calc, node, code)
    if there is not None:
        name, at = there
        return _Why("there", T(f"Die Formel liest {name}!{address(at)}, und dort steht schon {shown}",
                               f"The formula reads {name}!{address(at)}, and that cell already shows {shown}"),
                    T(f"Die Ursache liegt auf dem Blatt „{name}“ in {address(at)}: Ist der Fehler dort behoben, "
                      "verschwindet er hier auch.",
                      f"The cause is on the sheet “{name}” in {address(at)}: once the error is fixed there, it "
                      "disappears here too."), there=there)
    whole = _value(calc, node)
    part = _origin(calc, node, code) if _is(whole, code) else None
    if code == "#DIV/0!":
        why = _divide(sheet, calc, node, part)
        if why.kind in ("divide-empty", "divide-zero") and part is node:
            return _like(sheet, calc, pos, node) or why
        return why
    if code == "#N/A":
        return _missing(sheet, calc, node, part, shown)
    if code == "#VALUE!":
        return _wrong_kind(sheet, calc, part, shown)
    if code == "#NUM!":
        return _number(calc, part, node)
    if code == "#REF!":
        dynamic = _dynamic(sheet, calc, node)
        if dynamic is not None:
            return dynamic
    if part is not None and part is not node:
        return _Why("other", T(f"Der Fehler entsteht in {_say(part)}", f"The error arises in {_say(part)}"), "")
    return _Why("other", T("Woher der Fehler kommt, kann ich nicht nachrechnen", "I cannot recalculate where the error comes from"),
                "")


def _dynamic(sheet: Sheet, calc, node: Node) -> _Why | None:
    """#BEZUG! aus BEREICH.VERSCHIEBEN, das über den Rand des Blatts zeigt, oder aus INDIREKT mit einem Text, der
    auf kein Blatt der Mappe zeigt."""
    for part in _walk(node):
        if part.kind != "call" or part.value not in ("OFFSET", "INDIRECT") or not part.args:
            continue
        if part.value == "INDIRECT":
            missing = _missing_sheet(sheet, calc, part)
            if missing is not None:
                return missing
        try:
            calc.reference(part)
            continue
        except formel._Fail as fail:  # noqa: SLF001
            if not _is(fail.err, "#REF!"):
                continue
        except _FAILS:
            if part.value != "INDIRECT" or not _no_reference(sheet, _value(calc, part.args[0])):
                continue
        if part.value == "INDIRECT":
            text = _value(calc, part.args[0])
            empty = text is None or text == ""
            given = T("einen leeren Text", "an empty text") if empty else _shown(text)
            return _Why("indirect", T(f"{_say(part)} bekommt {given}, und das ist "
                                      "keine Zelle und kein Bereich",
                                      f"{_say(part)} gets {given}, which is neither a cell nor a range"),
                        T("Prüf den Text, aus dem INDIREKT den Bezug baut, zum Beispiel „B“&5 oder „Daten!B5“.",
                          "Check the text INDIRECT builds the reference from, for example “B”&5 or “Data!B5”."))
        try:
            base = calc.reference(part.args[0])
            rows, cols = calc._whole(part.args[1]), calc._whole(part.args[2])  # noqa: SLF001
        except (formel._Fail, *_FAILS):  # noqa: SLF001
            return None
        where = base.value[1] if base.kind in ("xref", "xrange") else base.value
        top, left = where[0], where[1]
        if top + rows < 1:
            edge = T(f"{-rows} Zeilen über {address((top, left))}, über den oberen Rand des Blatts hinaus",
                     f"{-rows} {'row' if rows == -1 else 'rows'} above {address((top, left))}, beyond the top edge of "
                     "the sheet")
        elif left + cols < 1:
            edge = T(f"{-cols} Spalten links von {address((top, left))}, über den linken Rand des Blatts hinaus",
                     f"{-cols} {'column' if cols == -1 else 'columns'} to the left of {address((top, left))}, beyond the "
                     "left edge of the sheet")
        else:
            edge = T("über den Rand des Blatts hinaus", "beyond the edge of the sheet")
        return _Why("offset", T(f"{_say(part)} zeigt {edge}", f"{_say(part)} points {edge}"),
                    T("Prüf die Zahl der Zeilen und Spalten in BEREICH.VERSCHIEBEN, oft fehlt dafür ein WENN am Anfang "
                      "der Tabelle.",
                      "Check the number of rows and columns in OFFSET; often an IF for the start of the table is "
                      "missing."))
    return None


def _no_reference(sheet: Sheet, text) -> bool:
    """Ein Text für INDIREKT, der weder Zelle noch Bereich noch ein Name der Mappe ist („Bx“)."""
    if not isinstance(text, str) or not text.strip() or "!" in text:
        return False
    names = {n.casefold() for n in getattr(sheet, "names", {}) or {}}
    return text.strip().casefold() not in names and re.fullmatch(r"[A-Za-z_\\][\w.]*", text.strip()) is not None


def _missing_sheet(sheet: Sheet, calc, part: Node) -> _Why | None:
    text = _value(calc, part.args[0]) if part.args else None
    if not isinstance(text, str) or "!" not in text:
        return None
    try:
        target = formel.parse(formel._upper_outside_quotes(text.strip()))  # noqa: SLF001
    except formel.Unknown:
        return None
    if target.kind not in ("xref", "xrange"):
        return None
    # Der Blattname, wie er im Text steht (nicht groß geschrieben wie zum Lesen).
    name = text.strip().rsplit("!", 1)[0].strip()
    name = name[1:-1].replace("''", "'") if len(name) > 1 and name[0] == name[-1] == "'" else name
    known = {n.casefold() for n in sheet.book} | ({sheet.name.casefold()} if sheet.name else set())
    if name.casefold() in known:
        return None
    return _Why("indirect", T(f"{_say(part)} zeigt auf das Blatt „{name}“, und das gibt es in dieser Mappe nicht",
                              f"{_say(part)} points to the sheet “{name}”, which does not exist in this workbook"),
                T("Prüf den Blattnamen im Text für INDIREKT; ein umbenanntes Blatt ändert INDIREKT nicht mit.",
                  "Check the sheet name in the text for INDIRECT; renaming a sheet does not update INDIRECT."))


_FIXED = {
    "#NULL!": ("Zwei Bereiche stehen nur mit einem Leerzeichen hintereinander, und sie überschneiden sich nicht",
               "Gemeint ist meist ein ; zwischen den Bereichen (oder ein : für einen zusammenhängenden Bereich)."),
    "#SPILL!": ("Das Ergebnis braucht mehrere Zellen, aber daneben oder darunter steht schon etwas",
                "Mach die Zellen frei, in die das Ergebnis laufen soll."),
    "#CALC!": ("Die Formel ergibt eine leere Liste, etwa FILTER ohne einen einzigen Treffer",
               "FILTER nimmt als drittes Argument, was dann stehen soll: FILTER(…;…;\"keine\")."),
}


def _fixed_why(code: str) -> tuple[str, str]:
    """Satz und Tipp zu einem Fehler, dessen Ursache der Code allein sagt (_FIXED), in der Sprache der Anfrage."""
    if not _in_english():
        return _FIXED[code]
    sep = _sep()
    return {
        "#NULL!": ("Two ranges are separated only by a space, and they do not overlap",
                   f"Usually a {sep} between the ranges is meant (or a : for a contiguous range)."),
        "#SPILL!": ("The result needs several cells, but there is already something next to or below it",
                    "Clear the cells the result should spill into."),
        "#CALC!": ("The formula gives an empty list, for example FILTER without a single match",
                   f"FILTER takes as its third argument what should appear then: FILTER(…{sep}…{sep}\"none\")."),
    }[code]


def _cell_word(sheet: Sheet, calc, node: Node) -> tuple[str, str, bool] | None:
    """Für eine Zelle in der Formel: wie sie heißt, was in ihr steht („leer“, „0“) und ob sie leer ist."""
    if node.kind not in ("ref", "xref"):
        return None
    name = address(node.value) if node.kind == "ref" else _say(node)
    value = _value(calc, node)
    if value is None or value == "":
        return name, T("leer", "empty"), True
    if isinstance(value, float) and not isinstance(value, bool) and value == 0:
        formula = node.kind == "ref" and sheet.cells.get(node.value) is not None and sheet.cells[node.value].formula
        return name, T("ergibt 0", "gives 0") if formula else T("ist 0", "is 0"), False
    return None


def _divide(sheet: Sheet, calc, node: Node, part: Node | None) -> _Why:
    if part is None:
        # Nicht nachzurechnen: Wenigstens ein Nenner, der leer oder 0 ist, sagt meistens schon alles.
        for item in _walk(node):
            if item.kind == "op" and item.value == "/":
                said = _cell_word(sheet, calc, item.args[1])
                if said:
                    return _zero(*said)
        return _Why("divide", T("Irgendwo in der Formel wird durch 0 oder eine leere Zelle geteilt",
                                "Somewhere in the formula, something is divided by 0 or by an empty cell"),
                    T("Meist fehlt ein Wert, oder ein Bereich ist leer.", "Usually a value is missing, or a range is empty."))
    if part.kind == "op" and part.value == "/":
        bottom = part.args[1]
        said = _cell_word(sheet, calc, bottom)
        if said:
            return _zero(*said)
        if bottom.kind == "num":
            return _Why("divide-literal", T("Die Formel teilt durch die Zahl 0", "The formula divides by the number 0"),
                        T("Durch 0 lässt sich nicht teilen: Prüf, welche Zahl oder Zelle hier gemeint ist.",
                          "You cannot divide by 0: check which number or cell is meant here."))
        return _Why("divide-expr", T(f"Die Formel teilt durch {_say(bottom)}, und das ergibt 0",
                                     f"The formula divides by {_say(bottom)}, and that gives 0"), "")
    if part.kind == "err":
        return _Why("literal", T(f"In der Formel steht {part.value} fest drin", f"The formula contains {part.value} as a fixed value"),
                    "")
    if part.kind != "call":
        return _Why("divide", T(f"Der Fehler entsteht in {_say(part)}", f"The error arises in {_say(part)}"), "")
    name, args = part.value, part.args
    if name in ("AVERAGE", "AVERAGEA") and args:
        return _Why("average", T(f"{_say(part)} findet keine einzige Zahl{_contents(sheet, calc, args)}",
                                 f"{_say(part)} does not find a single number{_contents(sheet, calc, args)}"),
                    T("Prüf den Bereich: Stehen die Zahlen woanders, oder sind sie als Text gespeichert (grünes Dreieck)?",
                      "Check the range: are the numbers somewhere else, or are they stored as text (green triangle)?"))
    if name in ("AVERAGEIF", "AVERAGEIFS") and len(args) >= 2:
        return _no_match(calc, name, args)
    if name in _SPREAD:
        least = _SPREAD[name]
        count = _count(calc, args)
        if not _in_english():
            has = {0: "keine", 1: "nur eine"}.get(count, str(count))
            need = "mindestens zwei Zahlen" if least == 2 else "mindestens eine Zahl"
            return _Why("spread", f"{_german(name)} braucht {need}, in {';'.join(_say(a) for a in args)} steht {has}",
                        "Prüf den Bereich: Stehen die Zahlen woanders, oder sind sie als Text gespeichert?")
        has = {0: "none", 1: "only one"}.get(count, str(count))
        need = "at least two numbers" if least == 2 else "at least one number"
        return _Why("spread", f"{name} needs {need}, but {(_sep() + ' ').join(_say(a) for a in args)} contains {has}",
                    "Check the range: are the numbers somewhere else, or are they stored as text?")
    if name in ("MOD", "QUOTIENT") and len(args) == 2:
        said = _cell_word(sheet, calc, args[1])
        what = (T(f"{said[0]}, und {said[0]} {'ist leer' if said[2] else said[1]}",
                  f"{said[0]}, and {said[0]} {'is empty' if said[2] else said[1]}") if said
                else T(f"{_say(args[1])}, und das ergibt 0", f"{_say(args[1])}, and that gives 0"))
        return _Why("divide", T(f"{_german(name)} teilt durch {what}", f"{name} divides by {what}"), "",
                    empty=bool(said and said[2]))
    return _Why("divide", T(f"Der Fehler entsteht in {_say(part)}", f"The error arises in {_say(part)}"),
                T("Meist fehlt dort ein Wert, oder ein Bereich ist leer.", "Usually a value is missing there, or a range is empty."))


def _zero(name: str, word: str, empty: bool) -> _Why:
    if empty:
        return _Why("divide-empty", T(f"Die Formel teilt durch {name}, und {name} ist leer",
                                      f"The formula divides by {name}, and {name} is empty"),
                    many=T("Die Formeln teilen durch {}, und die sind leer", "The formulas divide by {}, and those are empty"),
                    item=name,
                    many_tip=T("Trag dort die Werte ein, durch die geteilt werden soll, oder prüf, ob die Formeln auf die "
                               "richtigen Zellen zeigen.",
                               "Enter the values to divide by there, or check whether the formulas point to the right "
                               "cells."),
                    tip=T(f"Trag in {name} den Wert ein, durch den geteilt werden soll, oder prüf, ob die Formel auf die "
                          "richtige Zelle zeigt.",
                          f"Enter the value to divide by in {name}, or check whether the formula points to the right "
                          "cell."), empty=True)
    return _Why("divide-zero", T(f"Die Formel teilt durch {name}, und {name} {word}",
                                 f"The formula divides by {name}, and {name} {word}"),
                T(f"Prüf, ob in {name} wirklich 0 stehen soll und ob die Formel auf die richtige Zelle zeigt.",
                  f"Check whether {name} should really be 0 and whether the formula points to the right cell."))


def _like(sheet: Sheet, calc, pos: Pos, node: Node) -> _Why | None:
    """Teilen die Formeln darüber und darunter alle durch dieselbe Zelle (meist die Summe), ist sie wohl auch hier
    gemeint: =B4/B6 unter =B2/B5 und =B3/B5, und B6 ist leer."""
    bottom = node.args[1]
    if bottom.kind != "ref":
        return None
    shape = pattern(_numerator(sheet.cells[pos].formula or ""), pos)
    others: list[tuple[Pos, Pos, str]] = []
    for step in (-1, 1):
        r = pos[0] + step
        while 1 <= r and len(others) < 40:
            there = (r, pos[1])
            item, cell = sheet.nodes.get(there), sheet.cells.get(there)
            if item is None or cell is None or item.kind != "op" or item.value != "/" or item.args[1].kind != "ref":
                break
            if pattern(_numerator(cell.formula or ""), there) != shape:
                break
            others.append((there, item.args[1].value, (cell.formula or "").rpartition("/")[2].strip()))
            r += step
    targets = {target for _, target, _ in others}
    if len(others) < 2 or len(targets) != 1 or bottom.value in targets:
        return None
    target = next(iter(targets))
    value = _value(calc, Node("ref", target))
    head, sep, tail = (sheet.cells[pos].formula or "").rpartition("/")
    if not sep or tail.strip().replace("$", "") != address(bottom.value) or not isinstance(value, float) or value == 0:
        return None
    said = _cell_word(sheet, calc, bottom)
    if said is None:
        return None
    rows = sorted(p for p, _, _ in others)
    word = T("ist leer", "is empty") if said[2] else said[1]
    return _Why("divide-like", T(f"{_listed(rows)} teilen durch {address(target)}, diese Zeile aber durch {said[0]}, und "
                                 f"{said[0]} {word}",
                                 f"{_listed(rows)} divide by {address(target)}, but this row divides by {said[0]}, and "
                                 f"{said[0]} {word}"), "", fixed=f"{head}/{others[0][2]}", empty=said[2])


def _numerator(formula: str) -> str:
    return formula.rpartition("/")[0]


def _count(calc, args) -> int:
    try:
        values = calc.flat([a for a in args if isinstance(a, Node)])
    except (formel._Fail, *_FAILS):  # noqa: SLF001
        return 0
    return sum(1 for v in values if isinstance(v, float) and not isinstance(v, bool))


def _contents(sheet: Sheet, calc, args) -> str:
    """„: E2:E9 ist leer“ oder „: in E2:E9 steht nur Text“."""
    areas = [a for a in args if a.kind in ("range", "xrange")]
    if len(areas) != 1 or len(args) != 1:
        return ""
    try:
        values = calc.values_of(areas[0])
    except (formel._Fail, *_FAILS):  # noqa: SLF001
        return ""
    where = _say(areas[0])
    if all(v is None or v == "" for v in values):
        return T(f": {where} ist leer", f": {where} is empty")
    texts = [v for v in values if isinstance(v, str) and v != ""]
    if texts and all(formel._as_float(v) is not None for v in texts):  # noqa: SLF001
        return T(f": In {where} stehen die Zahlen als Text", f": in {where}, the numbers are stored as text")
    if texts:
        return T(f": In {where} steht nur Text", f": {where} contains only text")
    return ""


def _no_match(calc, name: str, args) -> _Why:
    """MITTELWERTWENN ohne Zeile, auf die die Bedingung passt, oder mit Treffern ohne Zahl."""
    if name == "AVERAGEIF":
        pairs, values = [(args[0], args[1])], args[2] if len(args) > 2 else args[0]
    else:
        pairs, values = list(zip(args[1::2], args[2::2], strict=False)), args[0]
    tests = [_criterion(calc, c) for _, c in pairs]
    count = _value(calc, Node("call", "COUNTIFS", tuple(x for pair in pairs for x in pair)))
    said = T(" und ", " and ").join(tests)
    if isinstance(count, float) and count > 0:
        if _in_english():
            rows = "One row matches" if count == 1 else f"{_shown(count)} rows match"
            return _Why("average-if", f"{rows} {said}, but {_say(values)} has no number in "
                                      f"{'that row' if count == 1 else 'those rows'}",
                        "Check whether numbers are stored as text there or whether the range has shifted.")
        rows = "eine Zeile" if count == 1 else f"{_shown(count)} Zeilen"
        return _Why("average-if", f"Auf {said} {'passt' if count == 1 else 'passen'} {rows}, aber in {_say(values)} "
                                  "steht dort keine Zahl", "Prüf, ob dort Zahlen als Text stehen oder ob der Bereich verrutscht ist.")
    if _in_english():
        which = "the condition" if len(tests) == 1 else "all conditions"
        return _Why("average-if", f"{name} finds no row that matches {said}",
                    f"Check the spelling of {which} (spaces; upper and lower case do not matter). If there is rightly "
                    f"no row, IFERROR(…{_sep()}\"\") shows nothing there.")
    which = "die Bedingung" if len(tests) == 1 else "alle Bedingungen"
    return _Why("average-if", f"{_german(name)} findet keine Zeile, auf die {said} {'passt' if len(tests) == 1 else 'passen'}",
                f"Prüf die Schreibweise von {which} (Leerzeichen, Groß und klein spielt keine Rolle). Gibt es zu Recht "
                "keine Zeile, zeigt WENNFEHLER(…;\"\") dort nichts an.")


def _criterion(calc, node: Node) -> str:
    value = _value(calc, node)
    if node.kind in ("ref", "xref") and value is not _UNKNOWN:
        return f"{_say(node)} ({_shown(value)})"
    return _shown(value) if value is not _UNKNOWN else _say(node)


def _missing(sheet: Sheet, calc, node: Node, part: Node | None, shown: str) -> _Why:
    if part is not None and part.kind == "call" and part.value == "NA":
        return _Why("na", T(f"Die Formel gibt mit NV() absichtlich {shown} aus", f"The formula returns {shown} on purpose with NA()"),
                    "")
    if part is not None and part.kind == "err":
        return _Why("literal", T(f"In der Formel steht {part.value} fest drin", f"The formula contains {part.value} as a fixed value"),
                    "")
    if part is not None and part.kind == "call" and part.value in _LOOKUPS and len(part.args) >= 2:
        key, table = part.args[0], part.args[1]
        value = _value(calc, key)
        function = _fn(part.value)
        if value is None or value == "":
            return _Why("lookup-empty", T(f"Der Suchwert {_say(key)} ist leer, darum findet {function} nichts",
                                          f"The lookup value {_say(key)} is empty, so {function} finds nothing"),
                        T(f"Trag in {_say(key)} ein, was gesucht werden soll.", f"Enter what should be looked up in {_say(key)}."),
                        empty=True)
        where = T(f"{_LOOKUPS[part.value]} {_say(table)}", f"{_LOOKUPS_EN.get(part.value, '')} {_say(table)}").strip()
        exact = _exactly(calc, part)
        if exact is not None:
            return _Why("lookup-rough", T(f"{function} sucht ungefähr, weil das letzte Argument fehlt oder WAHR ist, und "
                                          f"{_say(table)} ist nicht sortiert: So findet es {_shown(value)} nicht, obwohl es "
                                          f"dort steht (genau gesucht: {_shown(exact)})",
                                          f"{function} uses an approximate match because the last argument is missing or "
                                          f"TRUE, and {_say(table)} is not sorted: so it does not find {_shown(value)}, "
                                          f"even though it is there (exact match: {_shown(exact)})"),
                        _exact_tip(part.value))
        return _Why("lookup", T(f"{function} findet {_shown(value)} nicht in {where}",
                                f"{function} does not find {_shown(value)} in {where}"),
                    many=T(f"{function} findet {{}} nicht in {where}", f"{function} does not find {{}} in {where}"),
                    item=_shown(value),
                    many_tip=T("Stehen sie wirklich dort? Prüf die Schreibweise (Leerzeichen, Zahl als Text). Fehlen "
                               "Einträge zu Recht, zeigt WENNNV(…;\"fehlt\") einen eigenen Text statt des Fehlers.",
                               "Are they really there? Check the spelling (spaces, numbers stored as text). If entries "
                               f"are rightly missing, IFNA(…{_sep()}\"missing\") shows your own text instead of the error."),
                    tip=T(f"Steht {_shown(value)} wirklich dort? Prüf die Schreibweise (Leerzeichen, Zahl als Text). Fehlt "
                          "der Eintrag zu Recht, zeigt WENNNV(…;\"fehlt\") einen eigenen Text statt des Fehlers.",
                          f"Is {_shown(value)} really there? Check the spelling (spaces, number stored as text). If the "
                          f"entry is rightly missing, IFNA(…{_sep()}\"missing\") shows your own text instead of the "
                          "error."))
    if part is not None and part is not node:
        return _Why("other", T(f"Der Fehler entsteht in {_say(part)}: Dort fehlt ein Wert, den die Formel braucht",
                               f"The error arises in {_say(part)}: a value the formula needs is missing there"), "")
    calls = [c.value for c in _calls(node) if c.kind == "call"]
    if "NA" in calls:
        return _Why("na", T(f"Die Formel gibt mit NV() absichtlich {shown} aus", f"The formula returns {shown} on purpose with NA()"),
                    "")
    found = next((c for c in _calls(node) if c.kind == "call" and c.value in _LOOKUPS), None)
    if found is not None and _exactly(calc, found) is not None:
        return _Why("lookup-rough", T(f"{_german(found.value)} sucht ungefähr, weil das letzte Argument fehlt oder WAHR ist. "
                                      "In einer Tabelle, die nicht sortiert ist, findet es so Werte nicht, die dort stehen",
                                      f"{found.value} uses an approximate match because the last argument is missing or "
                                      "TRUE. In a table that is not sorted, it then misses values that are there"),
                    _exact_tip(found.value))
    if found is not None:
        return _Why("lookup", T(f"{_german(found.value)} liefert #NV, meist weil der Suchwert nicht in der Tabelle steht",
                                f"{found.value} returns {shown}, usually because the lookup value is not in the table"),
                    T("Prüf die Schreibweise (Leerzeichen, Zahl als Text).", "Check the spelling (spaces, number stored as text)."))
    return _Why("other", T("Ein Wert, den die Formel braucht, fehlt", "A value the formula needs is missing"), "")


def _exact_tip(function: str) -> str:
    return T(f"Schreib {'0' if function == 'MATCH' else 'FALSCH'} als letztes Argument, dann sucht Excel genau.",
             f"Put {'0' if function == 'MATCH' else 'FALSE'} as the last argument, then Excel looks for an exact match.")


def _exactly(calc, call: Node):
    """Was SVERWEIS, WVERWEIS oder VERGLEICH genau gesucht fände, wenn es jetzt ungefähr sucht; sonst None."""
    if not _rough(calc, call):
        return None
    index = {"VLOOKUP": 3, "HLOOKUP": 3, "MATCH": 2}[call.value]
    flag = Node("num", 0.0) if call.value == "MATCH" else Node("bool", False)
    args = list(call.args[:index]) + [Node("empty")] * (index - len(call.args))
    value = _value(calc, Node("call", call.value, (*args, flag)))
    return None if value is _UNKNOWN or value is None or isinstance(value, Err) else value


def _rough(calc, call: Node) -> bool:
    """Sucht SVERWEIS, WVERWEIS oder VERGLEICH ungefähr (letztes Argument fehlt, WAHR oder 1)?"""
    index = {"VLOOKUP": 3, "HLOOKUP": 3, "MATCH": 2}.get(call.value)
    if index is None:
        return False
    if index >= len(call.args) or call.args[index].kind == "empty":
        return True
    flag = _value(calc, call.args[index])
    return flag is True or (isinstance(flag, float) and flag != 0)


def _wrong_kind(sheet: Sheet, calc, part: Node | None, shown: str) -> _Why:
    if part is not None and part.kind in ("op", "neg", "pct") and (part.kind != "op" or part.value in _OPS):
        for child in part.args:
            value = _value(calc, child)
            if isinstance(value, str) and value != "" and formel._as_float(value) is None:  # noqa: SLF001
                if child.kind in ("ref", "xref"):
                    name = _say(child)
                    return _Why("text", T(f"Die Formel rechnet mit {name}, und dort steht Text ({_shown(value)})",
                                          f"The formula calculates with {name}, and it contains text ({_shown(value)})"),
                                many=T("Die Formeln rechnen mit {}, und dort steht Text",
                                       "The formulas calculate with {}, and those contain text"), item=name,
                                many_tip=T("Trag dort Zahlen ein oder lass die Zellen leer.",
                                           "Enter numbers there or leave the cells empty."),
                                tip=T(f"Trag in {name} eine Zahl ein oder lass die Zelle leer.",
                                      f"Enter a number in {name} or leave the cell empty."))
                return _Why("text", T(f"Die Formel rechnet mit {_say(child)}, und das ergibt Text ({_shown(value)})",
                                      f"The formula calculates with {_say(child)}, and that gives text ({_shown(value)})"), "")
    if part is not None and part.kind == "call":
        for child in part.args:
            value = _value(calc, child)
            if isinstance(value, str) and value != "" and formel._as_float(value) is None and formel.serial_of(value) is None:  # noqa: SLF001
                return _Why("text", T(f"{_german(part.value)} bekommt Text ({_shown(value)}) statt einer Zahl",
                                      f"{part.value} gets text ({_shown(value)}) instead of a number"), "")
        return _Why("other", T(f"Der Fehler entsteht in {_say(part)}: Dort bekommt die Funktion einen Wert der falschen Art",
                               f"The error arises in {_say(part)}: the function gets a value of the wrong kind there"), "")
    return _Why("other", T("Irgendwo rechnet die Formel mit Text oder mit einem ganzen Bereich statt einer Zelle",
                           "Somewhere the formula calculates with text or with a whole range instead of a single cell"), "")


def _number(calc, part: Node | None, node: Node | None = None) -> _Why:
    for call in ([part] if part is not None else []) + list(_calls(node)):
        why = _domain(calc, call)
        if why is not None:
            return why
    if part is not None:
        return _Why("number", T(f"Der Fehler entsteht in {_say(part)}: Dort bekommt eine Funktion eine Zahl, mit der sie "
                                "nicht rechnen kann",
                                f"The error arises in {_say(part)}: a function gets a number there that it cannot "
                                "calculate with"), "")
    return _Why("number", T("Eine Funktion in der Formel bekommt eine Zahl, mit der sie nicht rechnen kann",
                            "A function in the formula gets a number it cannot calculate with"), "")


def _calls(node: Node | None):
    """Die Funktionsaufrufe und Potenzen in node, die innersten zuerst."""
    if node is None:
        return
    for child in node.args:
        yield from _calls(child)
    if node.kind == "call" or (node.kind == "op" and node.value == "^"):
        yield node


def _domain(calc, call: Node) -> _Why | None:  # noqa: C901 - eine Regel je Funktion
    """Warum call #ZAHL! ergibt, wenn es an einer Zahl liegt, mit der die Funktion nicht rechnen kann, sonst None."""
    if not call.args:
        return None
    result = _value(calc, call)
    if result is not _UNKNOWN and not _is(result, "#NUM!"):
        return None  # Dieser Teil rechnet, der Fehler kommt woanders her.
    name, first = call.value, _value(calc, call.args[0])
    second = _value(calc, call.args[1]) if len(call.args) > 1 else _UNKNOWN
    number = isinstance(first, float) and not isinstance(first, bool)
    if call.kind == "op" or name == "POWER":
        if number and isinstance(second, float) and first < 0 and second != int(second):
            return _Why("number", T(f"{_say(call)} zieht eine Wurzel aus einer negativen Zahl ({_shown(first)})",
                                    f"{_say(call)} takes a root of a negative number ({_shown(first)})"), "")
        if number and isinstance(second, float) and first == 0 and second == 0:
            return _Why("number", T(f"{_say(call)} rechnet 0 hoch 0, das gibt es in Excel nicht",
                                    f"{_say(call)} calculates 0 to the power of 0, which Excel does not allow"), "")
        return None
    german = _fn(name)
    if name == "SQRT" and number and first < 0:
        return _Why("number", T(f"WURZEL bekommt {_shown(first)}, eine negative Zahl", f"SQRT gets {_shown(first)}, a negative number"),
                    "")
    if name in ("LN", "LOG", "LOG10") and number and first <= 0:
        return _Why("number", T(f"{german} bekommt {_shown(first)}: Den Logarithmus gibt es nur für Zahlen über 0",
                                f"{german} gets {_shown(first)}: the logarithm exists only for numbers above 0"), "")
    if name in ("ASIN", "ACOS") and number and not -1 <= first <= 1:
        return _Why("number", T(f"{german} bekommt {_shown(first)}, rechnet aber nur mit Zahlen von -1 bis 1",
                                f"{german} gets {_shown(first)}, but only works with numbers from -1 to 1"), "")
    if name == "FACT" and number and first < 0:
        return _Why("number", T(f"FAKULTÄT bekommt {_shown(first)}, eine negative Zahl", f"FACT gets {_shown(first)}, a negative number"),
                    "")
    if name in ("CEILING", "FLOOR") and number and isinstance(second, float) and first > 0 > second:
        return _Why("number", T(f"{german} bekommt eine positive Zahl ({_shown(first)}) und ein negatives Vielfaches "
                                f"({_shown(second)})",
                                f"{german} gets a positive number ({_shown(first)}) and a negative multiple "
                                f"({_shown(second)})"),
                    T(f"Nimm {_shown(-second)} als Vielfaches.", f"Use {_shown(-second)} as the multiple."))
    if name == "DATE" and number and not 0 <= first <= 9999:
        return _Why("number", T(f"DATUM bekommt das Jahr {_shown(first)}, Excel kennt nur die Jahre 1900 bis 9999",
                                f"DATE gets the year {_shown(first)}, but Excel only knows the years 1900 to 9999"), "")
    if name == "DATE" and number and len(call.args) == 3:
        month, day = second, _value(calc, call.args[2])
        if isinstance(month, float) and isinstance(day, float):
            sep = _sep()
            return _Why("number", T(f"DATUM({_shown(first)}; {_shown(month)}; {_shown(day)}) läge vor dem 1.1.1900, "
                                    "dem ersten Tag, den Excel kennt",
                                    f"DATE({_shown(first)}{sep} {_shown(month)}{sep} {_shown(day)}) would fall before "
                                    "January 1, 1900, the first day Excel knows"),
                        T("Prüf Monat und Tag: Negative Werte zählen zurück.", "Check month and day: negative values count backwards."))
    if name == "DATEDIF" and len(call.args) >= 2 and number and isinstance(second, float) and first > second:
        return _Why("number", T(f"DATEDIF bekommt das Enddatum ({_date_text(second, False)}) vor dem Anfangsdatum "
                                f"({_date_text(first, False)})",
                                f"DATEDIF gets the end date ({_date_text(second, False)}) before the start date "
                                f"({_date_text(first, False)})"),
                    T("Prüf, ob die Daten vertauscht oder falsch eingetragen sind.",
                      "Check whether the dates are swapped or entered incorrectly."))
    if name in ("SMALL", "LARGE") and len(call.args) == 2:
        count = _count(calc, call.args[:1])
        if isinstance(second, float) and (second < 1 or second > count):
            return _Why("number", T(f"{german} fragt nach dem {_shown(second)}. Wert, in {_say(call.args[0])} stehen "
                                    f"aber {count} Zahlen",
                                    f"{german} asks for value number {_shown(second)}, but {_say(call.args[0])} contains "
                                    f"{count} {'number' if count == 1 else 'numbers'}"), "")
    return None


# --- #NAME? ---


def _outside(formula: str) -> str:
    """Die Formel, in der Texte, Blattnamen in Hochkommas und [Spalten] von Tabellen durch Leerzeichen ersetzt sind
    (so bleiben die Stellen dieselben)."""
    out = "".join(p if i % 2 == 0 else " " * len(p) for i, p in enumerate(_QUOTED.split(formula)))
    while True:
        new = re.sub(r"\[[^\[\]]*\]", lambda m: " " * len(m.group(0)), out)
        if new == out:
            return out
        out = new


def _name(sheet: Sheet, calc, pos: Pos, node: Node | None) -> _Why:
    formula = sheet.cells[pos].formula or ""
    outside = _outside(formula)
    if _CURLY.search(outside):
        fixed = "".join('"' if _CURLY.fullmatch(ch) and outside[i] == ch else ch for i, ch in enumerate(formula))
        return _Why("name-quotes", T("In der Formel stehen typografische Anführungszeichen („…“), Excel braucht für Text "
                                     "gerade: \"…\"",
                                     "The formula contains typographic quotation marks („…“ or “…”); Excel needs straight "
                                     "ones for text: \"…\""), "", fixed=fixed)
    calls = [(m.start(1), m.end(1), m.group(1)) for m in _CALL.finditer(outside)]
    strange = [(a, b, n) for a, b, n in calls if n.upper() not in _ENGLISH and n.upper() not in ("NA",)]
    if strange:
        start, end, typed = strange[0]
        upper = typed.upper()
        if upper in _FROM_GERMAN:
            english = _FROM_GERMAN[upper]
            text = T(f"In der Datei steht {upper} mit dem deutschen Namen. Gespeicherte Formeln liest Excel nur mit den "
                     f"englischen Namen (hier {english}), darum erkennt es {upper} hier nicht",
                     f"The file contains {upper} with its German name. Excel reads saved formulas only with the English "
                     f"names (here {english}), so it does not recognize {upper} here")
            return _Why("name-german", text, T("Neu eingegeben rechnet sie: Einsetzen macht das, wie Doppelklick in die "
                                               "Zelle und Enter.",
                                               "Entered again, it calculates: “Apply” does that, like double-clicking the "
                                               "cell and pressing Enter."), fixed=formula[:start] + english + formula[end:])
        best = _suggestion(upper)
        if best is None:
            return _Why("name-unknown", T(f"Excel kennt keine Funktion „{typed}“", f"Excel does not know a function “{typed}”"),
                        T("Ist es ein Tippfehler? Kommt die Funktion aus einem Makro oder Add-in, ist es hier nicht geladen.",
                          "Is it a typo? If the function comes from a macro or add-in, that is not loaded here."))
        english, said = best
        return _Why("name-typo", T(f"Excel kennt keine Funktion „{typed}“. Gemeint ist wohl {said}",
                                   f"Excel does not know a function “{typed}”. You probably mean {said}"), "",
                    fixed=formula[:start] + english + formula[end:])
    lets = any(n.upper() in ("LET", "LAMBDA") for _, _, n in calls)
    bare = [] if lets else [m for m in _BARE.finditer(outside)
                            if not _CELL.fullmatch(m.group(1)) and m.group(1).upper() not in ("TRUE", "FALSE", "WAHR", "FALSCH")
                            and not _is_call(outside, m.end(1))]
    if bare:
        match = bare[0]
        word = match.group(1)
        joined = _NO_COLON.fullmatch(word)
        if joined:
            a, b = joined.group(1).upper(), joined.group(2).upper()
            return _Why("name-colon", T(f"Zwischen {a} und {b} fehlt der Doppelpunkt: {word} ist für Excel ein unbekannter Name",
                                        f"The colon between {a} and {b} is missing: for Excel, {word} is an unknown name"),
                        "", fixed=formula[:match.start(1)] + f"{a}:{b}" + formula[match.end(1):])
        return _Why("name-text", T(f"„{word}“ steht ohne Anführungszeichen: Dann sucht Excel einen Namen „{word}“ und findet "
                                   f"keinen. Als Text gehört es in Anführungszeichen: \"{word}\"",
                                   f"“{word}” has no quotation marks: then Excel looks for a name “{word}” and finds none. "
                                   f"As text, it belongs in quotation marks: \"{word}\""), "",
                    fixed=formula[:match.start(1)] + f'"{word}"' + formula[match.end(1):])
    value = _value(calc, node) if node is not None else _UNKNOWN
    if value is not _UNKNOWN and not isinstance(value, Err) and not _unclear(node):
        # text: was die Formel nachgerechnet ergibt (für den Hinweis zu einer Zelle).
        return _Why("name-reenter", _shown(value), "", fixed=_cleaned(formula))
    newest = _since(formula)
    return _Why("name-other", T("Excel erkennt einen Namen in der Formel nicht", "Excel does not recognize a name in the formula"),
                T("Prüf, ob alle Funktionen richtig geschrieben sind und Text in Anführungszeichen steht. Ist alles "
                  "richtig, gib die Formel neu ein (Doppelklick in die Zelle, dann Enter).",
                  "Check that all functions are spelled correctly and that text is in quotation marks. If everything "
                  "is right, enter the formula again (double-click the cell, then Enter).") + (f" {newest}." if newest else ""))


def _is_call(outside: str, end: int) -> bool:
    return outside[end:].lstrip().startswith("(")


def _suggestion(typed: str) -> tuple[str, str] | None:
    """Der bekannte Funktionsname, der typed am nächsten ist: (englisch, wie er im Hinweis heißt) oder None. Ähnelt
    typed dem englischen Namen, steht im Hinweis der deutsche mit dem englischen dahinter („SVERWEIS (englisch
    VLOOKUP)“), sonst nur der deutsche."""
    names = {**{e: e for e in _ENGLISH}, **_FROM_GERMAN}
    close = difflib.get_close_matches(typed, list(names), n=1, cutoff=0.75)
    if not close:
        return None
    english = names[close[0]]
    german = _german(english)
    if _in_english():
        return english, english
    if close[0] == english and english != german and close[0] not in _FROM_GERMAN:
        return english, f"{german} (englisch {english})"
    return english, german


def _since(formula: str) -> str | None:
    """„XVERWEIS gibt es erst ab Excel 2021“ für die erste neue Funktion in der Formel."""
    for match in _CALL.finditer(_outside(formula)):
        name = match.group(1).upper()
        if name in _SINCE:
            return T(f"{_german(name)} gibt es {_SINCE[name][0]}", f"{name} is available {_SINCE[name][1]}")
    return None


def _unrecognized(formula: str) -> bool:
    """So zeigt Excel (Office.js) eine Funktion, die es in dieser Zelle nicht erkannt hat: =@STABW.S(…)."""
    return re.search(r"@\s*(?:_xl(?:fn|ws)\.)?[A-Za-z_][A-Za-z0-9_.]*\s*\(", _outside(formula)) is not None


def _unclear(node: Node | None) -> bool:
    return node is None or any(item.kind == "unknown" for item in _walk(node))


def _cleaned(formula: str) -> str:
    """Die Formel ohne „_xlfn.“ und „@“, so wie man sie neu eintippt."""
    return "".join(p if i % 2 else re.sub(r"_xl(?:fn|ws)\.", "", p).replace("@", "")
                   for i, p in enumerate(_QUOTED.split(formula)))


# --- Gruppen und Texte ---


def _owners(sheet: Sheet, calc, findings: list[Finding], errors: dict[Pos, str]) -> dict[Pos, Finding]:
    """Zellen mit Fehler, deren Fehler die Korrektur eines anderen Hinweises beseitigt. Nur Hinweise auf einen Fehler:
    Eine Korrektur, die nur zu prüfen ist, setzt man vielleicht nicht ein."""
    owners: dict[Pos, Finding] = {}
    for item in findings:
        if not item.fix or item.level != "fehler":
            continue
        upstream: dict[Pos, object] = {}
        for where, text in item.fix.get("cells", {}).items():
            pos = formel.position(where)
            if isinstance(text, str) and text.startswith("="):
                try:
                    value = calc.run(formel.parse(text[1:]))
                except _FAILS:
                    continue
            else:
                value = text
            if pos not in errors:
                # Die Korrektur steht weiter oben (eine Formel als Text, eine Zahl als Text): Was die Zellen mit dem
                # Fehler damit rechnen, sagt der Hinweis dort („Darum ergibt B6 #WERT! statt 1584“).
                if sheet.dependents.get(pos) and not isinstance(value, Err) and value != "":
                    upstream[pos] = value
                continue
            if pos not in owners and not isinstance(value, Err) and isinstance(text, str) and text.startswith("="):
                owners[pos] = item
        if upstream:
            direct = {p for pos in upstream for p in sheet.dependents.get(pos, ())}
            for pos, _, after in changed_by(sheet, direct, upstream):
                if pos in errors and pos not in owners and not isinstance(after, Err):
                    owners[pos] = item
    return owners


def _grouped(sheet: Sheet, whys: dict[Pos, _Why], errors: dict[Pos, str], owners: dict[Pos, Finding]) -> list[_Group]:
    """Heruntergezogene Zellen mit demselben Fehler aus demselben Grund: ein Hinweis. #NAME?, weil Excel die Formeln
    beim Öffnen nicht erkannt hat: ein Hinweis für das ganze Blatt."""
    groups: list[_Group] = []
    again: _Group | None = None
    last: tuple | None = None
    for pos in sorted(whys, key=lambda p: (p[1], p[0])):
        why = whys[pos]
        if why.kind == "name-reenter":
            if again is None:
                again = _Group([], [])
                groups.append(again)
            again.cells.append(pos)
            again.whys.append(why)
            continue
        owner = owners.get(pos) if why.kind != "name-reenter" else None
        cell = sheet.cells[pos]
        shape = (pos[1], errors[pos], why.kind, id(owner), pattern(cell.formula, pos) if cell.formula else None, why.there)
        if last is not None and shape == last[0] and pos == (last[1][0] + 1, last[1][1]):
            groups[-1].cells.append(pos)
            groups[-1].whys.append(why)
        else:
            groups.append(_Group([pos], [why], owner))
        last = (shape, pos)
    groups.sort(key=lambda g: (g.cells[0][0], g.cells[0][1]))
    return groups


def _span(cells: list[Pos]) -> str:
    """„E3“, „E3 bis E8“ oder „H3, H6 und H9“."""
    if len(cells) == 1:
        return address(cells[0])
    if _below(cells) and len(cells) > 2:
        return T(f"{address(cells[0])} bis {address(cells[-1])}", f"{address(cells[0])} to {address(cells[-1])}")
    return _listed(cells)


def _below(cells: list[Pos]) -> bool:
    return all(b == (a[0] + 1, a[1]) for a, b in zip(cells, cells[1:], strict=False))


def _listed(cells: list[Pos]) -> str:
    names = [address(p) for p in cells[:MAX_LISTED]]
    if len(cells) > MAX_LISTED:
        return ", ".join(names) + T(f" und {len(cells) - MAX_LISTED} weiteren", f" and {len(cells) - MAX_LISTED} more")
    return names[0] if len(names) == 1 else ", ".join(names[:-1]) + T(f" und {names[-1]}", f" and {names[-1]}")


def _follows(cells: list[Pos]) -> str:
    """„ Daraus folgt der Fehler auch in E9 und E10.“"""
    return T(f" Daraus folgt der Fehler auch in {_listed(cells)}.", f" The error carries over to {_listed(cells)} as well.")


def _area(cells: list[Pos]) -> str:
    """Die Zellen zum Markieren: „E3:E8“ oder „H3,H6“."""
    if len(cells) > 1 and _below(cells):
        return f"{address(cells[0])}:{address(cells[-1])}"
    return ",".join(address(p) for p in cells[:50])


def _sentence(text: str) -> str:
    return text[:1].upper() + text[1:]


def _detail(sheet: Sheet, group: _Group, shown: dict[Pos, str]) -> str:
    """„E3 zeigt #DIV/0!: Die Formel teilt …“, für mehrere Zellen „E3 bis E8 zeigen #DIV/0!. Bei E3: …“."""
    first = group.cells[0]
    if len(group.cells) == 1:
        return T(f"{sheet.figure_label(first)} zeigt {shown[first]}: {_sentence(group.why.text)}.",
                 f"{sheet.figure_label(first)} shows {shown[first]}: {_sentence(group.why.text)}.")
    many = {why.many for why in group.whys}
    if len(many) == 1 and group.why.many:
        items = list(dict.fromkeys(why.item for why in group.whys))
        said = items[0] if len(items) == 1 else _joined(items)
        return T(f"{_span(group.cells)} zeigen {shown[first]}: {group.why.many.format(said)}.",
                 f"{_span(group.cells)} show {shown[first]}: {group.why.many.format(said)}.")
    return T(f"{_span(group.cells)} zeigen {shown[first]}. Bei {address(first)}: {_sentence(group.why.text)}; in den "
             "Zeilen darunter ist es genauso.",
             f"{_span(group.cells)} show {shown[first]}. In {address(first)}: {_sentence(group.why.text)}; it is the "
             "same in the rows below.")


def _joined(items: list[str]) -> str:
    """„„C“, „D“ und „E““, bei vielen nur die ersten."""
    if len(items) > MAX_LISTED:
        return ", ".join(items[:MAX_LISTED]) + T(f" und {len(items) - MAX_LISTED} weitere", f" and {len(items) - MAX_LISTED} more")
    return ", ".join(items[:-1]) + T(f" und {items[-1]}", f" and {items[-1]}")


def _along(sheet: Sheet, group: _Group, shown: dict[Pos, str]) -> str:
    """Was ein anderer Hinweis mitsagt, wenn seine Korrektur den Fehler behebt."""
    text = _detail(sheet, group, shown) + T(" Die Korrektur behebt das mit.", " The fix takes care of this as well.")
    if group.after:
        text += _follows(group.after)
    return text


def _finding(sheet: Sheet, group: _Group, shown: dict[Pos, str], fresh) -> Finding:
    if group.why.kind == "name-reenter":
        return _reenter(sheet, group, shown, fresh)
    first = group.cells[0]
    text = _detail(sheet, group, shown)
    if group.after:
        text += _follows(group.after)
    fix = None
    if all(why.fixed is not None for why in group.whys):
        fix = einsetzen.formulas(sheet, {pos: why.fixed for pos, why in zip(group.cells, group.whys, strict=True)},
                                 calc=fresh if group.why.kind.startswith("name") else None)
    tip = group.why.tip
    if len(group.cells) > 1 and group.why.many_tip and len({why.item for why in group.whys}) > 1:
        tip = group.why.many_tip
    if fix is not None:
        where = f" in {address(first)}" if len(group.cells) > 1 else ""
        down = (T(f", dann bis {address(group.cells[-1])} herunterziehen", f", then fill down to {address(group.cells[-1])}")
                if len(group.cells) > 1 else "")
        calc = fresh if group.why.kind.startswith("name") else sheet.evaluator()[0]
        value = _value(calc, formel.parse(fix["cells"][address(first)][1:])) if len(group.cells) == 1 else None
        said = (T(f", nachgerechnet {_shown(value)}", f", recalculated {_shown(value)}")
                if isinstance(value, float) and not isinstance(value, bool) else "")
        tip = T(f"Korrigiert{where}: {fix['german']}{said}{down}.", f"Corrected{where}: {fix['german']}{said}{down}.") + (
            f" {tip}" if tip else "")
    return Finding(f"error:{address(first)}", text, "fehler", tip, fix=fix, cell=_area(group.cells))


def _reenter(sheet: Sheet, group: _Group, shown: dict[Pos, str], fresh) -> Finding:
    """#NAME?, obwohl die Formeln richtig geschrieben sind: Excel hat sie beim Öffnen nicht erkannt."""
    cells, first = group.cells, group.cells[0]
    if len(cells) == 1:
        text = T(f"{sheet.figure_label(first)} zeigt {shown[first]}, obwohl die Formel richtig geschrieben ist: "
                 f"Nachgerechnet ergibt sie {group.why.text}.",
                 f"{sheet.figure_label(first)} shows {shown[first]} although the formula is spelled correctly: "
                 f"recalculated, it gives {group.why.text}.")
    else:
        text = T(f"{_span(cells)} zeigen {shown[first]}, obwohl die Formeln richtig geschrieben sind.",
                 f"{_span(cells)} show {shown[first]} although the formulas are spelled correctly.")
    text += T(" Excel hat sie beim Öffnen nicht erkannt: Die Datei kommt wohl aus einem anderen Programm, das neue "
              "Funktionen ohne den Zusatz speichert, den Excel dafür braucht.",
              f" Excel did not recognize {'it' if len(cells) == 1 else 'them'} when opening the file: it probably comes "
              "from another program that saves newer functions without the prefix Excel needs for them.")
    if any(_unrecognized(sheet.cells[p].formula or "") for p in cells):
        text += T(" Daher kommt auch das „@“ vor dem Funktionsnamen.",
                  " That is also where the “@” before the function name comes from.")
    if group.after:
        text += _follows(group.after)
    fix = einsetzen.formulas(sheet, {p: why.fixed for p, why in zip(cells, group.whys, strict=True)}, calc=fresh)
    tip = T("Einsetzen gibt die Formeln neu ein, wie Doppelklick in die Zelle und Enter. Für ein ganzes Blatt geht auch: "
            "Suchen und Ersetzen, „=“ durch „=“ ersetzen.",
            "“Apply” enters the formulas again, like double-clicking the cell and pressing Enter. For a whole sheet, "
            "this also works: Find and Replace, replace “=” with “=”.") if fix else T(
           "Gib die Formeln neu ein: Doppelklick in die Zelle, dann Enter. Für ein ganzes Blatt geht auch: Suchen und "
           "Ersetzen, „=“ durch „=“ ersetzen.",
           "Enter the formulas again: double-click the cell, then Enter. For a whole sheet, this also works: Find and "
           "Replace, replace “=” with “=”.")
    newest = next((s for s in (_since(sheet.cells[p].formula or "") for p in cells) if s), None)
    if newest:
        tip += T(f" Bleibt danach #NAME?, kennt diese Excel-Version die Funktion noch nicht: {newest}.",
                 f" If #NAME? remains afterwards, this Excel version does not know the function yet: {newest}.")
    return Finding(f"error:{address(first)}", text, "fehler", tip, fix=fix, cell=_area(cells))
