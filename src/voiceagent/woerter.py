"""Wörter in Spaltenköpfen und Beschriftungen auf Englisch, Französisch und Italienisch, zu den deutschen, die die
Prüfungen selbst kennen. Spitäler in der Schweiz führen ihre Listen auch auf Französisch oder Italienisch, Exporte
kommen oft auf Englisch: „Somme“, „Totale“ und „Date de sortie“ sollen so viel sagen wie „Summe“ und „Austritt“.

Nur Wörter, die in allen vier Sprachen dasselbe meinen, machen Hinweise lauter. „Media“ heißt auf Italienisch
Mittelwert, auf Englisch Medien, und „Part“ auf Französisch Anteil, auf Englisch ein Bauteil: Solche Wörter machen
Hinweise nur leiser (eine Spalte von Mittelwerten zählt in keiner Summe mit), oder sie zählen erst, wenn die Zahlen
dazu passen (Anteile zwischen 0 und 1)."""

import re
import unicodedata

# Eine Zeile oder Kennzahl über alles: „Somme“, „Totale“, „Total général“.
TOTAL = {"somme", "sommes", "somma", "totale", "totali", "totaux", "complessivo", "complessiva", "complessivi",
         "subtotale"}
# Wörter, die sagen: über die ganze Spalte, nicht nur eine Gruppe.
ALL = TOTAL | {"tous", "toutes", "tout", "tutti", "tutte", "globale", "global", "général", "generale"}
# Kennzahl-Wörter und was sie rechnen.
STATS = {"somme": "SUM", "somma": "SUM", "totale": "SUM", "moyenne": "AVERAGE", "moyen": "AVERAGE", "moyens": "AVERAGE",
         "medio": "AVERAGE", "médiane": "MEDIAN", "mediane": "MEDIAN", "mediana": "MEDIAN", "nombre": "COUNT",
         "conteggio": "COUNT", "massimo": "MAX", "minimo": "MIN"}
# Spalten mit einer Quote je Zeile: Bruttomarge, EBIT-Marge, Umsatzrendite, Ausfallquote. Ihre Summe sagt nichts.
RATE_ENDINGS = ("marge", "margen", "margin", "margins", "rendite", "quote", "quoten", "marges", "margine", "margini")


def rate_word(words) -> bool:
    """Nennt einer der Wörter eine Quote (Bruttomarge, EBIT-Marge, Umsatzrendite)?"""
    return any(w.endswith(RATE_ENDINGS) for w in words)


# Mittelwerte, auch die mehrdeutigen: Sie machen Hinweise nur leiser.
MEANS = {"moyenne", "moyennes", "moyen", "moyens", "médiane", "mediane", "media", "medie", "medio", "medi", "mediana",
         "mean", "average", "avg"}
# Eine Zeile, die so heißt, ist keine Datenzeile. „Media“ kann auf Englisch ein Posten in der Liste sein.
SUMMARY_ROWS = TOTAL | MEANS - {"media"}
# Füllwörter in Beschriftungen („Coût moyen pour les cas“, „Durata media delle degenze“). Keine kurzen wie „de“ oder
# „il“: Als Länderkürzel (DE, IL) sind sie oft selbst die Gruppe, nach der eine Kennzahl fragt.
FILLER = {"les", "pour", "aux", "dans", "sur", "della", "dello", "delle", "degli", "dei", "nella", "nelle"}
# Ein Anteil in Prozent: eindeutig in jeder Sprache.
SHARE = r"pourcentage|percentuale|percentage|proporzione"
# Ein Anteil nur, wenn die Zahlen Brüche zwischen 0 und 1 sind: „Part femmes“, „Quota donne“, „Market share“, aber
# nicht „Part no“ (Bauteil), „Share price“ oder „Sales quota“.
SHARE_IF_FRACTIONS = r"\bpart\b|\bquota\b|\bshare\b|\bproportion\b"
# Eine Zahl, die nie negativ ist.
NOT_NEGATIVE = (r"length of stay|duration|durée|durata|degenza|séjour|soggiorno|\bcosts?\b|coûts?\b|couts?\b|\bcost[oi]\b|"
                r"\bprice\b|\bprix\b|\bprezzo\b|\bheight\b|\btaille\b|altezza|quantity|quantité|quantità|"
                r"number of|nombre de|numero di|\bjours\b|\bgiorni\b|\bdays\b|\blits\b|\bletti\b|\bbeds\b")
# Datumsspalten: Anfang, Ende, Geburt.
START = r"\bentrée\b|\bentree\b|ammissione|ricovero|\bingresso\b|\bdébut\b|\bdebut\b|\binizio\b|\barrivée\b|\barrivo\b"
END = r"\bsortie\b|dimissione|\buscita\b|\bfin\b|\bfine\b|\bdécès\b|\bdeces\b|decesso"
BIRTH = r"naissance|nascita"
START_WORDS = {"entrée", "entree", "ammissione", "ricovero", "ingresso", "début", "debut", "inizio"}
END_WORDS = {"sortie", "dimissione", "uscita", "décès", "deces", "decesso"}
BIRTH_WORDS = {"naissance", "nascita"}
AGE_WORDS = {"âge", "età"}
# Eine Nummer, die einen Fall benennt („N° de cas“, „N. caso“, „Case no“). Eine Person darf in einer Liste von
# Besuchen mehrmals vorkommen, ein Fall oder eine Rechnung nicht.
ID = r"numéro|numero|\bcas\b|\bcaso\b|\bcase\b|dossier|paziente|n°"
PERSON = r"paziente|\bclient|\bcliente"
UNIQUE = r"\bcas\b|\bcaso\b|\bcase\b|facture|fattura|invoice"
# Zahlen, die etwas benennen statt messen (Jahr, Monat, Postleitzahl).
NAMING = r"\bannée\b|\banno\b|\bmois\b|\bmese\b|\bjour\b|\bgiorno\b|semaine|settimana|trimestre|\bnpa\b|code postal|codice"
# Ein Satz, der in der ganzen Spalte gleich sein sollte (MwSt).
RATE = r"\btaux\b|\btasso\b|aliquota|\btva\b|\biva\b"
# Eine Veränderung darf negativ sein und über 100 % gehen.
CHANGES = ("variation", "variazione", "évolution", "croissance", "crescita", "écart", "ecart", "scarto")
# Text, der in einer Zahlenspalte für „fehlt“ steht.
PLACEHOLDERS = {"unknown", "missing", "inconnu", "manquant", "mancante", "sconosciuto", "n.d.", "nd"}
# Grenzen, die es nicht gibt, mit dem Wort in den anderen Sprachen.
AGE = r"\bâge\b|\betà\b"
BMI = r"\bimc\b"
TEMPERATURE = r"temp[eé]rature|temperatura"
PULSE = r"\bpouls\b|\bpolso\b|\bpulse\b|fréquence cardiaque|frequenza cardiaca"
SYSTOLIC = r"sistol"


def starts_with(text: str, words) -> bool:
    """„Somme des coûts“ beginnt mit „somme“, „Sommer“ nicht: anders als bei deutschen Wortanfängen („Gesamtsumme“)
    zählt hier nur das ganze Wort."""
    pattern = "|".join(map(re.escape, sorted(words, key=len, reverse=True)))
    return re.match(rf"(?:{pattern})(?!\w)", spaced(text).strip().casefold()) is not None


# Wie Köpfe geschrieben sind: „VerkaufAuto“, „Verkauf_Auto“, „Verkauf.Auto“ und „Verkauf Auto“ sind dasselbe,
# „EBITMarge“ ist „EBIT Marge“, „UmsatzEUR“ ist „Umsatz EUR“, „Umsatz2025“ ist „Umsatz 2025“. Abkürzungen wie MwSt,
# USt oder KPIs bleiben ganz (vor dem Großbuchstaben stehen weniger als zwei Kleinbuchstaben).
_CAMEL = re.compile(r"(?<=[a-zäöüß]{2})(?=[A-ZÄÖÜ](?:[a-zäöüß]{2}|[A-ZÄÖÜ]+(?![a-zäöüß])))"
                    r"|(?<=[A-ZÄÖÜ]{2})(?=[A-ZÄÖÜ][a-zäöüß]{2})"
                    r"|(?<=[a-zäöüß]{3})(?=\d)")
# Punkt nur zwischen längeren Wörtern („Mio.EUR“), nicht in Abkürzungen wie p.a. oder z.B.
_JOINERS = re.compile(r"(?<=[^\W\d_])[_·](?=[^\W\d_])|(?<=[^\W\d_]{2})\.(?=[^\W\d_]{2})")
_SPACES = re.compile(r"[\u00a0\u2007\u202f\t\r\n]+")


def spaced(text) -> str:
    """Der Text mit Leerzeichen zwischen den Wörtern, bevor er klein geschrieben und zerlegt wird. Auch Umlaute,
    die ein Mac als zwei Zeichen speichert (u + ¨), werden zu einem."""
    text = unicodedata.normalize("NFC", str(text or ""))
    text = _SPACES.sub(" ", text)
    text = _JOINERS.sub(" ", text)
    return _CAMEL.sub(" ", text)


def fold(word: str) -> str:
    """Ein Wort zum Vergleichen: klein, ß als ss, Umlaut als ae/oe/ue („Größe“ = „Groesse“ = „GRÖSSE“)."""
    word = unicodedata.normalize("NFC", word).casefold()
    for a, b in (("ß", "ss"), ("ä", "ae"), ("ö", "oe"), ("ü", "ue")):
        word = word.replace(a, b)
    return word
