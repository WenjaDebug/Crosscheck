"""Ausreißer innerhalb einer Gruppe: Laborwerte im Langformat mischen Kalium um 4 und Natrium um 140 in einer Spalte.
Kalium 45 fällt nur gegen die übrigen Kaliumwerte auf. Steht in der Zeile dazu eine andere Einheit, die sich nur in der
Vorsilbe unterscheidet (g/l statt g/dl), rechnet der Hinweis genau um und Einsetzen schreibt Wert und Einheit."""

import random

from test_workbook import xlsx

from voiceagent import pruefer
from voiceagent.gruppen import _factor

PARAMS = [("Hämoglobin", "g/dl", 11.5, 16.5, 1), ("Kreatinin", "µmol/l", 55, 115, 0), ("Kalium", "mmol/l", 3.4, 5.1, 1),
          ("Natrium", "mmol/l", 134, 146, 0), ("Leukozyten", "G/l", 3.8, 11.0, 1), ("CRP", "mg/l", 0.5, 45, 1),
          ("Thrombozyten", "G/l", 140, 420, 0)]


def _labor(seed: int = 1, **changes) -> dict:
    rng = random.Random(seed)
    cells = {"A1": "Patient", "B1": "Parameter", "C1": "Wert", "D1": "Einheit"}
    for i in range(210):
        r = i + 2
        name, unit, low, high, digits = PARAMS[i % len(PARAMS)]
        value = round(rng.lognormvariate(1.3, 1.0), 1) if name == "CRP" else round(rng.uniform(low, high), digits)
        cells.update({f"A{r}": f"P{100 + i // 7:03d}", f"B{r}": name, f"C{r}": float(value), f"D{r}": unit})
    cells.update(changes)
    return cells


def _hints(cells: dict) -> dict:
    done = pruefer.Checker(None).check_file("labor.xlsx", xlsx(("Labor", cells)))
    assert done["state"] == "ok", done
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


def test_a_value_ten_times_off_its_parameter_gets_the_right_number():
    # Zeile 18 ist Kalium, 26 Natrium, 37 Hämoglobin: gegen die ganze Spalte fällt keiner auf.
    found = _hints(_labor(C18=45.0, C26=14.1, C37=142.0))
    assert set(found) == {"scale:C18", "scale:C26", "scale:C37"}
    kalium = found["scale:C18"]
    assert "C18 (Patient P102) = 45 liegt weit weg von den übrigen Werten in Wert, wo Parameter „Kalium“ ist" in kalium["text"]
    assert "Durch 10 geteilt passt es genau dazu" in kalium["text"] and kalium["fix"]["cells"] == {"C18": 4.5}
    assert found["scale:C26"]["fix"]["cells"] == {"C26": 141.0}


def test_another_unit_in_the_row_is_converted_exactly():
    found = _hints(_labor(C37=142.0, D37="g/l"))
    hint = found["rowunit:C37"]
    assert "In D37 steht dazu „g/l“, sonst „g/dl“. 142 g/l sind 14,2 g/dl." in hint["text"]
    assert hint["fix"]["cells"] == {"C37": 14.2, "D37": "g/dl"} and hint["fix"]["was"]["D37"] == {"text": "g/l"}
    # Die Einheit allein auf „g/dl“ zu setzen, machte daraus 142 g/dl: kein eigener Hinweis dazu.
    assert not [k for k in found if k.endswith("D37")]


def test_a_unit_with_another_factor_is_only_named():
    """Kreatinin in mg/dl: Der Faktor ist 88,4, keine Zehnerpotenz. Mal 100 wäre falsch."""
    hint = _hints(_labor(C52=0.9, D52="mg/dl"))["rowunit:C52"]
    assert "In D52 steht dazu „mg/dl“, sonst „µmol/l“" in hint["text"] and hint.get("fix") is None


def test_a_clean_lab_list_and_a_loose_grouping_stay_quiet():
    assert _hints(_labor(2)) == {}
    rng = random.Random(7)
    cells = {"A1": "Fall", "B1": "Station", "C1": "Kosten"}
    for i in range(200):
        r = i + 2
        cells.update({f"A{r}": f"F{i}", f"B{r}": ["Kardio", "Chir", "Neuro", "Onko"][i % 4],
                      f"C{r}": round(rng.lognormvariate(9 + (i % 4) * 0.2, 0.7), 2)})
    assert not [k for k in _hints(cells) if k.startswith(("outlier:", "scale:", "rowunit:"))]


def test_prices_per_article():
    rng = random.Random(6)
    cells = {"A1": "Bestellung", "B1": "Artikel", "C1": "Menge", "D1": "Preis CHF"}
    articles = [("Spritze 5 ml", 0.35), ("Infusion NaCl", 2.8), ("Verband", 4.5), ("Katheter", 19.9), ("Pumpe", 1450.0)]
    for i in range(150):
        r, (article, price) = i + 2, articles[i % 5]
        cells.update({f"A{r}": f"B{1000 + i}", f"B{r}": article, f"C{r}": float(rng.randint(1, 50)),
                      f"D{r}": round(price * rng.uniform(0.95, 1.05), 2)})
    cells["D20"] = 1990.0
    assert _hints(cells)["scale:D20"]["fix"]["cells"] == {"D20": 19.9}


def test_units_that_differ_only_in_their_prefix():
    assert _factor("g/l", "g/dl") == 0.1 and _factor("mmol/l", "µmol/l") == 1000.0 and _factor("kg", "g") == 1000.0
    assert _factor("G/l", "/nl") == 1.0 and _factor("mmol/L", "mmol/l") == 1.0 and _factor("U/l", "mU/l") == 1000.0
    assert _factor("mg/dl", "µmol/l") is None and _factor("%", "g/l") is None and _factor("G/l", "g/l") is None
    assert _factor("mmHg", "kPa") is None and _factor("°C", "K") is None
