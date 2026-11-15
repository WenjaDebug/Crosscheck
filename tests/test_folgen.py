"""Eine Folge beim Ursprung: 1,74 in Größe (cm) macht den BMI derselben Zeile unmöglich. Gemeldet wird die Größe, mit
174 zum Einsetzen und dem BMI als Folge, nicht dazu noch der BMI für sich. Ein Wert mal 10, 100 oder 1000 daneben
(ein verrutschtes Komma, m statt cm, g statt kg) bekommt die richtige Zahl, wo genau ein Faktor passt."""

from test_workbook import xlsx

from voiceagent import pruefer


def _hints(cells: dict) -> dict:
    done = pruefer.Checker(None).check_file("studie.xlsx", xlsx(("Daten", cells)))
    assert done["state"] == "ok", done
    return {h["key"]: h for s in done["sheets"] for h in s["hints"]}


def _studie(**changes) -> dict:
    cells = {f"{c}1": h for c, h in zip("ABCDE", ["Proband", "Alter", "Gewicht (kg)", "Größe (cm)", "BMI"])}
    for r in range(2, 32):
        cells.update({f"A{r}": f"S{r:03d}", f"B{r}": float(25 + (r * 7) % 40), f"C{r}": float(60 + (r * 11) % 30),
                      f"D{r}": float(162 + (r * 13) % 28)})
    cells.update(changes)
    for r in range(2, 32):
        weight, height = cells[f"C{r}"], cells[f"D{r}"]
        cells[f"E{r}"] = ("=", f"C{r}/(D{r}/100)^2", weight / (height / 100) ** 2)
    return cells


def test_a_height_in_metres_is_the_error_and_the_bmi_its_consequence():
    found = _hints(_studie(D10=1.74))
    hint = found["scale:D10"]
    assert hint["level"] == "fehler" and hint["fix"]["cells"] == {"D10": 174.0}
    assert "D10 (Größe (cm) S010) = 1,74 liegt weit weg von den übrigen Werten in Größe (cm)" in hint["text"]
    assert "Mal 100 passt es genau dazu, wohl ein verrutschtes Komma oder eine andere Einheit" in hint["text"]
    assert "Darum ergibt E10 (BMI S010)" in hint["text"] and hint["text"].endswith("Richtig wäre 174 statt 1,74.")
    # Der BMI ist nur die Folge: kein eigener Hinweis, auch kein Ausreißer.
    assert not any(k.endswith("E10") or k.startswith(("impossible:E", "outlier:D10")) for k in found)


def test_grams_instead_of_kilograms_are_divided():
    hint = _hints(_studie(C12=72000.0))["scale:C12"]
    assert "Durch 1000 geteilt passt es genau dazu, wohl" in hint["text"] and hint["fix"]["cells"] == {"C12": 72.0}


def test_an_impossible_input_names_the_impossible_result_once():
    found = _hints(_studie(C20=-99.0))
    hint = found["impossible:C20"]
    assert "In Gewicht (kg) steht -99 in C20" in hint["text"] and "Darum ist auch E20 (BMI S020)" in hint["text"]
    assert not any(k.startswith(("impossible:E", "outlier:C20")) for k in found)


def test_a_wide_column_keeps_the_plain_outlier():
    """Beträge von 100 bis 5000: Ob 1,5 mal 100 oder mal 1000 gemeint ist, sagt die Spalte nicht."""
    cells = {"A1": "Fall", "B1": "Kosten"}
    for r in range(2, 32):
        cells.update({f"A{r}": f"F{r}", f"B{r}": float(100 + (r * 397) % 4900)})
    cells["B9"] = 1.5
    found = _hints(cells)
    assert not any(k.startswith("scale:") for k in found)


def test_a_clean_study_has_nothing_to_say():
    assert _hints(_studie()) == {}
