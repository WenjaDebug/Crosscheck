"""Was der Name sagt, gegen das, was die Formel rechnet (kopf.py, massstab.py): je Fehlerart die Mappe mit dem Fehler,
dieselbe richtig (still) und der Name in anderen Schreibweisen (Leerzeichen, Bindestrich, Groß/klein, Umlaute)."""

from datetime import date

import pytest
from test_dritte_serie import PERCENT, TIME, _hints

MONTHS = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]
SALES = [100900, 80800, 131100, 83800, 95700, 136900, 110700, 132100, 134800, 87800, 132200, 135200]
COSTS = [62400, 54000, 83900, 58800, 65300, 95300, 78500, 99200, 82800, 63400, 100200, 83700]
ROWS = range(5, 17)


def book(heads: dict, row, extra: dict | None = None, name: str = "Blatt") -> tuple:
    """Kopf in Zeile 4, die Monate in A5:A16; row(r, i) gibt die Zellen der Zeile."""
    cells = {"A1": "Auswertung 2026", "A4": "Monat", **{f"{c}4": h for c, h in heads.items()}}
    for i, r in enumerate(ROWS):
        cells[f"A{r}"] = MONTHS[i]
        cells.update(row(r, i))
    cells.update(extra or {})
    return ((name, cells),)


def mine(found: dict) -> dict:
    return {key: h for (_, key), h in found.items() if h["level"] in ("fehler", "pruefen")}


def kinds(found: dict) -> set:
    return {k.split(":")[0] for k in mine(found)}


def sales_costs(r, i):
    return {f"B{r}": SALES[i], f"C{r}": COSTS[i]}


# --- Vorzeichen ---


@pytest.mark.parametrize("head", ["Abweichung (Ist − Plan)", "Abweichung Ist-Plan", "ABW. IST minus PLAN",
                                  "Abweichung (Ist–Plan)", "Abweichung Ist - Plan"])
def test_sign_against_the_name(head):
    heads = {"B": "Ist", "C": "Plan", "D": head}
    wrong = mine(_hints(*book(heads, lambda r, i: {**sales_costs(r, i), f"D{r}": f"=C{r}-B{r}"})))
    assert "sign:D5" in wrong, wrong
    assert wrong["sign:D5"]["fix"]["cells"]["D5"] == "=B5-C5" and wrong["sign:D5"]["fix"]["cells"]["D16"] == "=B16-C16"
    assert "Das Vorzeichen ist vertauscht" in wrong["sign:D5"]["text"]
    assert "sign" not in kinds(_hints(*book(heads, lambda r, i: {**sales_costs(r, i), f"D{r}": f"=B{r}-C{r}"})))


# --- Kosten negativ, Rabatt ---


def test_negative_costs_subtracted_again():
    heads = {"B": "Umsatz", "C": "Kosten", "D": "Gewinn"}
    neg = lambda r, i: {f"B{r}": SALES[i], f"C{r}": -COSTS[i]}  # noqa: E731
    wrong = mine(_hints(*book(heads, lambda r, i: {**neg(r, i), f"D{r}": f"=B{r}-C{r}"})))
    assert wrong["negcost:D5"]["fix"]["cells"]["D5"] == "=B5+C5", wrong
    assert "schon negativ" in wrong["negcost:D5"]["text"]
    assert "negcost" not in kinds(_hints(*book(heads, lambda r, i: {**neg(r, i), f"D{r}": f"=B{r}+C{r}"})))
    # Positive Kosten abgezogen: richtig.
    assert "negcost" not in kinds(_hints(*book(heads, lambda r, i: {**sales_costs(r, i), f"D{r}": f"=B{r}-C{r}"})))


@pytest.mark.parametrize("head", ["Netto nach Rabatt", "Endpreis", "Zahlbetrag", "netto-Betrag"])
def test_discount_added(head):
    heads = {"B": "Listenpreis", "C": "Rabatt (CHF)", "D": head}
    row = lambda r, i: {f"B{r}": SALES[i] / 100, f"C{r}": round(SALES[i] / 1000, 2)}  # noqa: E731
    wrong = mine(_hints(*book(heads, lambda r, i: {**row(r, i), f"D{r}": f"=B{r}+C{r}"})))
    assert wrong["rabatt:D5"]["fix"]["cells"]["D5"] == "=B5-C5", wrong
    assert "rabatt" not in kinds(_hints(*book(heads, lambda r, i: {**row(r, i), f"D{r}": f"=B{r}-C{r}"})))


# --- Brutto ohne MwSt. ---


@pytest.mark.parametrize("head", ["Brutto", "Betrag inkl. MwSt.", "inkl MWST", "Bruttobetrag"])
def test_gross_without_vat(head):
    heads = {"B": "Netto", "D": head}
    extra = {"F2": "MwSt.-Satz", "G2": (0.081, PERCENT)}
    wrong = mine(_hints(*book(heads, lambda r, i: {f"B{r}": SALES[i], f"D{r}": f"=B{r}"}, extra)))
    assert wrong["gross:D5"]["fix"]["cells"]["D5"] == "=B5*(1+$G$2)", wrong
    right = book(heads, lambda r, i: {f"B{r}": SALES[i], f"D{r}": f"=B{r}*(1+$G$2)"}, extra)
    assert "gross" not in kinds(_hints(*right))


# --- Prozent ---


def test_percent_times_hundred_with_percent_format():
    heads = {"B": "Umsatz", "C": "Kosten", "D": "Marge", "E": "Marge in % vom Umsatz"}
    row = lambda r, i, f: {**sales_costs(r, i), f"D{r}": f"=B{r}-C{r}", f"E{r}": (f, PERCENT)}  # noqa: E731
    wrong = mine(_hints(*book(heads, lambda r, i: row(r, i, f"=D{r}/B{r}*100"))))
    assert wrong["percent2:E5"]["fix"]["cells"]["E5"] == "=D5/B5", wrong
    assert "percent2" not in kinds(_hints(*book(heads, lambda r, i: row(r, i, f"=D{r}/B{r}"))))


def test_share_without_percent_format():
    heads = {"B": "Umsatz", "C": "Kosten", "D": "Marge", "E": "Marge (% vom Umsatz)"}
    row = lambda r, i: {**sales_costs(r, i), f"D{r}": f"=B{r}-C{r}", f"E{r}": (f"=D{r}/B{r}", PERCENT)}  # noqa: E731
    extra = {"A18": "Summe", "B18": "=SUM(B5:B16)", "D18": "=SUM(D5:D16)", "A20": "Gesamtmarge (%)",
             "B20": "=D18/B18"}
    wrong = mine(_hints(*book(heads, row, extra)))
    item = wrong["percentfmt:B20"]
    assert item["fix"]["format"] == {"B20": "E5"} and item["fix"]["german"] == "Prozentformat wie E5", item
    assert "percentfmt" not in kinds(_hints(*book(heads, row, {**extra, "B20": ("=D18/B18", PERCENT)})))


# --- Zeit ---


@pytest.mark.parametrize("head", ["Dauer (min)", "Dauer in Minuten", "Minuten", "Dauer (Min.)"])
def test_time_difference_in_the_wrong_unit(head):
    heads = {"B": "Beginn", "C": "Ende", "D": head}
    row = lambda r, i, f: {f"B{r}": (0.3 + i / 100, TIME), f"C{r}": (0.5 + i / 100, TIME), f"D{r}": f}  # noqa: E731
    wrong = mine(_hints(*book(heads, lambda r, i: row(r, i, f"=(C{r}-B{r})*24"))))
    assert wrong["timeunit:D5"]["fix"]["cells"]["D5"] == "=(C5-B5)*1440", wrong
    assert "soll Minuten zeigen, die Formel ergibt aber Stunden" in wrong["timeunit:D5"]["text"]
    assert "timeunit" not in kinds(_hints(*book(heads, lambda r, i: row(r, i, f"=(C{r}-B{r})*1440"))))


# --- Monat und Jahr ---


@pytest.mark.parametrize("year,month", [("Miete pro Jahr", "Miete pro Monat"), ("Jahresmiete", "Miete/Monat"),
                                        ("Miete jährlich", "monatliche Miete"), ("Miete p.a.", "Miete mtl.")])
def test_monthly_from_yearly(year, month):
    heads = {"B": year, "C": month}
    wrong = mine(_hints(*book(heads, lambda r, i: {f"B{r}": SALES[i], f"C{r}": f"=B{r}"})))
    assert wrong["period:C5"]["fix"]["cells"]["C5"] == "=B5/12", wrong
    assert "period" not in kinds(_hints(*book(heads, lambda r, i: {f"B{r}": SALES[i], f"C{r}": f"=B{r}/12"})))


# --- Index, Fenster ---


@pytest.mark.parametrize("head", ["Index (Jan = 100)", "Umsatz-Index Jan=100", "Index (JAN = 100)"])
def test_index_on_the_wrong_base(head):
    heads = {"B": "Umsatz", "H": head}
    wrong = mine(_hints(*book(heads, lambda r, i: {f"B{r}": SALES[i], f"H{r}": f"=B{r}/$B$6*100"})))
    assert wrong["indexbase:H5"]["fix"]["cells"]["H5"] == "=B5/$B$5*100", wrong
    assert "indexbase" not in kinds(_hints(*book(heads, lambda r, i: {f"B{r}": SALES[i], f"H{r}": f"=B{r}/$B$5*100"})))


@pytest.mark.parametrize("head", ["Ø 3 Monate", "Gleitender Durchschnitt (3 Mon.)", "Mittel letzte 3", "Ø 3-Monate"])
def test_window_with_the_wrong_length(head):
    heads = {"B": "Umsatz", "F": head}
    wrong = mine(_hints(*book(heads, lambda r, i: {f"B{r}": SALES[i], **({f"F{r}": f"=AVERAGE(B{r - 3}:B{r})"}
                                                                          if r >= 8 else {})})))
    assert wrong["span:F8"]["fix"]["cells"]["F8"] == "=AVERAGE(B6:B8)", wrong
    clean = book(heads, lambda r, i: {f"B{r}": SALES[i], **({f"F{r}": f"=AVERAGE(B{r - 2}:B{r})"} if r >= 7 else {})})
    assert "span" not in kinds(_hints(*clean))


# --- Marge durch die Kosten ---


def test_margin_divided_by_costs_is_a_markup():
    heads = {"B": "Umsatz", "C": "Kosten", "D": "Gewinn", "E": "Marge %"}
    row = lambda r, i, f: {**sales_costs(r, i), f"D{r}": f"=B{r}-C{r}", f"E{r}": (f, PERCENT)}  # noqa: E731
    wrong = mine(_hints(*book(heads, lambda r, i: row(r, i, f"=D{r}/C{r}"))))
    assert wrong["markup:E5"]["level"] == "pruefen" and wrong["markup:E5"]["fix"]["cells"]["E5"] == "=D5/B5", wrong
    assert "markup" not in kinds(_hints(*book(heads, lambda r, i: row(r, i, f"=D{r}/B{r}"))))
    markup = {**heads, "E": "Aufschlag %"}
    assert "markup" not in kinds(_hints(*book(markup, lambda r, i: row(r, i, f"=D{r}/C{r}"))))


# --- Kennzahlen neben der Tabelle ---


def figures(extra):
    heads = {"B": "Umsatz", "C": "Kosten"}
    return book(heads, sales_costs, extra)


@pytest.mark.parametrize("label", ["Bester Monat", "Höchster Umsatz-Monat", "TOP-Monat", "Umsatzstärkster Monat"])
def test_best_with_min(label):
    wrong = mine(_hints(*figures({"E5": label, "F5": "=INDEX(A5:A16,MATCH(MIN(B5:B16),B5:B16,0))"})))
    assert wrong["extreme:F5"]["fix"]["cells"]["F5"] == "=INDEX(A5:A16,MATCH(MAX(B5:B16),B5:B16,0))", wrong
    assert "extreme" not in kinds(_hints(*figures({"E5": label, "F5": "=INDEX(A5:A16,MATCH(MAX(B5:B16),B5:B16,0))"})))
    # Bei Kosten ist der tiefste Wert der beste.
    assert "extreme" not in kinds(_hints(*figures({"E5": "Bester Kostenmonat",
                                                    "F5": "=INDEX(A5:A16,MATCH(MIN(C5:C16),C5:C16,0))"})))


@pytest.mark.parametrize("label", ["Std.-Abw. Umsatz", "Standardabweichung Umsätze", "Median UMSATZ"])
def test_figure_over_the_wrong_column(label):
    func = "MEDIAN" if "Median" in label else "STDEV"
    wrong = mine(_hints(*figures({"E5": label, "F5": f"={func}(C5:C16)"})))
    assert wrong["statcol:F5"]["fix"]["cells"]["F5"] == f"={func}(B5:B16)", wrong
    assert "statcol" not in kinds(_hints(*figures({"E5": label, "F5": f"={func}(B5:B16)"})))


# --- Verschoben, vertauscht ---


def test_reference_from_the_row_above():
    heads = {"B": "Umsatz", "C": "Kosten", "D": "Marge"}
    wrong = mine(_hints(*book(heads, lambda r, i: {**sales_costs(r, i), f"D{r}": f"=B{r - 1}-C{r}"})))
    assert wrong["shifted:D5"]["fix"]["cells"]["D5"] == "=B5-C5" and wrong["shifted:D5"]["level"] == "fehler", wrong
    # Mit „Vormonat“ im Namen ist es gewollt.
    meant = {**heads, "D": "Umsatz Vormonat minus Kosten"}
    assert "shifted" not in kinds(_hints(*book(meant, lambda r, i: {**sales_costs(r, i), f"D{r}": f"=B{r - 1}-C{r}"})))


def test_revenue_and_costs_swapped():
    heads = {"B": "Umsatz", "C": "Kosten", "D": "Marge"}
    found = mine(_hints(*book(heads, lambda r, i: {f"B{r}": COSTS[i], f"C{r}": SALES[i], f"D{r}": f"=B{r}-C{r}"})))
    assert found["swapped:B4"]["fix"]["cells"] == {"B4": "Kosten", "C4": "Umsatz"}, found
    assert "swapped" not in kinds(_hints(*book(heads, lambda r, i: {**sales_costs(r, i), f"D{r}": f"=B{r}-C{r}"})))


def test_until_today_with_a_fixed_date():
    heads = {"B": "Umsatz"}
    row = lambda r, i: {f"A{r}": date(2026, i + 1, 1), f"B{r}": SALES[i]}  # noqa: E731
    found = mine(_hints(*book(heads, row, {"E5": "Umsatz bis heute",
                                           "F5": '=SUMIFS(B5:B16,A5:A16,"<="&DATE(2026,6,30))'})))
    assert found["ytd:F5"]["level"] == "pruefen", found
    quiet = book(heads, row, {"E5": "Umsatz Januar bis Juni", "F5": '=SUMIFS(B5:B16,A5:A16,"<="&DATE(2026,6,30))'})
    assert "ytd" not in kinds(_hints(*quiet))


# --- Währungen, Titel ---


def test_two_currencies_without_a_rate():
    heads = {"B": "Betrag (CHF)", "C": "Betrag (EUR)", "D": "Total (CHF)"}
    extra = {"F2": "Kurs EUR/CHF", "G2": 0.94}
    row = lambda r, i: {f"B{r}": SALES[i] / 100, f"C{r}": COSTS[i] / 100}  # noqa: E731
    wrong = mine(_hints(*book(heads, lambda r, i: {**row(r, i), f"D{r}": f"=B{r}+C{r}"}, extra)))
    assert wrong["currency:D5"]["fix"]["cells"]["D5"] == "=B5+C5*$G$2", wrong
    right = book(heads, lambda r, i: {**row(r, i), f"D{r}": f"=B{r}+C{r}*$G$2"}, extra)
    assert "currency" not in kinds(_hints(*right))


def test_title_and_header_disagree():
    heads = {"B": "Umsatz (€)", "C": "Kosten (€)", "D": "Marge (€)"}
    row = lambda r, i: {**sales_costs(r, i), f"D{r}": f"=B{r}-C{r}"}  # noqa: E731
    found = mine(_hints(*book(heads, row, {"A2": "Alle Beträge in Tausend Euro"})))
    assert any(k.startswith("title:") for k in found), found
    assert not any(k.startswith("title:") for k in mine(_hints(*book(heads, row, {"A2": "Alle Beträge in Euro"}))))


def test_real_thousands_get_converted_by_einsetzen():
    heads = {"B": "Umsatz (T€)", "C": "Kosten (€)", "D": "Marge (€)", "G": "Kumuliert (€)"}
    row = lambda r, i: {f"B{r}": SALES[i] / 1000, f"C{r}": COSTS[i], f"D{r}": f"=B{r}-C{r}",  # noqa: E731
                        f"G{r}": f"=B{r}" if r == 5 else f"=G{r - 1}+B{r}"}
    found = mine(_hints(*book(heads, row)))
    item = next(h for k, h in found.items() if k.startswith("unit:"))
    assert item["level"] == "fehler" and item["fix"]["cells"]["D5"] == "=B5*1000-C5", item
    assert item["fix"]["cells"]["G5"] == "=(B5)*1000" and item["fix"]["cells"]["G6"] == "=G5+B6*1000"
