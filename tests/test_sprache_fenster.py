"""Das Seitenfenster auf Deutsch oder Englisch: Jeder Satz, den es selbst sagt, hat eine englische Fassung mit
denselben Platzhaltern, jeder feste Text im HTML auch, und Korrekturen gibt es so, wie englisches Excel sie zeigt."""

import json
import re
from html.parser import HTMLParser
from pathlib import Path

from test_addin import payload
from test_excel_live import FakeExcel, client_with

ADDIN = Path(__file__).resolve().parents[1] / "src" / "voiceagent" / "web" / "addin"


def _english() -> dict:
    script = (ADDIN / "sprache.js").read_text(encoding="utf-8")
    body = re.search(r"const EN = \{\n(.*?)\n\};", script, re.S).group(1)
    pairs = re.findall(r'^\s*("(?:[^"\\]|\\.)*"):\s*\n?\s*("(?:[^"\\]|\\.)*"),?\s*$', body, re.M)
    return {json.loads(k): json.loads(v) for k, v in pairs}


def _said() -> set[str]:
    script = (ADDIN / "pruefer.js").read_text(encoding="utf-8")
    return {json.loads(m) for m in re.findall(r'tr\(("(?:[^"\\]|\\.)*")', script)}


class _Texts(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.texts, self.skip, self.own = [], 0, []

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        self.own.append("data-en" in attrs)
        if tag in ("script", "style"):
            self.skip += 1
        for name in ("title", "aria-label", "placeholder"):
            if attrs.get(name):
                self.texts.append(" ".join(attrs[name].split()))

    def handle_endtag(self, tag):
        if self.own:
            self.own.pop()
        if tag in ("script", "style"):
            self.skip -= 1

    def handle_data(self, data):
        text = " ".join(data.split())
        if text and not self.skip and not any(self.own) and re.search(r"[A-Za-zÄÖÜäöü]{2}", text):
            self.texts.append(text)


def test_every_sentence_of_the_pane_has_an_english_version():
    english = _english()
    assert len(english) > 200
    missing = sorted(_said() - set(english))
    assert not missing, missing
    for german, said in english.items():
        assert re.findall(r"\{\d+\}", german.split("|")[0]) == re.findall(r"\{\d+\}", said) or \
            sorted(re.findall(r"\{\d+\}", german)) == sorted(re.findall(r"\{\d+\}", said)), german
        assert not re.search(r"[äöüÄÖÜß„]", said), said


def test_every_fixed_text_of_the_pane_has_an_english_version():
    parser = _Texts()
    parser.feed((ADDIN / "pruefer.html").read_text(encoding="utf-8"))
    english = _english()
    allowed = {"Gegenprobe", "✓", "Sprache / Language", "Gegenprobe __VERSION__"}
    missing = sorted({t for t in parser.texts if t not in english and t not in allowed})
    assert not missing, missing


def test_a_fix_comes_as_english_excel_shows_it_too(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    result = client.post("/api/addin/pruefen", json=payload(), headers={"X-VoiceAgent": "1"}).json()
    fixes = [h["fix"] for h in result["hints"] if h.get("fix", {}).get("german", "").startswith("=")]
    assert fixes and all(f["english"] == f["german"].replace("SUMME", "SUM").replace(";", ",") for f in fixes), fixes


def test_the_pane_loads_the_language_before_its_script(tmp_path):
    _, client = client_with(tmp_path, FakeExcel(None))
    page = client.get("/addin/pruefer.html").text
    assert page.index("/addin/sprache.js?v=") < page.index("/addin/pruefer.js?v=")
    assert client.get("/addin/sprache.js").status_code == 200
