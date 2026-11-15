"""Herunterladen als ZIP, entpacken, doppelklicken: Die Startdateien müssen das auch ohne git und ohne Terminal-
Kenntnisse schaffen, am Mac und unter Windows, in Ordnern mit Leerzeichen und Umlauten."""

import os
import re
import shutil
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
MAC = ("VoiceAgent.command", "ExcelAddin.command", "scripts/install-mac.sh")
bash = pytest.mark.skipif(shutil.which("bash") is None or os.name == "nt", reason="bash fehlt")


@bash
@pytest.mark.parametrize("name", MAC)
def test_mac_scripts_parse(name):
    assert subprocess.run(["bash", "-n", str(ROOT / name)], capture_output=True).returncode == 0


def test_mac_scripts_keep_their_execute_bit_in_git():
    """GitHub packt die ZIP mit git archive, das das Ausführungsrecht mitnimmt (Unix-Rechte in der ZIP)."""
    if shutil.which("git") is None or not (ROOT / ".git").exists():
        pytest.skip("kein Git-Klon")
    listed = subprocess.run(["git", "ls-files", "-s", *MAC], cwd=ROOT, capture_output=True, text=True).stdout
    assert listed.count("100755") == len(MAC), listed


def test_mac_launchers_install_with_bash_and_say_what_comes_next():
    for name in ("VoiceAgent.command", "ExcelAddin.command"):
        text = (ROOT / name).read_text(encoding="utf-8")
        # Mit bash aufgerufen geht es auch, wenn das Entpacken die Ausführungsrechte verloren hat.
        assert "bash ./scripts/install-mac.sh" in text and 'cd "$(dirname "$0")"' in text
        assert "Noch nicht fertig, oben steht, was zu tun ist." in text


FAKE = """#!/bin/bash
echo "$0 $*" >> "{log}"
exec python3 -c 'import sys; sys.exit(0 if {ok} else 1)'
"""


def fake_python(folder: Path, name: str, ok: bool, log: Path) -> Path:
    path = folder / name
    path.write_text(FAKE.format(log=log, ok=ok), encoding="utf-8")
    path.chmod(0o755)
    return path


def source(script: str, tmp_path: Path, path_dirs=()) -> subprocess.CompletedProcess:
    env = {**os.environ, "PATH": os.pathsep.join([*map(str, path_dirs), os.environ["PATH"]]), "HOME": str(tmp_path)}
    env.pop("VOICEAGENT_PYTHON", None)
    return subprocess.run(["bash", "-c", f'source "{ROOT / "scripts" / "install-mac.sh"}"\n{script}'],
                          capture_output=True, text=True, env=env, cwd=tmp_path)


@bash
def test_mac_takes_the_first_python_that_is_new_enough(tmp_path):
    log = tmp_path / "gefragt.txt"
    old = fake_python(tmp_path, "python3.9", False, log)
    new = fake_python(tmp_path, "python3.13", True, log)
    done = source(f'python_candidates() {{ echo "{tmp_path}/fehlt"; echo "{old}"; echo "{new}"; }}\nfind_python', tmp_path)
    assert done.returncode == 0 and done.stdout.strip() == str(new)
    done = source(f'python_candidates() {{ echo "{old}"; }}\nfind_python', tmp_path)
    assert done.returncode == 1 and done.stdout == ""


@bash
def test_mac_never_wakes_the_command_line_tools_placeholder(tmp_path):
    """/usr/bin/python3 ohne Befehlszeilentools fragt, ob macOS sie installieren soll. Nur fragen, wenn xcode-select
    sagt, dass sie da sind."""
    tools = tmp_path / "bin"
    tools.mkdir()
    (tools / "xcode-select").write_text("#!/bin/bash\nexit 2\n", encoding="utf-8")
    (tools / "xcode-select").chmod(0o755)
    script = 'python_candidates() { echo /usr/bin/python3; }\nnew_enough() { echo "gefragt: $1" >&2; return 0; }\nfind_python'
    done = source(script, tmp_path, [tools])
    assert done.returncode == 1 and "gefragt" not in done.stderr
    (tools / "xcode-select").write_text("#!/bin/bash\necho /Library/Developer/CommandLineTools\n", encoding="utf-8")
    if Path("/usr/bin/python3").exists():
        done = source(script, tmp_path, [tools])
        assert done.returncode == 0 and "gefragt: /usr/bin/python3" in done.stderr


@bash
def test_mac_looks_in_the_usual_places_first(tmp_path):
    done = source("VOICEAGENT_PYTHON=/eigenes/python3\npython_candidates", tmp_path)
    found = done.stdout.split("\n")
    assert found[0] == "/eigenes/python3"
    assert found[1] == "/Library/Frameworks/Python.framework/Versions/3.13/bin/python3"
    assert "/opt/homebrew/opt/python@3.13/bin/python3.13" in found and "/usr/local/bin/python3.12" in found


@bash
def test_mac_without_python_opens_python_org_and_says_what_to_do(tmp_path):
    tools = tmp_path / "bin"
    tools.mkdir()
    (tools / "open").write_text(f'#!/bin/bash\necho "$@" > "{tmp_path}/geoeffnet.txt"\n', encoding="utf-8")
    (tools / "open").chmod(0o755)
    done = source("explain_python", tmp_path, [tools])
    assert (tmp_path / "geoeffnet.txt").read_text().strip() == "https://www.python.org/downloads/macos/"
    assert "Python 3.13" in done.stdout and "noch einmal doppelklicken" in done.stdout


@bash
def test_mac_tip_for_downloads_only_there(tmp_path):
    (tmp_path / "Downloads" / "VoiceAgent-main").mkdir(parents=True)
    done = subprocess.run(["bash", "-c", f'source "{ROOT / "scripts" / "install-mac.sh"}"\ncheck_place'], input="\n",
                          capture_output=True, text=True, env={**os.environ, "HOME": str(tmp_path)},
                          cwd=tmp_path / "Downloads" / "VoiceAgent-main")
    assert "Benutzerordner" in done.stdout
    done = subprocess.run(["bash", "-c", f'source "{ROOT / "scripts" / "install-mac.sh"}"\ncheck_place'], input="",
                          capture_output=True, text=True, env={**os.environ, "HOME": str(tmp_path)}, cwd=tmp_path)
    assert done.stdout == ""


@bash
def test_mac_offline_check_needs_no_data(tmp_path):
    text = (ROOT / "scripts" / "install-mac.sh").read_text(encoding="utf-8")
    assert 'socket.create_connection(("pypi.org", 443), 5)' in text and "explain_offline" in text


# --- Windows ---


def bat(name: str) -> str:
    data = (ROOT / name).read_bytes()
    # cmd liest .bat in der Codepage der Konsole: nur ASCII, mit Windows-Zeilenenden.
    assert data.isascii() and b"\n" not in data.replace(b"\r\n", b""), name
    return data.decode("ascii")


@pytest.mark.parametrize("name", ["VoiceAgent.bat", "ExcelAddin.bat"])
def test_bat_quotes_every_path_and_notices_an_unpacked_zip(name):
    text = bat(name)
    assert 'pushd "%~dp0"' in text
    # Jede Verwendung von %~dp0 steht in Anführungszeichen (Leerzeichen, Umlaute, Klammern im Ordnernamen).
    for line in text.splitlines():
        for match in re.finditer(r"%~dp0", line):
            assert line[: match.start()].count('"') % 2 == 1, line
    assert "Alle extrahieren" in text and "goto nicht_entpackt" in text


def ps1(name: str) -> str:
    data = (ROOT / "scripts" / name).read_bytes()
    # Windows PowerShell liest Skripte ohne BOM nicht als UTF-8, dann wären die Umlaute kaputt.
    assert data.startswith(b"\xef\xbb\xbf") and b"\n" not in data.replace(b"\r\n", b""), name
    return data[3:].decode("utf-8")


def test_windows_install_finds_python_without_the_store_stub():
    text = ps1("install-windows.ps1")
    assert "Set-Location -LiteralPath" in text
    assert '-like "*\\WindowsApps\\*"' in text and "sys.version_info >= (3, 10)" in text
    for place in ('Programs\\Python\\Python$version\\python.exe', '"pythoncore-3.*"', "Get-Command python -All"):
        assert place in text, place
    assert 'Start-Process $PythonDownload' in text and '"https://www.python.org/downloads/windows/"' in text
    assert "Test-Online" in text and "Ich komme nicht ins Internet" in text
    start = ps1("start-windows.ps1")
    assert "Set-Location -LiteralPath" in start and "$LASTEXITCODE -ne 0" in start


@pytest.mark.skipif(shutil.which("pwsh") is None, reason="pwsh fehlt: PowerShell-Syntax hier nicht prüfbar")
@pytest.mark.parametrize("name", ["install-windows.ps1", "start-windows.ps1"])
def test_powershell_syntax(name):
    script = ("$e = $null; [void][System.Management.Automation.Language.Parser]::ParseFile("
              f"'{ROOT / 'scripts' / name}', [ref]$null, [ref]$e); if ($e) {{ $e; exit 1 }}")
    assert subprocess.run(["pwsh", "-NoProfile", "-Command", script], capture_output=True).returncode == 0


# --- Anleitung und Hilfe ---


def slugs(markdown: str) -> set[str]:
    """Die Anker, die GitHub aus den Überschriften macht."""
    found = set()
    for line in markdown.splitlines():
        if line.startswith("#"):
            title = line.lstrip("#").strip().lower()
            found.add(re.sub(r"[^\w\- ]", "", title).replace(" ", "-"))
    return found


@pytest.mark.parametrize("name", ["README.md", "README.de.md", "docs/FEATURES.md", "docs/FUNKTIONEN.md", "docs/HILFE.md",
                                  "docs/HELP.md"])
def test_links_images_and_anchors_exist(name):
    path = ROOT / name
    text = path.read_text(encoding="utf-8")
    for target in re.findall(r"\]\(([^)\s]+)\)|src=\"([^\"]+)\"", text):
        target = target[0] or target[1]
        if target.startswith(("http://", "https://", "mailto:")):
            continue
        file, _, anchor = target.partition("#")
        linked = (path.parent / file) if file else path
        assert linked.exists(), f"{name}: {target}"
        if anchor and linked.suffix == ".md":
            assert anchor in slugs(linked.read_text(encoding="utf-8")), f"{name}: {target}"


def test_readmes_start_with_five_steps_and_a_small_picture():
    for name, heading, picture in (("README.md", "## Install in 5 steps", "seitenfenster-en.png"),
                                   ("README.de.md", "## In 5 Schritten", "seitenfenster-de.png")):
        text = (ROOT / name).read_text(encoding="utf-8")
        assert text.index(picture) < text.index(heading) and len(text.split()) < 900
        assert "docs/HILFE.md" in text or "docs/HELP.md" in text
        data = (ROOT / "docs" / "bilder" / picture).read_bytes()
        assert data.startswith(b"\x89PNG") and len(data) < 200_000


def test_help_covers_every_known_problem():
    hilfe = (ROOT / "docs" / "HILFE.md").read_text(encoding="utf-8")
    for topic in ("Python fehlt", "Kein Internet", "Dennoch öffnen", "Trotzdem ausführen", "Passwort", "Ja** klicken",
                  "Knopf Prüfer fehlt", "Clear Web Cache", "Wef", "Excel 2016 oder 2019", "Port belegt", "Virenscanner",
                  "--entfernen", "Alles deinstallieren"):
        assert topic in hilfe, topic
    help_en = (ROOT / "docs" / "HELP.md").read_text(encoding="utf-8")
    assert slugs(help_en) and len(slugs(help_en)) == len(slugs(hilfe))
