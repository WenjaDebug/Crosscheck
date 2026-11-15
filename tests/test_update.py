"""Selbst aktualisieren: Vor dem Start die neueste Version holen, aber nie etwas kaputt machen."""

import io
import shutil
import subprocess
import tarfile
from pathlib import Path

import pytest

from voiceagent.update import MARKER, drop_retired, pull, stand, sync_packages, version

pytestmark = pytest.mark.skipif(shutil.which("git") is None, reason="git fehlt")


def git(cwd, *args):
    return subprocess.run(["git", *args], cwd=cwd, check=True, capture_output=True, text=True, encoding="utf-8").stdout.strip()


def commit(repo: Path, name: str, text: str) -> None:
    (repo / name).write_text(text, encoding="utf-8")
    git(repo, "add", name)
    git(repo, "-c", "user.name=Test", "-c", "user.email=test@example.com", "commit", "-q", "-m", f"{name} geändert")


@pytest.fixture
def clones(tmp_path, monkeypatch):
    """Ein GitHub zum Mitnehmen: ein leeres Repository, ein eigener Klon und einer, der neue Versionen hochlädt."""
    monkeypatch.delenv("VOICEAGENT_NO_UPDATE", raising=False)
    remote = tmp_path / "github.git"
    git(tmp_path, "init", "-q", "--bare", "-b", "main", str(remote))
    upstream = tmp_path / "entwicklung"
    git(tmp_path, "clone", "-q", str(remote), str(upstream))
    git(upstream, "symbolic-ref", "HEAD", "refs/heads/main")
    commit(upstream, "README.md", "Version 1")
    git(upstream, "push", "-q", "origin", "main")
    mine = tmp_path / "VoiceAgent"
    git(tmp_path, "clone", "-q", str(remote), str(mine))
    return upstream, mine


def test_new_version_is_pulled(clones):
    upstream, mine = clones
    assert pull(mine) == "Gegenprobe ist aktuell."
    commit(upstream, "README.md", "Version 2")
    commit(upstream, "neu.txt", "dazu")
    git(upstream, "push", "-q", "origin", "main")
    assert pull(mine) == "Neue Version geladen (2 Änderungen)."
    assert (mine / "README.md").read_text(encoding="utf-8") == "Version 2"


def test_own_changes_and_other_branches_are_left_alone(clones):
    upstream, mine = clones
    commit(upstream, "README.md", "Version 2")
    git(upstream, "push", "-q", "origin", "main")
    (mine / "README.md").write_text("Meine Notiz", encoding="utf-8")
    assert "eigene Änderungen" in pull(mine)
    assert (mine / "README.md").read_text(encoding="utf-8") == "Meine Notiz"
    git(mine, "checkout", "-q", "README.md")
    git(mine, "checkout", "-q", "-b", "feature/test")
    assert "Zweig „feature/test“" in pull(mine)
    assert (mine / "README.md").read_text(encoding="utf-8") == "Version 1"


def test_offline_just_starts_the_old_version(clones, tmp_path):
    _, mine = clones
    # GitHub nicht erreichbar. (Löschen geht unter Windows nicht, Git-Objekte sind schreibgeschützt.)
    git(mine, "remote", "set-url", "origin", str(tmp_path / "gibt-es-nicht.git"))
    assert "bisherigen Version" in pull(mine)


def test_switched_off_and_without_git(clones, tmp_path, monkeypatch):
    _, mine = clones
    assert pull(tmp_path / "kein-klon") == ""
    monkeypatch.setenv("VOICEAGENT_NO_UPDATE", "1")
    assert "ausgeschaltet" in pull(mine)


def test_version_shows_the_commit(clones):
    _, mine = clones
    info = version(mine)
    assert info["branch"] == "main" and len(info["commit"]) >= 7 and info["date"].count("-") == 2


class Pip:
    def __init__(self, code=0):
        self.code, self.calls = code, []

    def __call__(self, args):
        self.calls.append(args)
        return subprocess.CompletedProcess(args, self.code)


def test_packages_are_installed_only_when_they_changed(tmp_path):
    (tmp_path / "pyproject.toml").write_text('dependencies = ["fastapi"]', encoding="utf-8")
    venv = tmp_path / "venv"
    venv.mkdir()
    pip = Pip()
    assert sync_packages(tmp_path, venv, run=pip) is True
    assert pip.calls[0][1:5] == ["-m", "pip", "install", "-e"]
    assert sync_packages(tmp_path, venv, run=pip) is False and len(pip.calls) == 1
    (tmp_path / "pyproject.toml").write_text('dependencies = ["fastapi", "pypdf"]', encoding="utf-8")
    assert sync_packages(tmp_path, venv, run=pip) is True and len(pip.calls) == 2


def test_packages_of_the_old_speech_parts_are_removed():
    """Whisper, Tesseract und ihre Begleiter braucht Gegenprobe nicht mehr: Jeder Start nimmt sie aus .venv, außer ein
    anderes Paket braucht sie noch."""
    pip = Pip()
    env = lambda: ({"fastapi", "faster-whisper", "numpy", "pillow", "pandas"}, {"numpy", "starlette"})  # noqa: E731
    assert drop_retired(run=pip, environment=env) == ["faster-whisper", "pillow"]
    assert pip.calls[0][1:5] == ["-m", "pip", "uninstall", "-y"] and pip.calls[0][-2:] == ["faster-whisper", "pillow"]
    assert drop_retired(run=Pip(), environment=lambda: ({"fastapi"}, set())) == []
    assert drop_retired(run=Pip(code=1), environment=env) == []


def test_the_real_environment_can_be_read():
    from voiceagent.update import _environment

    installed, needed = _environment()
    assert "fastapi" in installed and "starlette" in needed


def test_failed_install_is_tried_again(tmp_path):
    (tmp_path / "pyproject.toml").write_text("x", encoding="utf-8")
    assert sync_packages(tmp_path, tmp_path, run=Pip(code=1)) is False
    assert not (tmp_path / MARKER).exists()


def test_launchers_update_before_starting():
    root = Path(__file__).resolve().parents[1]
    mac = (root / "VoiceAgent.command").read_text(encoding="utf-8")
    assert mac.index("-m voiceagent.update") < mac.index("exec .venv/bin/voiceagent")
    assert mac.rstrip().endswith('main "$@"\nexit')
    windows = (root / "scripts" / "start-windows.ps1").read_text(encoding="utf-8-sig")
    assert windows.index("-m voiceagent.update") < windows.index("voiceagent.exe @args")


# --- Aus einer ZIP von GitHub (ohne .git) ---


def zip_folder(tmp_path, text="9f062b75817394dd66163baf11c299896136a3c0 1794334526"):
    """Ein entpackter Ordner wie aus „Code → Download ZIP“: kein .git, aber stand.txt von GitHub ausgefüllt."""
    root = tmp_path / "VoiceAgent-main"
    (root / "src" / "voiceagent").mkdir(parents=True)
    (root / "src" / "voiceagent" / "stand.txt").write_text(text + "\n", encoding="utf-8")
    venv = tmp_path / "venv"
    venv.mkdir()
    return root, venv


def test_zip_install_pulls_nothing_and_says_so_only_when_newer(tmp_path, monkeypatch):
    from voiceagent.update import hint

    monkeypatch.delenv("VOICEAGENT_NO_UPDATE", raising=False)
    root, venv = zip_folder(tmp_path)
    assert pull(root) == ""
    asked = []
    newer = lambda: asked.append(1) or 1794334526 + 86400 * 3  # noqa: E731
    said = hint(root, venv, fetch=newer, now=lambda: 1_800_000_000)
    assert said.startswith("Es gibt eine neuere Version von Gegenprobe (vom ") and "github.com/WenjaDebug/Crosscheck" in said
    # Höchstens einmal am Tag gefragt: Der nächste Start sagt dasselbe, ohne GitHub zu fragen.
    assert hint(root, venv, fetch=newer, now=lambda: 1_800_000_000 + 3600) == said and len(asked) == 1
    assert hint(root, venv, fetch=newer, now=lambda: 1_800_000_000 + 86400) == said and len(asked) == 2


def test_zip_install_up_to_date_offline_or_unknown_stays_quiet(tmp_path):
    from voiceagent.update import CHECKED, hint

    root, venv = zip_folder(tmp_path)
    assert hint(root, venv, fetch=lambda: 1794334526, now=lambda: 1_800_000_000) == ""
    # Offline: nichts sagen, und erst morgen wieder fragen (sonst wartet jeder Start auf das Netz).
    (venv / CHECKED).unlink()
    asked = []
    offline = lambda: asked.append(1) or None  # noqa: E731
    assert hint(root, venv, fetch=offline, now=lambda: 1_800_000_000) == ""
    assert hint(root, venv, fetch=offline, now=lambda: 1_800_000_000 + 60) == "" and len(asked) == 1
    # Ein kaputter Zwischenstand stört nicht.
    (venv / CHECKED).write_text("{kaputt", encoding="utf-8")
    assert hint(root, venv, fetch=lambda: 1794334526 + 10, now=lambda: 1_800_000_000).startswith("Es gibt")
    # Im Git-Klon steht der Platzhalter: dann nie fragen.
    plain, venv2 = zip_folder(tmp_path / "klon", "$Format:%H %ct$")
    assert hint(plain, venv2, fetch=lambda: pytest.fail("darf nicht fragen")) == ""
    (plain / ".git").mkdir()
    (plain / "src" / "voiceagent" / "stand.txt").write_text("9f062b7 1794334526", encoding="utf-8")
    assert hint(plain, venv2, fetch=lambda: pytest.fail("darf nicht fragen")) == ""


def test_the_repository_file_is_a_placeholder_filled_by_github():
    root = Path(__file__).resolve().parents[1]
    assert (root / "src" / "voiceagent" / "stand.txt").read_text(encoding="utf-8").strip() == "$Format:%H %ct$"
    assert "src/voiceagent/stand.txt export-subst" in (root / ".gitattributes").read_text(encoding="utf-8")
    assert stand(root / "src" / "voiceagent" / "stand.txt") is None


def test_git_archive_fills_in_the_stand(tmp_path):
    """So packt GitHub „Download ZIP“: git archive mit export-subst."""
    repo = tmp_path / "repo"
    repo.mkdir()
    git(repo, "init", "-q", "-b", "main")
    commit(repo, ".gitattributes", "stand.txt export-subst\n")
    commit(repo, "stand.txt", "$Format:%H %ct$\n")
    out = subprocess.run(["git", "archive", "HEAD", "stand.txt"], cwd=repo, check=True, capture_output=True).stdout
    text = tarfile.open(fileobj=io.BytesIO(out)).extractfile("stand.txt").read().decode("utf-8")
    (tmp_path / "stand.txt").write_text(text, encoding="utf-8")
    found = stand(tmp_path / "stand.txt")
    assert found and found[0] == git(repo, "rev-parse", "HEAD")


def test_version_of_a_zip_install(tmp_path):
    root, _ = zip_folder(tmp_path)
    info = version(root)
    assert info["commit"] == "9f062b7" and info["date"] == "2026-11-10" and info["branch"] is None


def test_latest_reads_the_date_from_github_without_account(monkeypatch):
    import urllib.error
    import urllib.request

    from voiceagent import update

    seen = []

    def answer(request, timeout, context):
        seen.append(request)
        body = b'[{"sha": "abc", "commit": {"committer": {"date": "2026-11-12T08:00:00Z"}}}]'
        return io.BytesIO(body)

    monkeypatch.setattr(urllib.request, "urlopen", answer)
    assert update.latest() == 1794470400
    assert seen[0].full_url == "https://api.github.com/repos/WenjaDebug/Crosscheck/commits?sha=main&per_page=1"
    # Kein Konto, kein Schlüssel, nichts aus dem Ordner.
    assert "Authorization" not in seen[0].headers and seen[0].data is None

    def offline(request, timeout, context):
        raise urllib.error.URLError("keine Verbindung")

    monkeypatch.setattr(urllib.request, "urlopen", offline)
    assert update.latest() is None
