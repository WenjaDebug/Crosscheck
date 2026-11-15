#!/bin/bash
# Installiert Gegenprobe auf einem Mac: eine eigene Python-Umgebung im Ordner .venv mit den drei Paketen.
# Aufruf: Doppelklick auf VoiceAgent.command oder ExcelAddin.command (oder im Terminal: ./scripts/install-mac.sh).
#
# Python 3.10 oder neuer nimmt es, wo es schon eins findet: von python.org, von Homebrew oder sonst auf dem PATH.
# Fehlt eins, installiert es Python 3.13 mit Homebrew, falls Homebrew da ist, sonst öffnet es python.org und sagt,
# was zu tun ist. Das Python der Befehlszeilentools (/usr/bin/python3) ist meist 3.9 und damit zu alt; ohne die
# Befehlszeilentools ist es nur ein Platzhalter, der beim Aufruf fragt, ob macOS sie installieren soll. Darum ruft
# dieses Skript es nur auf, wenn `xcode-select -p` sagt, dass sie da sind.
# Ein anderes Python lässt sich mit VOICEAGENT_PYTHON vorgeben.

PYTHON_DOWNLOAD="https://www.python.org/downloads/macos/"

# Alle Stellen, an denen ein passendes Python liegen kann, das empfohlene zuerst.
python_candidates() {
  [ -n "${VOICEAGENT_PYTHON:-}" ] && echo "$VOICEAGENT_PYTHON"
  local version prefix
  for version in 3.13 3.12 3.11 3.10 3.14; do
    echo "/Library/Frameworks/Python.framework/Versions/$version/bin/python3"
    for prefix in /opt/homebrew /usr/local; do
      echo "$prefix/opt/python@$version/bin/python$version"
      echo "$prefix/bin/python$version"
    done
  done
  command -v python3 2>/dev/null
}

# Das Platzhalter-Python von macOS nur fragen, wenn die Befehlszeilentools wirklich da sind (oder wenn du es selbst
# mit VOICEAGENT_PYTHON vorgibst).
usable_path() {
  if [ "$1" = "${VOICEAGENT_PYTHON:-}" ]; then
    [ -x "$1" ]
    return
  fi
  case "$1" in
    /usr/bin/python3) xcode-select -p >/dev/null 2>&1 ;;
    *) [ -x "$1" ] ;;
  esac
}

# Python 3.10 oder neuer?
new_enough() {
  "$1" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)' >/dev/null 2>&1
}

find_python() {
  local candidate
  while IFS= read -r candidate; do
    [ -n "$candidate" ] || continue
    if usable_path "$candidate" && new_enough "$candidate"; then
      echo "$candidate"
      return 0
    fi
  done < <(python_candidates)
  return 1
}

find_brew() {
  if command -v brew >/dev/null 2>&1; then
    command -v brew
    return 0
  fi
  # Homebrew liegt auf Apple-Silicon-Macs unter /opt/homebrew, auf älteren unter /usr/local.
  local candidate
  for candidate in /opt/homebrew/bin/brew /usr/local/bin/brew; do
    if [ -x "$candidate" ]; then
      echo "$candidate"
      return 0
    fi
  done
  return 1
}

explain_python() {
  echo
  echo "Gegenprobe braucht Python (ein kostenloses Programm), und das fehlt auf diesem Mac noch."
  echo "Ich öffne dir jetzt die Seite python.org. So geht es weiter:"
  echo "  1. Dort Python 3.13 herunterladen (bei „Python 3.13…“ auf „macOS 64-bit universal2 installer“ klicken)."
  echo "  2. Die geladene Datei doppelklicken und mit „Fortfahren“ und „Installieren“ durchklicken."
  echo "     Der Mac fragt dabei nach deinem Passwort."
  echo "  3. Danach diese Datei (VoiceAgent.command oder ExcelAddin.command) noch einmal doppelklicken."
  open "$PYTHON_DOWNLOAD" >/dev/null 2>&1 || echo "Die Seite: $PYTHON_DOWNLOAD"
}

# Kommt der Mac ins Internet? Ohne TLS und ohne etwas zu schicken: nur die Verbindung zu pypi.org.
online() {
  "$1" -c 'import socket; socket.create_connection(("pypi.org", 443), 5).close()' >/dev/null 2>&1
}

explain_offline() {
  echo
  echo "Ich komme nicht ins Internet. Beim ersten Start lädt Gegenprobe einmal ihre drei Pakete"
  echo "(FastAPI, Uvicorn, pypdf). Danach geht alles ohne Internet."
  echo "Verbinde den Mac mit dem Internet (WLAN oder Kabel) und doppelklick die Datei noch einmal."
}

# In Downloads, auf dem Schreibtisch, in Dokumente oder in iCloud fragt macOS bei jedem Programm nach, ob es dort
# lesen darf, auch beim Prüfer im Hintergrund. Später verschieben geht nicht mehr (die .venv hängt am Ort).
check_place() {
  case "$PWD/" in
    "$HOME/Downloads/"* | "$HOME/Desktop/"* | "$HOME/Documents/"* | "$HOME/Library/Mobile Documents/"*)
      echo "Tipp: Dieser Ordner liegt in Downloads, auf dem Schreibtisch, in Dokumente oder in iCloud."
      echo "Dort fragt der Mac immer wieder, ob Gegenprobe ihn lesen darf. Besser ist dein Benutzerordner"
      echo "(im Finder: Gehe zu → Benutzerordner). Später verschieben geht nicht mehr so einfach."
      echo "So geht es: dieses Fenster schließen, den Ordner in den Benutzerordner ziehen, dort noch einmal doppelklicken."
      read -r -p "Oder Enter drücken, dann bleibt er hier. " || true
      ;;
  esac
}

main() {
  set -uo pipefail
  cd "$(dirname "$0")/.." || exit 1
  check_place

  echo "1/3 Ich suche Python 3.10 oder neuer …"
  local python
  if ! python="$(find_python)"; then
    local brew
    if brew="$(find_brew)"; then
      echo "    Keins gefunden. Ich installiere Python 3.13 mit Homebrew …"
      eval "$("$brew" shellenv)"
      if ! brew install python@3.13; then
        echo "Homebrew konnte Python nicht installieren."
        explain_python
        exit 1
      fi
      python="$(brew --prefix python@3.13)/bin/python3.13"
    else
      explain_python
      exit 1
    fi
  fi
  echo "    $python"

  echo "2/3 Ich lege eine eigene Python-Umgebung im Ordner .venv an …"
  if ! "$python" -m venv .venv; then
    echo "Das ging nicht. Hat dieser Ordner einen Schreibschutz, oder ist die Festplatte voll?"
    exit 1
  fi

  echo "3/3 Ich lade und installiere Gegenprobe und ihre drei Pakete, das dauert eine Minute …"
  .venv/bin/python -m pip install --upgrade pip --quiet --disable-pip-version-check >/dev/null 2>&1 || true
  if ! .venv/bin/python -m pip install -e . --quiet --disable-pip-version-check; then
    if ! online .venv/bin/python; then
      explain_offline
    else
      echo
      echo "Die Pakete ließen sich nicht installieren. Was oben steht, hilft beim Fragen (Bildschirmfoto)."
      echo "Hilfe: docs/HILFE.md im Ordner von Gegenprobe."
    fi
    exit 1
  fi

  echo
  echo "Installiert."
}

# Nur ausführen, wenn es gestartet und nicht von einem Test eingelesen wird.
if [ "${BASH_SOURCE[0]}" = "$0" ]; then
  main "$@"
fi
