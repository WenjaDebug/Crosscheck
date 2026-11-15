#!/bin/bash
# Startet Gegenprobe per Doppelklick im Finder. Beim ersten Mal wird alles installiert (scripts/install-mac.sh),
# danach schaut es vor jedem Start nach einer neuen Version (ausschalten: VOICEAGENT_NO_UPDATE=1): Ein Git-Klon holt
# sie, eine ZIP-Installation sagt nur, dass es eine gibt.
# Alles steht in main(): Bash liest das Skript so ganz, bevor ein Update es ändern kann.
main() {
  cd "$(dirname "$0")" || exit 1

  for candidate in /opt/homebrew/bin/brew /usr/local/bin/brew; do
    if [ -x "$candidate" ]; then
      eval "$("$candidate" shellenv)"
      break
    fi
  done

  if [ ! -x .venv/bin/voiceagent ]; then
    echo "Gegenprobe wird zum ersten Mal eingerichtet."
    # Mit bash: So geht es auch, wenn das Entpacken die Ausführungsrechte des Skripts verloren hat.
    bash ./scripts/install-mac.sh || { echo; read -r -p "Noch nicht fertig, oben steht, was zu tun ist. Enter schließt das Fenster."; exit 1; }
    echo
    echo "Gleich öffnet sich Gegenprobe im Browser. Für das Seitenfenster in Excel doppelklick danach ExcelAddin.command."
  fi

  .venv/bin/python -m voiceagent.update

  echo "Gegenprobe startet. Zum Beenden dieses Fenster schließen oder Ctrl + C drücken."
  exec .venv/bin/voiceagent "$@"
}
main "$@"
exit
