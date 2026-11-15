#!/bin/bash
# Richtet den Excel-Prüfer einmal als Seitenfenster in Excel ein (am Mac; unter Windows ExcelAddin.bat). Er läuft danach still im
# Hintergrund: In Excel unter Start auf „Prüfer“ klicken genügt. Fehlt die Installation, legt es sie zuerst an.
# Wieder entfernen: ./ExcelAddin.command --entfernen
main() {
  cd "$(dirname "$0")" || exit 1
  if [ ! -x .venv/bin/python ]; then
    echo "Gegenprobe wird zum ersten Mal eingerichtet."
    # Mit bash: So geht es auch, wenn das Entpacken die Ausführungsrechte des Skripts verloren hat.
    bash ./scripts/install-mac.sh || { echo; read -r -p "Noch nicht fertig, oben steht, was zu tun ist. Enter schließt das Fenster."; exit 1; }
    echo
    echo "Jetzt richte ich das Seitenfenster in Excel ein."
  fi
  # Ohne „Enter schließt das Fenster“ am Ende: Im offenen Terminal schluckte das sonst die nächste eingefügte Zeile.
  # Per Doppelklick bleibt das Fenster trotzdem offen, bis du es schließt.
  .venv/bin/python -m voiceagent.addin "$@"
}
main "$@"
exit
