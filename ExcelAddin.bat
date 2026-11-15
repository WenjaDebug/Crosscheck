@echo off
rem Richtet den Excel-Pruefer einmal als Seitenfenster in Excel ein (Windows). Er laeuft danach still im
rem Hintergrund: In Excel unter Start auf "Pruefer" klicken genuegt. Beim ersten Mal wird zuerst alles installiert.
rem Wieder entfernen: ExcelAddin.bat --entfernen
rem Alle Pfade in Anfuehrungszeichen: Der Ordner darf Leerzeichen und Umlaute haben (C:\Users\Juerg Mueller\...).
pushd "%~dp0"
if not exist "%~dp0scripts\install-windows.ps1" goto nicht_entpackt
if not exist ".venv\Scripts\voiceagent.exe" (
  echo Gegenprobe wird zum ersten Mal eingerichtet.
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\install-windows.ps1"
  if errorlevel 1 (
    echo.
    echo Noch nicht fertig, oben steht, was zu tun ist.
    pause
    exit /b 1
  )
  echo.
  echo Jetzt richte ich das Seitenfenster in Excel ein.
)
".venv\Scripts\python.exe" -m voiceagent.addin %*
set result=%errorlevel%
rem Das Fenster bleibt offen, damit man lesen kann, was es gesagt hat.
pause
exit /b %result%

:nicht_entpackt
echo Gegenprobe steckt noch in der ZIP-Datei. Bitte zuerst entpacken:
echo Rechtsklick auf die ZIP-Datei, "Alle extrahieren...", dann im entpackten Ordner ExcelAddin.bat doppelklicken.
pause
exit /b 1
