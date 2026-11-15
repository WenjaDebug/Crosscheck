@echo off
rem Startet Gegenprobe per Doppelklick im Explorer. Beim ersten Mal wird alles installiert.
rem Alle Pfade in Anfuehrungszeichen: Der Ordner darf Leerzeichen und Umlaute haben (C:\Users\Juerg Mueller\...).
rem pushd statt cd: geht auch auf einem Netzlaufwerk (\\server\...).
pushd "%~dp0"
if not exist "%~dp0scripts\start-windows.ps1" goto nicht_entpackt
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\start-windows.ps1" %*
if errorlevel 1 pause
popd
exit /b

:nicht_entpackt
echo Gegenprobe steckt noch in der ZIP-Datei. Bitte zuerst entpacken:
echo Rechtsklick auf die ZIP-Datei, "Alle extrahieren...", dann im entpackten Ordner VoiceAgent.bat doppelklicken.
pause
exit /b 1
