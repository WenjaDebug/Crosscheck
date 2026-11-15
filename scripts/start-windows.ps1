# Startet Gegenprobe unter Windows. Beim ersten Mal wird alles installiert (install-windows.ps1), danach schaut es
# vor jedem Start nach einer neuen Version: Ein Git-Klon holt sie, eine ZIP-Installation sagt nur, dass es eine gibt.
# Aufruf: Doppelklick auf VoiceAgent.bat
$ErrorActionPreference = "Stop"
Set-Location -LiteralPath (Split-Path -Parent $PSScriptRoot)

if (-not (Test-Path -LiteralPath ".venv\Scripts\voiceagent.exe")) {
    Write-Host "Gegenprobe wird zum ersten Mal eingerichtet."
    & (Join-Path $PSScriptRoot "install-windows.ps1")
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "Noch nicht fertig, oben steht, was zu tun ist." -ForegroundColor Yellow
        exit 1
    }
    Write-Host ""
    Write-Host "Gleich öffnet sich Gegenprobe im Browser. Für das Seitenfenster in Excel doppelklick danach ExcelAddin.bat."
}

# Vor jedem Start nach einer neuen Version schauen (ausschalten: VOICEAGENT_NO_UPDATE=1). Scheitert das, startet die bisherige.
& .venv\Scripts\python.exe -m voiceagent.update

Write-Host "Gegenprobe startet. Zum Beenden dieses Fenster schließen oder Strg + C drücken."
& .venv\Scripts\voiceagent.exe @args
exit $LASTEXITCODE
