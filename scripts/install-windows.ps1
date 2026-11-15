# Installiert Gegenprobe unter Windows: eine eigene Python-Umgebung im Ordner .venv mit den drei Paketen.
# Aufruf: Doppelklick auf VoiceAgent.bat oder ExcelAddin.bat (oder: powershell -ExecutionPolicy Bypass -File scripts\install-windows.ps1)
#
# Python 3.10 oder neuer nimmt es, wo es schon eins findet. Fehlt eins, installiert es Python 3.13 mit winget (nur für
# diesen Benutzer, ohne Adminrechte); geht das nicht, öffnet es python.org und sagt, was zu tun ist.
# Ein python.exe unter ...\WindowsApps zählt nie: Ohne echtes Python öffnet es nur den Microsoft Store.
# Ein anderes Python lässt sich mit VOICEAGENT_PYTHON vorgeben, z. B. C:\Python313\python.exe.
# Alle Pfade mit -LiteralPath oder in Anführungszeichen: Der Ordner darf Leerzeichen, Umlaute und [ ] im Namen haben.
$ErrorActionPreference = "Stop"
Set-Location -LiteralPath (Split-Path -Parent $PSScriptRoot)

$PythonDownload = "https://www.python.org/downloads/windows/"

function Stop-Install([string]$message) {
    Write-Host ""
    Write-Host $message -ForegroundColor Yellow
    exit 1
}

trap {
    Stop-Install "Etwas ist schiefgegangen: $($_.Exception.Message)`nMach davon ein Bildschirmfoto, wenn du um Hilfe fragst. Hilfe: docs\HELP.md bzw. docs\HILFE.md im Ordner."
}

function Test-Python([string]$exe) {
    # Ein echtes Python 3.10 oder neuer?
    if (-not $exe -or $exe -like "*\WindowsApps\*" -or -not (Test-Path -LiteralPath $exe)) { return $false }
    try {
        # Windows PowerShell macht aus Fehlermeldungen von Python sonst einen Abbruch.
        $answer = & $exe -c "import sys; print(sys.version_info >= (3, 10))" 2>$null
        return ($LASTEXITCODE -eq 0 -and "$answer".Trim() -eq "True")
    }
    catch { return $false }
}

function Find-Python {
    if ($env:VOICEAGENT_PYTHON) { return $env:VOICEAGENT_PYTHON }
    $candidates = New-Object System.Collections.Generic.List[string]
    foreach ($version in @("313", "312", "311", "310", "314")) {
        $candidates.Add((Join-Path $env:LOCALAPPDATA "Programs\Python\Python$version\python.exe"))
        $candidates.Add((Join-Path $env:ProgramFiles "Python$version\python.exe"))
    }
    # Der „Python install manager“ von python.org legt Python hier ab.
    $managed = Join-Path $env:LOCALAPPDATA "Python"
    if (Test-Path -LiteralPath $managed) {
        Get-ChildItem -LiteralPath $managed -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "pythoncore-3.*" } | Sort-Object Name -Descending |
            ForEach-Object { $candidates.Add((Join-Path $_.FullName "python.exe")) }
    }
    foreach ($candidate in $candidates) {
        if (Test-Python $candidate) { return $candidate }
    }
    # Der Starter py vom Installer von python.org kennt jedes installierte Python (nicht der aus dem Store).
    $py = Get-Command py -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($py -and $py.Source -notlike "*\WindowsApps\*") {
        try {
            $found = & $py.Source -3 -c "import sys; print(sys.executable)" 2>$null
            if ($LASTEXITCODE -eq 0 -and $found -and (Test-Python "$found".Trim())) { return "$found".Trim() }
        }
        catch { }
    }
    foreach ($command in @(Get-Command python -All -CommandType Application -ErrorAction SilentlyContinue)) {
        if (Test-Python $command.Source) { return $command.Source }
    }
    return $null
}

function Install-Python {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { return }
    Write-Host "    Keins gefunden. Ich installiere Python 3.13 mit winget, nur für dich, ohne Adminrechte ..."
    try {
        & winget install --exact --id Python.Python.3.13 --silent --scope user --accept-package-agreements --accept-source-agreements
    }
    catch { }
}

function Show-PythonHelp {
    Write-Host ""
    Write-Host "Gegenprobe braucht Python (ein kostenloses Programm), und das fehlt auf diesem PC noch."
    Write-Host "Ich öffne dir jetzt die Seite python.org. So geht es weiter:"
    Write-Host "  1. Dort bei `"Python 3.13 ...`" auf `"Windows installer (64-bit)`" klicken."
    Write-Host "  2. Die geladene Datei doppelklicken. Hast du keine Adminrechte, unten das Häkchen bei"
    Write-Host "     `"Use admin privileges when installing py.exe`" wegnehmen. Dann auf `"Install Now`" klicken."
    Write-Host "  3. Danach VoiceAgent.bat (oder ExcelAddin.bat) noch einmal doppelklicken."
    try { Start-Process $PythonDownload } catch { Write-Host "Die Seite: $PythonDownload" }
}

function Test-Online {
    # Nur die Verbindung zu pypi.org, ohne etwas zu schicken.
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $ok = $client.ConnectAsync("pypi.org", 443).Wait(5000) -and $client.Connected
        $client.Close()
        return $ok
    }
    catch { return $false }
}

Write-Host "1/3 Ich suche Python 3.10 oder neuer ..."
$python = Find-Python
if (-not $python) {
    Install-Python
    $python = Find-Python
}
if (-not $python) {
    Show-PythonHelp
    if (-not (Test-Online)) {
        Write-Host ""
        Write-Host "Gerade komme ich übrigens nicht ins Internet. Das brauchst du dafür einmal."
    }
    exit 1
}
Write-Host "    $python"

Write-Host "2/3 Ich lege eine eigene Python-Umgebung im Ordner .venv an ..."
if (-not (Test-Path -LiteralPath ".venv\Scripts\python.exe")) {
    & $python -m venv .venv
    if ($LASTEXITCODE -ne 0) { Stop-Install "Das ging nicht. Ist der Ordner schreibgeschützt oder die Festplatte voll?" }
}

Write-Host "3/3 Ich lade und installiere Gegenprobe und ihre drei Pakete, das dauert eine Minute ..."
try { & .venv\Scripts\python.exe -m pip install --upgrade pip --quiet --disable-pip-version-check 2>$null | Out-Null } catch { }
& .venv\Scripts\python.exe -m pip install -e . --quiet --disable-pip-version-check
if ($LASTEXITCODE -ne 0) {
    if (-not (Test-Online)) {
        Stop-Install ("Ich komme nicht ins Internet. Beim ersten Start lädt Gegenprobe einmal ihre drei Pakete (FastAPI, Uvicorn, pypdf).`n" +
            "Danach geht alles ohne Internet. Verbinde den PC mit dem Internet und doppelklick die Datei noch einmal.")
    }
    Stop-Install "Die Pakete ließen sich nicht installieren. Was oben steht, hilft beim Fragen (Bildschirmfoto). Hilfe: docs\HILFE.md im Ordner."
}

Write-Host ""
Write-Host "Installiert."
exit 0
