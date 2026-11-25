# Prueft die Umgebung und gibt die Pfade aus, die der Agent braucht.
# Aufruf:  powershell -NoProfile -ExecutionPolicy Bypass -File app\check_env.ps1
# Laeuft auch im Constrained Language Mode der Schulrechner: keine .NET-Aufrufe, nur Cmdlets.
$app = Split-Path -Parent $MyInvocation.MyCommand.Path

$candidates = @()
if ($env:LIBREOFFICE_PROGRAM) { $candidates += (Join-Path $env:LIBREOFFICE_PROGRAM 'python.exe') }
$candidates += (Join-Path $env:ProgramFiles 'LibreOffice\program\python.exe')
if (${env:ProgramFiles(x86)}) { $candidates += (Join-Path ${env:ProgramFiles(x86)} 'LibreOffice\program\python.exe') }
$candidates += (Join-Path $env:LOCALAPPDATA 'Programs\LibreOffice\program\python.exe')

$python = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $python) {
    Write-Output "LIBREOFFICE_PYTHON=NICHT GEFUNDEN"
    Write-Output "LibreOffice ist nicht installiert oder liegt an einem unbekannten Ort."
    Write-Output "Abhilfe: LibreOffice installieren oder die Umgebungsvariable LIBREOFFICE_PROGRAM auf den Ordner ...\LibreOffice\program setzen."
    exit 1
}
Write-Output "LIBREOFFICE_PYTHON=$python"

$test = "import sys; sys.path.insert(0, r'$app'); import headless_core as h, uno; print('DOCUMENTS=' + h.documents_folder()); print('IMPORTS=OK'); print('SOFFICE=' + str(h.SOFFICE)); print('SOFFICE_FOUND=' + str(h.SOFFICE.is_file()))"
& $python -c $test
if ($LASTEXITCODE -ne 0) { Write-Output "IMPORTS=FEHLER"; exit 1 }
