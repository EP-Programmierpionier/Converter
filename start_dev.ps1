[CmdletBinding()]
param(
    [switch]$SkipInstall
)

$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

Write-Host "=== NWG-Bericht Converter: Dev Start ===" -ForegroundColor Cyan

$venvPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'

function Test-VenvHealthy {
    if (-not (Test-Path $venvPython)) { return $false }

    $cfg = Join-Path $PSScriptRoot '.venv\pyvenv.cfg'
    if (-not (Test-Path $cfg)) { return $false }

    $cfgText = Get-Content -Path $cfg -Raw
    $exeLine = ($cfgText -split "`n" | Where-Object { $_ -match '^executable\s*=\s*' } | Select-Object -First 1)
    if (-not $exeLine) { return $false }

    $baseExe = ($exeLine -replace '^executable\s*=\s*', '').Trim()
    if ([string]::IsNullOrWhiteSpace($baseExe)) { return $false }

    return (Test-Path $baseExe)
}

if (-not (Test-VenvHealthy)) {
    Write-Host "Lege .venv an..." -ForegroundColor Yellow

    $py = $null
    if (Get-Command py -ErrorAction SilentlyContinue) { $py = 'py' }
    elseif (Get-Command python -ErrorAction SilentlyContinue) { $py = 'python' }

    if (-not $py) {
        throw 'Python wurde nicht gefunden (weder `py` noch `python`).'
    }

    if (Test-Path .\.venv) {
        Remove-Item .\.venv -Recurse -Force
    }

    # Bewusst über den Python Launcher, damit wir die gewünschte Version treffen.
    if ($py -eq 'py') {
        & py -3.11 -m venv .venv
    } else {
        & python -m venv .venv
    }
}

Write-Host "Aktiviere .venv..." -ForegroundColor Yellow
. (Join-Path $PSScriptRoot '.venv\Scripts\Activate.ps1')

if (-not $SkipInstall) {
    Write-Host "Installiere Abhängigkeiten..." -ForegroundColor Yellow
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
}

Write-Host "Starte Anwendung..." -ForegroundColor Green
python .\NWG_Converter.py
