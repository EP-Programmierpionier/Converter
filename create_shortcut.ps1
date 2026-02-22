[CmdletBinding()]
param(
	[string]$ShortcutName = 'NWG-Bericht Converter',
	[string]$DesktopPath = [Environment]::GetFolderPath('Desktop')
)

$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

function Resolve-ExePath {
	$candidates = @(
		(Join-Path $PSScriptRoot '..\NWG-Bericht-Converter.exe'),
		(Join-Path $PSScriptRoot '.\NWG-Bericht-Converter.exe'),
		(Join-Path $PSScriptRoot '.\Release\NWG-Bericht-Converter.exe'),
		(Join-Path $PSScriptRoot '.\dist\NWG-Bericht-Converter.exe')
	)

	foreach ($path in $candidates) {
		if (Test-Path $path) { return (Resolve-Path $path).Path }
	}

	return $null
}

$exePath = Resolve-ExePath
if (-not $exePath) {
	Write-Host 'Keine .exe gefunden.' -ForegroundColor Yellow
	Write-Host 'Erst builden (build.bat oder python build_app.py), dann erneut ausführen.'
	exit 1
}

$iconPath = Join-Path $PSScriptRoot '.\Vorlagen\Converter_logo.ico'
if (-not (Test-Path $iconPath)) {
	$iconPath = $null
}

$shortcutPath = Join-Path $DesktopPath ("$ShortcutName.lnk")

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $exePath
$shortcut.WorkingDirectory = Split-Path -Path $exePath -Parent
if ($iconPath) {
	$shortcut.IconLocation = $iconPath
}
$shortcut.Save()

Write-Host "Shortcut erstellt: $shortcutPath" -ForegroundColor Green
