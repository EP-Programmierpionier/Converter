@echo off
setlocal
cd /d "%~dp0"

echo === NWG-Bericht Converter: Build ===
echo.

set "PY="
set "VENV_PY=%~dp0.venv\Scripts\python.exe"

if exist "%VENV_PY%" (
  "%VENV_PY%" -c "import sys" >nul 2>nul
  if not errorlevel 1 (
    set "PY=%VENV_PY%"
  )
)

if "%PY%"=="" (
  where py >nul 2>nul && set "PY=py -3.11"
)

if "%PY%"=="" (
  where python >nul 2>nul && set "PY=python"
)

if "%PY%"=="" (
  echo FEHLER: Python wurde nicht gefunden.
  echo Bitte Python installieren oder .venv anlegen.
  pause
  exit /b 1
)

%PY% build_app.py

echo.
pause
