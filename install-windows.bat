@echo off
title QCAA Word Count Tool
cd /d "%~dp0"

echo.
echo ============================================================
echo   QCAA WORD COUNT TOOL
echo ============================================================
echo.

set "PY="
py -3 --version >nul 2>nul && set "PY=py -3"
if not defined PY python --version >nul 2>nul && set "PY=python"
if not defined PY python3 --version >nul 2>nul && set "PY=python3"

if not defined PY goto :no_python

%PY% setup.py %*
goto :end

:no_python
echo ERROR: Python 3 was not found.
echo.
echo Install Python 3.15 (current stable) from:
echo   https://www.python.org/downloads/
echo.
echo Tick "Add Python to PATH" during setup, then try again.
echo.
start "" "https://www.python.org/downloads/"

:end
echo.
pause