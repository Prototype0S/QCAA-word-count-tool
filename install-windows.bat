@echo off
setlocal
title QCAA Word Count Tool
cd /d "%~dp0"

echo.
echo ============================================================
echo   QCAA WORD COUNT TOOL
echo ============================================================
echo.

REM --- Warn if running from Downloads or a temporary folder ---------
echo %~dp0 | findstr /i /c:"\\Downloads\\" /c:"\\Temp\\" >nul
if %errorlevel% equ 0 (
    echo NOTE: You are running the tool from Downloads or Temp.
    echo       It is better to move the folder somewhere permanent
    echo       like Documents\qcaa-word-count first.
    echo.
)

REM --- Check Python is available and recent enough -------------------
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not on PATH.
    echo.
    echo   1. Download Python 3.10+ from https://www.python.org/downloads/
    echo   2. On the first screen of the installer, tick
    echo      "Add Python to PATH"
    echo   3. Then double-click this file again.
    echo.
    pause
    exit /b 1
)

for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo Found Python %PYVER%
python -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python %PYVER% is too old. This tool needs Python 3.10 or newer.
    echo Install a newer version from https://www.python.org/downloads/
    pause
    exit /b 1
)

REM --- Create the virtual environment if needed ----------------------
if not exist ".venv\Scripts\python.exe" (
    echo Setting up virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo ERROR: Could not create virtual environment.
        pause
        exit /b 1
    )
)

REM --- Install dependencies if needed --------------------------------
".venv\Scripts\python.exe" -c "import docx" >nul 2>nul
if errorlevel 1 (
    echo Installing dependencies...
    ".venv\Scripts\python.exe" -m pip install --quiet --upgrade pip
    ".venv\Scripts\python.exe" -m pip install --quiet python-docx
    if errorlevel 1 (
        echo ERROR: Could not install python-docx.
        echo.
        echo If you are on a school network, PyPI might be blocked.
        echo Try from a home network or a phone hotspot.
        pause
        exit /b 1
    )
)

REM --- Run the tool --------------------------------------------------
echo Starting...
echo.
echo If the tool asks for a file, a file picker will open.
echo Navigate to your .docx file and click Open.
echo.
echo The HTML preview will open in your browser.
echo This window stays open until you press a key at the end.
echo.

if not "%~1"=="" (
    ".venv\Scripts\python.exe" qcaa_word_count.py "%~1"
) else (
    ".venv\Scripts\python.exe" qcaa_word_count.py
)
if errorlevel 1 (
    echo.
    echo The tool exited with an error. See the message above.
)

echo.
pause
endlocal