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
echo %~dp0 | findstr /i /c:"\Downloads\" /c:"\Temp\" >nul
if %errorlevel% equ 0 (
    echo NOTE: You are running the tool from Downloads or Temp.
    echo       It is better to move the folder somewhere permanent
    echo       like Documents\qcaa-word-count first.
    echo.
)

REM --- Check Python is available -------------------------------------
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not on PATH.
    echo.
    echo   RECOMMENDED: Python 3.15 (much faster for this tool)
    echo   Download:    https://www.python.org/downloads/release/python-3150/
    echo   Pick:        "Windows installer (64-bit)"
    echo.
    echo   1. Run the installer.
    echo   2. On the first screen, tick "Add Python to PATH".
    echo   3. Then double-click this file again.
    echo.
    pause
    exit /b 1
)

for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo Found Python %PYVER%

REM --- Recommend 3.15, but require 3.10 -------------------------------
python -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python %PYVER% is too old. This tool needs Python 3.10 or newer.
    echo.
    echo   RECOMMENDED: Python 3.15 (much faster for this tool)
    echo   Download:    https://www.python.org/downloads/release/python-3150/
    echo.
    pause
    exit /b 1
)

python -c "import sys; sys.exit(0 if sys.version_info >= (3,15) else 1)" >nul 2>nul
if errorlevel 1 (
    echo.
    echo NOTE: You are on Python %PYVER%. The tool works, but 3.15 is
    echo       noticeably faster. If you want the speedup, upgrade from:
    echo       https://www.python.org/downloads/release/python-3150/
    echo.
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