@echo off
setlocal
title QCAA Word Count Tool
cd /d "%~dp0"

echo.
echo ============================================================
echo   QCAA WORD COUNT TOOL
echo ============================================================
echo.

REM --- Check Python is available -------------------------------------
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
        pause
        exit /b 1
    )
)

REM --- Run the tool --------------------------------------------------
echo Starting...
echo.
".venv\Scripts\python.exe" qcaa_word_count.py
if errorlevel 1 (
    echo.
    echo The tool exited with an error. See the message above.
)

echo.
pause
endlocal