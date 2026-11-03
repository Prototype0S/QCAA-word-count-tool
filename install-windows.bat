@echo off
setlocal EnableExtensions EnableDelayedExpansion
title QCAA Word Count Tool

chcp 65001 >nul 2>nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

pushd "%~dp0" >nul 2>nul
if errorlevel 1 (
    echo ERROR: Could not enter the launcher folder.
    goto :close
)

echo.
echo ============================================================
echo   QCAA WORD COUNT TOOL
echo ============================================================
echo.

if not exist "qcaa_word_count.py" (
    echo ERROR: qcaa_word_count.py is missing.
    echo        Keep this launcher in the same folder as the tool.
    goto :close
)
if not exist "qcaa_word_count\NUL" (
    echo ERROR: The qcaa_word_count\ folder is missing.
    echo        Keep the whole folder together.
    goto :close
)

echo "!CD!" | findstr /i /c:"\Downloads" /c:"\Temp" >nul
if not errorlevel 1 (
    echo NOTE: This folder is in Downloads or a temporary location.
    echo       Move it to Documents so your virtual environment is kept.
    echo.
)

REM --- Find Python 3.10+ ------------------------------------------
set "PY_CMD="
call :try_python "py -3"
if not defined PY_CMD call :try_python "python"
if not defined PY_CMD call :try_python "python3"
if not defined PY_CMD goto :no_python

%PY_CMD% --version >nul 2>nul
if errorlevel 1 goto :no_python

for /f "tokens=2" %%v in ('%PY_CMD% --version 2^>^&1') do set "PYVER=%%v"
if not defined PYVER goto :no_python
echo Found Python !PYVER! (!PY_CMD!)

REM --- Create / repair the virtual environment --------------------
set "VENV_PY=.venv\Scripts\python.exe"
set "VENV_OK=0"
if exist "!VENV_PY!" (
    "!VENV_PY!" -c "import sys" >nul 2>nul
    if not errorlevel 1 set "VENV_OK=1"
)

if "!VENV_OK!"=="0" (
    echo Setting up virtual environment...
    if exist ".venv" rmdir /s /q ".venv" >nul 2>nul
    %PY_CMD% -m venv .venv
    if errorlevel 1 goto :venv_failed
    if not exist "!VENV_PY!" goto :venv_failed
    "!VENV_PY!" -c "import sys" >nul 2>nul
    if errorlevel 1 goto :venv_failed
)

REM --- Install dependencies ---------------------------------------
"!VENV_PY!" -c "import docx, docx.opc" >nul 2>nul
if errorlevel 1 (
    echo Installing dependencies ^(python-docx^)...
    "!VENV_PY!" -m pip install --quiet --upgrade pip >nul 2>nul
    "!VENV_PY!" -m pip install --quiet python-docx
    if errorlevel 1 goto :pip_failed
)

REM --- Smoke test -------------------------------------------------
REM Quiet on success, loud on failure. Catches "installer worked
REM but the tool won't run" before the user hits a traceback.
if exist "selftest.py" (
    "!VENV_PY!" selftest.py >nul 2>nul
    if errorlevel 1 (
        echo.
        echo WARNING: The tool's own sanity check failed.
        echo          The installer worked but the tool may not run.
        echo          Details:
        echo.
        "!VENV_PY!" selftest.py
        echo.
        set "TOOL_EXIT=1"
        goto :close
    )
)

REM --- Run --------------------------------------------------------
echo Starting...
echo.
"!VENV_PY!" qcaa_word_count.py %*
set "TOOL_EXIT=!errorlevel!"
if not "!TOOL_EXIT!"=="0" (
    echo.
    echo The tool exited with an error ^(code !TOOL_EXIT!^). See above.
)

goto :close

:try_python
%~1 -c "import sys; raise SystemExit(0 if sys.version_info >= (3,10) else 1)" >nul 2>nul
if not errorlevel 1 set "PY_CMD=%~1"
exit /b 0

:no_python
echo.
echo ERROR: Python 3.10 or newer was not found.
echo.
echo   Install Python 3.15 ^(current stable^) from:
echo       https://www.python.org/downloads/
echo   Pick "Windows installer ^(64-bit^)".
echo.
echo   1. Run the installer.
echo   2. On the first screen, tick "Add Python to PATH".
echo   3. Double-click this file again.
echo.
start "" "https://www.python.org/downloads/"
set "TOOL_EXIT=1"
goto :close

:venv_failed
echo.
echo ERROR: Could not create the virtual environment.
echo.
echo   Reinstall Python 3.15 from:
echo       https://www.python.org/downloads/
echo   During setup, tick "Add Python to PATH".
echo.
set "TOOL_EXIT=1"
goto :close

:pip_failed
echo.
echo ERROR: Could not install python-docx.
echo.
echo   Check your internet connection. On a school network, PyPI may
echo   be blocked; try a phone hotspot or ask IT for the proxy address.
echo.
set "TOOL_EXIT=1"
goto :close

:close
popd >nul 2>nul
echo.
pause
endlocal & exit /b %TOOL_EXIT%