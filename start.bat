@echo off
rem Keepa screener - Windows launcher. Double-click this file to start.
setlocal EnableExtensions
cd /d "%~dp0"
title Keepa screener

echo.
echo   Keepa screener
echo   ==============
echo.

rem ---- 1. Find Python 3.9 or newer --------------------------------------
set "PY="
py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>&1 && set "PY=py -3"
if not defined PY (
    python -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>&1 && set "PY=python"
)
if not defined PY (
    echo   Python 3.9 or newer is not installed on this computer.
    echo.
    echo   1. The Python download page will open now. Download and run the installer.
    echo   2. IMPORTANT: on the first screen tick "Add python.exe to PATH".
    echo   3. When it has finished, double-click start.bat again.
    echo.
    start "" "https://www.python.org/downloads/windows/"
    pause
    exit /b 1
)

rem ---- 2. Private environment so nothing else on the PC is touched ---------
if not exist ".venv\Scripts\python.exe" (
    echo   First start: setting things up. This takes a minute or two, only once.
    echo.
    %PY% -m venv .venv
    if errorlevel 1 (
        echo   Could not create the Python environment in this folder.
        echo   Move the folder somewhere you can write to, such as Documents, and try again.
        pause
        exit /b 1
    )
)
set "VPY=.venv\Scripts\python.exe"

rem ---- 3. Install or update the packages when requirements.txt changes ----
fc /b requirements.txt ".venv\installed-requirements.txt" >nul 2>&1
if errorlevel 1 (
    echo   Installing the packages the screener needs...
    "%VPY%" -m pip install --disable-pip-version-check -q --upgrade pip >nul 2>&1
    "%VPY%" -m pip install --disable-pip-version-check -q -r requirements.txt
    if errorlevel 1 (
        echo.
        echo   Installing the packages failed. Check the internet connection and try again.
        echo   If you are on a work network, a firewall or proxy may be blocking it.
        pause
        exit /b 1
    )
    copy /y requirements.txt ".venv\installed-requirements.txt" >nul
    echo   Done.
    echo.
)

rem ---- 4. Start the app ---------------------------------------------------
"%VPY%" app.py
if errorlevel 1 (
    echo.
    echo   The screener stopped with an error, see the message above.
    pause
)
endlocal
