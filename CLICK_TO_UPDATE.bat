@echo off
setlocal
title Prospect LLC - Run Updates
cd /D "%~dp0"

echo ==================================================
echo    Updating the Dump Trucking program
echo ==================================================
echo.

REM ---- Find Python the same way CLICK_TO_RUN.bat does ----
call "%~dp0find_python.bat"
if not defined PYTHON (
    echo    PROBLEM: Python is not installed on this computer.
    echo.
    echo    Install it from https://www.python.org/downloads/
    echo    Tick "Add Python to PATH" while installing, then run this file again.
    echo.
    pause
    exit /b 1
)

REM ---- Git is what downloads the update, so check it before starting ----
git --version >nul 2>&1
if errorlevel 1 (
    echo    PROBLEM: Git is not installed on this computer.
    echo.
    echo    Install it from https://git-scm.com/download/win
    echo    then run this file again.
    echo.
    pause
    exit /b 1
)

echo [Step 1 of 2] Downloading the newest version of the program...
echo.
git pull
if errorlevel 1 (
    echo.
    echo    PROBLEM: The newest version could not be downloaded.
    echo.
    echo    This usually means a file in this folder was edited by hand,
    echo    or this computer is not connected to the internet.
    echo.
    pause
    exit /b 1
)

echo.
echo [Step 2 of 2] Installing everything the program needs...
echo.
%PYTHON% -m pip install --upgrade pip
%PYTHON% -m pip install --upgrade pandas FreeSimpleGUI openpyxl Pillow numpy pywin32
if errorlevel 1 (
    echo.
    echo    PROBLEM: Something could not be installed.
    echo    Check that this computer is connected to the internet.
    echo.
    pause
    exit /b 1
)

echo.
echo ==================================================
echo    All up to date!
echo.
echo    Close this window, then double-click
echo    CLICK_TO_RUN.bat to make driver logs.
echo ==================================================
echo.
pause
