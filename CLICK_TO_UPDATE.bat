@echo off
setlocal
title Prospect LLC - Run Updates

REM ---------------------------------------------------------------------------
REM This script replaces itself, which needs care.
REM
REM Windows reads a batch file a piece at a time while it runs, remembering how
REM far into the file it has read. The "git pull" below can rewrite this very
REM file, and Windows then carries on reading the new file from the old
REM position, landing in the middle of a line and trying to run it. That is
REM where errors like this come from:
REM
REM     'connected' is not recognized as an internal or external command
REM
REM So before doing anything, this copies itself into the temp folder and hands
REM over to that copy. The file being run is then never the file being replaced.
REM ---------------------------------------------------------------------------

REM The folder holding the program, without the trailing backslash, which would
REM otherwise escape the closing quote when passed on
set "HERE=%~dp0"
set "HERE=%HERE:~0,-1%"

REM Being given a folder means this is already the temp copy, so get on with it
if "%~1"=="" goto :stage_copy
set "REPO=%~1"
goto :run_update

REM Hand over to the temp copy. This is deliberately not a "call": a call would
REM come back here afterwards, into the file git pull had just replaced, which
REM is the very problem being avoided. Handing over for good never returns.
:stage_copy
copy /Y "%~f0" "%TEMP%\prospect_llc_update.bat" >nul 2>&1
if not exist "%TEMP%\prospect_llc_update.bat" goto :no_temp_copy
"%TEMP%\prospect_llc_update.bat" "%HERE%"
exit /b

REM No temp copy could be made, so carry on from here instead. The update still
REM works, and only the rare pull that changes this file could interrupt it.
:no_temp_copy
set "REPO=%HERE%"

REM If the program folder did not come through, stop with something useful rather
REM than a puzzling complaint about Python. This is what the very first run after
REM an older version of this script updated it can look like.
:run_update
if exist "%REPO%\find_python.bat" goto :folder_ok
echo.
echo    PROBLEM: The program folder could not be found.
echo.
echo    This can happen on the first run straight after an update.
echo    Close this window and double-click CLICK_TO_UPDATE.bat again.
echo.
pause
exit /b 1

:folder_ok
cd /D "%REPO%"

echo ==================================================
echo    Updating the Dump Trucking program
echo ==================================================
echo.

REM ---- Find Python the same way CLICK_TO_RUN.bat does ----
call "%REPO%\find_python.bat"
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
    echo    or this computer is not on the internet.
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
    echo    Check that this computer is on the internet.
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
