@echo off
setlocal
title Prospect LLC - Driver Logs

call "%~dp0Resources\find_python.bat"
if defined PYTHON goto :python_ok

echo.
if defined PYTHON_TOO_OLD echo    PROBLEM: This computer has Python %PYTHON_TOO_OLD%, but the program needs %PYTHON_MIN% or newer.
if not defined PYTHON_TOO_OLD echo    PROBLEM: Python is not installed on this computer.
echo.
echo    Install Python %PYTHON_MIN% or newer from https://www.python.org/downloads/
echo    Tick "Add Python to PATH" while installing, then run this file again.
echo.
pause
exit /b 1

:python_ok
cd /D "%~dp0Resources"

REM ---- Get out of the way, if this computer has a Python that lets us ---------
REM Double-clicking a .bat opens a console window, and it stays for as long as the
REM batch file runs: start the program from here in the ordinary way and that black
REM window sits behind it the whole session. Handing it to the windowless Python
REM with "start" instead means this file has nothing left to wait for, so it ends
REM at once and takes the console with it. The window is still there while Python
REM is being looked for above, but it is gone by the time the form appears.
REM
REM Nothing is lost by not waiting. Everything that can go wrong once the program
REM is running says so in a window of its own, and tax.py puts up a message box of
REM its own if it cannot even load - which, with no console, is the one failure
REM that would otherwise happen in complete silence
if not defined PYTHONW goto :with_console
start "" "%PYTHONW%" tax.py
exit /b

REM No windowless Python to be had, so run it the old way and keep the console,
REM which is then the only place a message can appear
:with_console
%PYTHON% tax.py
if errorlevel 1 (
    echo.
    echo    The program stopped before it finished.
    echo.
    echo    If the message above mentions a missing module, close this window and
    echo    double-click CLICK_TO_UPDATE.bat to install what it needs, then try again.
    echo.
    pause
)
