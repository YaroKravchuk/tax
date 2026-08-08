@echo off
setlocal
title Prospect LLC - Driver Logs

call "%~dp0find_python.bat"
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
