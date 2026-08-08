@echo off
REM Works out which Python to use and hands it back in the PYTHON variable.
REM
REM CLICK_TO_RUN.bat and CLICK_TO_UPDATE.bat both call this, so the program always
REM runs on the same Python its packages were installed into. "py" is the Windows
REM Python launcher and picks the newest version installed, which keeps everything
REM working after Python is upgraded. Naming a version here, as this used to with
REM python3.13, breaks the moment a computer gets a newer Python.
REM
REM There is deliberately no setlocal in this file. It would throw the PYTHON
REM variable away again before the script that called us could read it.

set "PYTHON="
py --version >nul 2>&1 && set "PYTHON=py"
if not defined PYTHON (
    python --version >nul 2>&1 && set "PYTHON=python"
)
if not defined PYTHON (
    python3 --version >nul 2>&1 && set "PYTHON=python3"
)
