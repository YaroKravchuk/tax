@echo off
REM Works out which Python to use and hands it back in the PYTHON variable.
REM
REM CLICK_TO_RUN.bat, here beside it, and CLICK_TO_UPDATE.bat, in the folder above,
REM both call this, so the program always runs on the same Python its packages were
REM installed into. "py" is the Windows Python launcher and picks the newest version
REM installed, which keeps everything working after Python is upgraded. Naming a
REM version here, as this used to with python3.13, breaks the moment a computer
REM gets a newer Python.
REM
REM There is deliberately no setlocal in this file. It would throw the variables
REM below away again before the script that called us could read them.
REM
REM Hands back:
REM   PYTHON           how to start Python, empty if nothing suitable was found
REM   PYTHON_VERSION   the version being used, e.g. 3.13
REM   PYTHON_TOO_OLD   version that was found but is below the minimum, if any
REM   PYTHON_UNTESTED  version being used, when it is newer than has been tried

REM Oldest Python the packages will install on. numpy needs 3.10, pandas and
REM Pillow need 3.9, openpyxl needs 3.8, so 3.10 is the real floor. The program's
REM own code is far less fussy and would run on much older versions.
set "PYTHON_MIN_MAJOR=3"
set "PYTHON_MIN_MINOR=10"

REM Newest Python this has been tried on. Going above it is only mentioned, never
REM blocked: a version that works must not be locked out the way naming python3.13
REM once locked out everything newer. A brand new Python usually only causes
REM trouble because its packages have no ready-made downloads yet.
set "PYTHON_MAX_MAJOR=3"
set "PYTHON_MAX_MINOR=14"

set "PYTHON_MIN=%PYTHON_MIN_MAJOR%.%PYTHON_MIN_MINOR%"
set "PYTHON_MAX=%PYTHON_MAX_MAJOR%.%PYTHON_MAX_MINOR%"

set "PYTHON="
set "PYTHON_VERSION="
set "PYTHON_TOO_OLD="
set "PYTHON_UNTESTED="

REM Try each way of starting Python and keep the first that is new enough
for %%C in (py python python3) do call :consider %%C
goto :eof


REM ---- Look at one candidate, and take it if it is new enough -----------------
REM The Python below deliberately avoids quote characters, which would end the
REM command early, and avoids the greater-than and less-than signs, which Windows
REM would read as redirecting to a file. Comparing with min() and max() gets the
REM same answer using only characters that are safe to write here.
:consider
if defined PYTHON goto :eof

set "FOUND="
for /f "delims=" %%V in ('%1 -c "import sys;print(str(sys.version_info[0])+chr(46)+str(sys.version_info[1]))" 2^>nul') do set "FOUND=%%V"
if not defined FOUND goto :eof

%1 -c "import sys;sys.exit(0 if min(sys.version_info[:2],(%PYTHON_MIN_MAJOR%,%PYTHON_MIN_MINOR%))==(%PYTHON_MIN_MAJOR%,%PYTHON_MIN_MINOR%) else 1)" >nul 2>&1
if errorlevel 1 goto :too_old

set "PYTHON=%1"
set "PYTHON_VERSION=%FOUND%"

%1 -c "import sys;sys.exit(0 if max(sys.version_info[:2],(%PYTHON_MAX_MAJOR%,%PYTHON_MAX_MINOR%))==(%PYTHON_MAX_MAJOR%,%PYTHON_MAX_MINOR%) else 1)" >nul 2>&1
if errorlevel 1 set "PYTHON_UNTESTED=%FOUND%"
goto :eof


REM ---- Remember the too-old one, so the message can name it -------------------
:too_old
if not defined PYTHON_TOO_OLD set "PYTHON_TOO_OLD=%FOUND%"
goto :eof
