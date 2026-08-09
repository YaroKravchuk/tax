@echo off
REM Works out which Python to use and hands it back in the PYTHON variable.
REM
REM CLICK_TO_RUN.bat and CLICK_TO_UPDATE.bat both call this, so the program always
REM runs on the same Python its packages were installed into. "py" is the Windows
REM Python launcher and picks the newest version installed, which keeps everything
REM working after Python is upgraded. Naming a version here, as this used to with
REM python3.13, breaks the moment a computer gets a newer Python.
REM
REM There is deliberately no setlocal in this file. It would throw the variables
REM below away again before the script that called us could read them.
REM
REM Hands back:
REM   PYTHON           how to start Python, empty if nothing suitable was found
REM   PYTHONW          how to start it without a console window, empty if there is
REM                    no such Python, or none of the same version as PYTHON
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
set "PYTHONW="
set "PYTHON_VERSION="
set "PYTHON_TOO_OLD="
set "PYTHON_UNTESTED="

REM Try each way of starting Python and keep the first that is new enough
for %%C in (py python python3) do call :consider %%C
if defined PYTHON call :consider_windowless
goto :eof


REM ---- Find the same Python again, in the form that opens no console ----------
REM Every Python on Windows comes with a second copy of itself built to run without
REM a console window: pyw beside py, pythonw beside python. CLICK_TO_RUN.bat uses it
REM so the program is not left sitting in front of an empty black window.
REM
REM It has to be the same version as the one chosen above, or the program would be
REM started by a Python that never had its packages installed - and, having no
REM console, would fail without a word on screen. Anything not confirmed to match is
REM handed back empty, which puts CLICK_TO_RUN.bat back on the ordinary Python
:consider_windowless
if /i "%PYTHON%"=="py" set "PYTHONW=pyw"
if /i "%PYTHON%"=="python" set "PYTHONW=pythonw"
if /i "%PYTHON%"=="python3" set "PYTHONW=pythonw"
if not defined PYTHONW goto :eof

REM Redirecting its output is what makes a windowless Python answer at all
set "WINDOWLESS_VERSION="
for /f "delims=" %%V in ('%PYTHONW% -c "import sys;print(str(sys.version_info[0])+chr(46)+str(sys.version_info[1]))" 2^>nul') do set "WINDOWLESS_VERSION=%%V"
if not "%WINDOWLESS_VERSION%"=="%PYTHON_VERSION%" set "PYTHONW="
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
