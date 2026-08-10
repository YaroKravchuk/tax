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
if exist "%REPO%\Resources\find_python.bat" goto :folder_ok
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
call "%REPO%\Resources\find_python.bat"
if defined PYTHON goto :python_ok

if defined PYTHON_TOO_OLD echo    PROBLEM: This computer has Python %PYTHON_TOO_OLD%, but the program needs %PYTHON_MIN% or newer.
if not defined PYTHON_TOO_OLD echo    PROBLEM: Python is not installed on this computer.
echo.
echo    Install Python %PYTHON_MIN% or newer from https://www.python.org/downloads/
echo    Tick "Add Python to PATH" while installing, then run this file again.
echo.
pause
exit /b 1

REM A Python newer than has been tried still gets used. It is only worth saying so
REM here, where packages are installed, because that is what a brand new Python
REM tends to trip up.
:python_ok
if not defined PYTHON_UNTESTED goto :version_noted
echo    Note: Python %PYTHON_UNTESTED% is newer than %PYTHON_MAX%, the newest tried so far.
echo    Carrying on. If installing below fails, this is the likely reason.
echo.

:version_noted

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

echo [Step 1 of 3] Downloading the newest version of the program...
echo.

REM Put .gitignore back in plain sight for as long as the pull takes. It is hidden
REM again once the pull is done, and Windows can refuse to let a hidden file be
REM overwritten, which would fail the pull for a reason nothing on screen explains
attrib -h "%REPO%\.gitignore" >nul 2>&1

git pull
if errorlevel 1 (
    echo.
    echo    PROBLEM: The newest version could not be downloaded.
    echo.
    echo    This usually means a file in this folder was edited by hand,
    echo    or this computer is not on the internet.
    echo.
    attrib +h "%REPO%\.gitignore" >nul 2>&1
    pause
    exit /b 1
)

REM ---- Keep the housekeeping file out of sight -------------------------------
REM .gitignore is what stops the driver logs and invoices - customer names, site
REM addresses and rates - from being uploaded to the public repository, so it has
REM to stay. It is of no interest to anyone using the program though, and Windows,
REM unlike a Mac, shows it in the folder like any other file. Hiding it puts it
REM where Git already keeps its own .git folder: still there, just not in the way.
REM This cannot go straight after the pull: attrib would replace the pull's own
REM result code before the check above had the chance to read it
attrib +h "%REPO%\.gitignore" >nul 2>&1

echo.
echo [Step 2 of 3] Installing everything the program needs...
echo.
%PYTHON% -m pip install --upgrade pip
%PYTHON% -m pip install --upgrade pandas ttkbootstrap openpyxl Pillow numpy pywin32
if errorlevel 1 (
    echo.
    echo    PROBLEM: Something could not be installed.
    echo    Check that this computer is on the internet.
    echo.
    pause
    exit /b 1
)

echo.
echo [Step 3 of 3] Making the shortcut that opens the program...
echo.
call :make_shortcut

echo.
echo ==================================================
echo    All up to date!
echo.
echo    Now double-click
if defined SHORTCUT_MADE echo    %SHORTCUT_LABEL% to make driver logs.
if not defined SHORTCUT_MADE echo    Resources\CLICK_TO_RUN.bat to make driver logs.
echo ==================================================
echo.

REM Nothing left to read but the line above, so this shuts itself rather than asking
REM to be dismissed. Long enough to take in that it worked, and a key press cuts it
REM short. Only this ending closes itself: every other way out of this script is
REM something going wrong, and those all wait, because the window shutting is what
REM would take the explanation with it
echo    This window closes on its own in a moment.
timeout /t 6 >nul
exit /b


REM ---- Build the shortcut that starts the program without a console window ----
REM Double-clicking a .bat makes Windows open a console before a single line of it
REM runs, so no batch file can start the program without one appearing. A shortcut
REM can: it names pythonw.exe, which is the copy of Python built as a window
REM program rather than a console one, and Explorer starts it directly. Nothing in
REM between, which is also why it opens quicker than the batch file it replaces -
REM none of the looking for Python done here is repeated at every start.
REM
REM It is built here, and not kept in the repository, because it holds the full path
REM to this computer's Python and to this very folder, both of which are different on
REM the next computer. Built again on every update, so upgrading Python mends it.
REM
REM Resources\CLICK_TO_RUN.bat is kept as the spare way in. Nothing here can report
REM a failure once the program is running, and that batch file, with its console,
REM remains the way to see what went wrong
:make_shortcut
REM Windows hides the .lnk ending, so the label is what is actually seen in the folder
REM - which is why the shortcut can carry the name the batch file used to have without
REM the two ever looking alike. The batch file itself now lives in Resources
set "SHORTCUT_LABEL=CLICK_TO_RUN"
set "SHORTCUT=%REPO%\%SHORTCUT_LABEL%.lnk"
set "SHORTCUT_MADE="

REM Clear out the one from last time first. Left in place, a shortcut still pointing
REM at a Python that has since been upgraded away would pass the check below and be
REM announced as freshly made, while doing nothing at all when double-clicked
del /f /q "%SHORTCUT%" >nul 2>&1

REM The name this shortcut was given before, cleared away so the folder is not left
REM with two of them, one of which nothing builds any more
del /f /q "%REPO%\Driver Logs.lnk" >nul 2>&1

REM Ask Python only for the folder it lives in, and put the name of its windowless
REM twin on the end here. Asking for the whole path would need quote characters in
REM the Python, which would end the command early, exactly as find_python.bat avoids
set "PYTHON_FOLDER="
for /f "delims=" %%D in ('%PYTHON% -c "import os,sys;print(os.path.dirname(sys.executable))" 2^>nul') do set "PYTHON_FOLDER=%%D"

if not defined PYTHON_FOLDER goto :no_shortcut
set "PYTHONW_PATH=%PYTHON_FOLDER%\pythonw.exe"
if not exist "%PYTHONW_PATH%" goto :no_shortcut

REM Start in is not decoration: the program writes its files to the folder above and
REM loads its templates from beside itself, so it has to begin in Resources, exactly
REM as CLICK_TO_RUN.bat arranges with its cd
powershell -NoProfile -Command "$s = (New-Object -ComObject WScript.Shell).CreateShortcut('%SHORTCUT%'); $s.TargetPath = '%PYTHONW_PATH%'; $s.Arguments = 'tax.py'; $s.WorkingDirectory = '%REPO%\Resources'; $s.Description = 'Prospect LLC driver logs and invoices'; $s.Save()" >nul 2>&1

if not exist "%SHORTCUT%" goto :no_shortcut
set "SHORTCUT_MADE=1"
echo    Made %SHORTCUT_LABEL%, which opens the program without a black window.
goto :eof

REM Not being able to make it costs nothing: the batch file in Resources still works,
REM and is what the message at the end will go on pointing at
:no_shortcut
echo    Note: the shortcut could not be made on this computer.
echo    Nothing is broken - use Resources\CLICK_TO_RUN.bat instead.
goto :eof
