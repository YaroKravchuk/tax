# Dump Trucking Invoice Generator

A Python application that automates the creation of driver logs and invoices for dump trucking operations. Users input project details through a simple GUI, and the app generates formatted Excel workbooks containing itemized driver logs and corresponding invoices.

## Installation

### On Windows (the easy way)

1. Install Python 3.10 or newer from python.org, ticking "Add Python to PATH" while installing
2. Install Git from https://git-scm.com/download/win
3. Double-click **CLICK_TO_UPDATE.bat**

That downloads the newest version of the program, installs everything it needs, and
makes a **Driver Logs** shortcut in this folder to open it with.
Run it again any time to pick up the latest changes; no command prompt required.

### By hand

1. Install Python 3.10 or newer from python.org

Python 3.10 is the floor because numpy needs it; pandas and Pillow need 3.9, openpyxl 3.8.
Newer versions than that are always allowed. `Resources\find_python.bat` stops with a clear message
if the only Python it finds is too old, and merely mentions it if one is newer than has
been tried, since a working Python must never be locked out.

2. Install required packages:
```bash
pip install pandas ttkbootstrap openpyxl Pillow numpy
```

You may need to use pip3 instead of pip

3. Install what is needed for PDF output:

- On Windows, the PDF is made with Excel, which needs one extra package:
```bash
pip install pywin32
```
- On a Mac, Excel cannot be driven automatically, so the PDF is made with LibreOffice:
```bash
brew install --cask libreoffice
```

Without these the Excel files are still created, and only the PDF step is skipped with a message.

## Usage

1. Double-click the **Driver Logs** shortcut, or run the main script by hand:
```bash
python tax.py
```
You may need to use python3 instead of python

The shortcut is made by `CLICK_TO_UPDATE.bat`, so run that once before looking for it. It
points straight at `pythonw.exe`, the copy of Python built as a window program rather than
a console one, so the program opens on its own with no black window behind it. It is also
the quickest way in, since none of the looking for Python is repeated at every start.

**CLICK_TO_RUN.bat** still works and is the one to use when something is wrong: it keeps a
console open, which is where anything the program cannot report in a window of its own will
appear. It and `CLICK_TO_UPDATE.bat` find Python through `Resources\find_python.bat`, which
uses whichever version is installed, so neither stops working when Python is upgraded.

2. In the GUI:
    - Check the line at the top saying when the records were last saved. If edits were just
      made and it still says hours ago, the file being edited is not the file the program is
      reading. Saving the records while the form is open turns that line red, because the
      form is then working from older data
    - Select the appropriate "Dump Trucking" year sheet
    - Find the project. Click the Project box and the most recently worked projects drop
      down under it, each with its customer and load count beside it; typing narrows them,
      and they disappear again once one is picked. Any part of the address or of the
      customer works, in any order, and the two can be mixed: `delridge`, `86th seattle`,
      `all terrain` or `terrain auburn` are all enough. Pick one with the mouse, or
      with the arrow keys and Enter, and Escape puts the list away. The line underneath
      then confirms the customer, the number of loads and the dates, so the right project
      can be checked before anything is made
    - Adjust the date range. It is filled in with everything the project has, so narrow
      it only when part of the job is wanted. Dates can be typed, or picked from the
      calendar behind the button at the end of each box
    - Under "Create", choose whether to make the driver logs, the invoice, or both
    - Under "Options", check "Taxable" if applicable. It greys out when no invoice is
      being made, since it describes the invoice and nothing else. Leave "Also save as
      PDF" ticked to get a PDF next to each Excel file
    - Click Generate. Anything that stops the run is said against the box it belongs to,
      with that box outlined in red. A small window then says what is happening, and shows
      how far along it is, since building the files and making the PDFs takes a while

3. The program generates:
    - `DRIVER LOGS__[ProjectID]__[dates].xlsx` and the matching `.pdf`
    - `INVOICE__[ProjectID]__[dates].xlsx` and the matching `.pdf`
