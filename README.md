# Dump Trucking Invoice Generator

A Python application that automates the creation of driver logs and invoices for dump trucking operations. Users input project details through a simple GUI, and the app generates formatted Excel workbooks containing itemized driver logs and corresponding invoices.

## Installation

### On Windows (the easy way)

1. Install Python 3.10 or newer from python.org, ticking "Add Python to PATH" while installing
2. Install Git from https://git-scm.com/download/win
3. Double-click **CLICK_TO_UPDATE.bat**

That downloads the newest version of the program and installs everything it needs.
Run it again any time to pick up the latest changes; no command prompt required.

### By hand

1. Install Python 3.10 or newer from python.org

Python 3.10 is the floor because numpy needs it; pandas and Pillow need 3.9, openpyxl 3.8.
Newer versions than that are always allowed. `find_python.bat` stops with a clear message
if the only Python it finds is too old, and merely mentions it if one is newer than has
been tried, since a working Python must never be locked out.

2. Install required packages:
```bash
pip install pandas FreeSimpleGUI openpyxl Pillow numpy
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

1. Double-click **CLICK_TO_RUN.bat**, or run the main script by hand:
```bash
python tax.py
```
You may need to use python3 instead of python

Both `CLICK_TO_RUN.bat` and `CLICK_TO_UPDATE.bat` find Python through `find_python.bat`,
which uses whichever version is installed. They will not stop working when Python is upgraded.

2. In the GUI:
    - Select the appropriate "Dump Trucking" year sheet
    - Find the project. The box under Project ID searches as you type, and the list
      underneath shows what matches, most recently worked first. Any part of the address
      works, in any order: typing `delridge` or `86th seattle` is enough. Pick one with
      the mouse, or with the arrow keys and Enter. The line under the list confirms the
      customer, the number of loads and the dates, so the right project can be checked
      before anything is made
    - Adjust the date range. It is filled in with everything the project has, so narrow
      it only when part of the job is wanted
    - Check "Taxable" if applicable
    - Leave "Also Save as PDF" checked to get a PDF next to each Excel file
    - Click Submit. A small window then says what is happening, since building the files
      and making the PDFs takes a while

3. The program generates:
    - `DRIVER LOGS__[ProjectID]__[dates].xlsx` and the matching `.pdf`
    - `INVOICE__[ProjectID]__[dates].xlsx` and the matching `.pdf`
