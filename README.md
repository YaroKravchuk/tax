# Dump Trucking Invoice Generator

A Python application that automates the creation of driver logs and invoices for dump trucking operations. Users input project details through a simple GUI, and the app generates formatted Excel workbooks containing itemized driver logs and corresponding invoices.

## Installation

### On Windows (the easy way)

1. Install Python 3.x from python.org, ticking "Add Python to PATH" while installing
2. Install Git from https://git-scm.com/download/win
3. Double-click **CLICK_TO_UPDATE.bat**

That downloads the newest version of the program and installs everything it needs.
Run it again any time to pick up the latest changes; no command prompt required.

### By hand

1. Install Python 3.x from python.org

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

1. Run the main script
```bash
python tax.py
```
You may need to use python3 instead of python

2. In the GUI:
    - Select the appropriate "Dump Trucking" year sheet
    - Enter the Project ID
    - Specify date range (optional)
    - Check "Taxable" if applicable
    - Leave "Also Save as PDF" checked to get a PDF next to each Excel file
    - Click Submit

3. The program generates:
    - `DRIVER LOGS__[ProjectID]__[dates].xlsx` and the matching `.pdf`
    - `INVOICE__[ProjectID]__[dates].xlsx` and the matching `.pdf`
