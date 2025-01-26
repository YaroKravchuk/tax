# Dump Trucking Invoice Generator

A Python application that automates the creation of driver logs and invoices for dump trucking operations. Users input project details through a simple GUI, and the app generates formatted Excel workbooks containing itemized driver logs and corresponding invoices.

## Installation

1. Install Python 3.x from python.org

2. Install required packages:
```bash
pip install pandas PySimpleGUI openpyxl Pillow
```

You may need to use pip3 instead of pip

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
    - Click Submit

3. The program generates:
    - `Driver_Logs_[ProjectID].xlsx`
    - `Invoice_[ProjectID].xlsx`
