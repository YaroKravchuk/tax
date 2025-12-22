import pandas as pd
import PySimpleGUI
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font


# Function to create the layout of the UI
def create_layout():
    PySimpleGUI.ChangeLookAndFeel('GreenTan')
    form = PySimpleGUI.FlexForm('Audit', default_element_size=(40, 1))
    # Open the Excel file
    excel_file = pd.ExcelFile('../../Dump Trucking BookRecords - TEST.xlsx')
    # Get all sheet names
    sheet_names = [name for name in excel_file.sheet_names if 'Dump Trucking' in name]
    # Check if current year exists in sheet names, and if so, set it as default
    current_year = str(datetime.now().year)
    default_sheet = next((name for name in sheet_names if current_year in name), None)
    layout = [
        [PySimpleGUI.Text('Generate Audit Daily Load Tickets!', size=(30, 1), font=("Helvetica", 25))],
        [PySimpleGUI.Text('"Dump Trucking" Year Sheet')],
        [PySimpleGUI.Combo(sheet_names, size=(30, 1), default_value=default_sheet)],
        [PySimpleGUI.Text('Project ID')],
        [PySimpleGUI.InputText()],
        [PySimpleGUI.Text('Start Date')],
        [PySimpleGUI.InputText()],
        [PySimpleGUI.Text('End Date')],
        [PySimpleGUI.InputText()],
        [PySimpleGUI.Checkbox('Taxable')],
        [PySimpleGUI.Checkbox('Include Driver Logs', default=True)],
        [PySimpleGUI.Checkbox('Include Invoice', default=True)],
        [PySimpleGUI.Text('_' * 80)],
        [PySimpleGUI.Submit(), PySimpleGUI.Cancel()]
    ]
    return form, layout


def collect_UI_input():
    form, layout = create_layout()
    window = form.Layout(layout)
    button, values = window.Read()

    try:
        start_date = pd.to_datetime(values[2]) if values[2] else pd.NaT
        end_date = pd.to_datetime(values[3]) if values[3] else pd.NaT
    except ValueError:
        raise ValueError("Invalid date format. Please enter a valid date (e.g. MM/DD/YYYY, YYYY-MM-DD)")

    return create_materials(
        values[0],     # Sheet Name
        values[1],     # Project ID
        start_date,    # Start Date
        end_date,      # End Date
        values[4],     # Taxable checkbox
        values[5],     # Driver Log checkbox
        values[6]      # Invoice checkbox
    )


# Function to create new materials like workbooks, the data table, and template sheets
def create_materials(sheet_name, project_id, start_date, end_date, taxable, should_create_driver_logs, should_create_invoice):
    driver_log_wb = load_workbook(filename='MASTER_DumpTruck_TimeSheet_ProspectLLC_2025_FINAL.xlsx')
    driver_log_template = driver_log_wb["2024 Version"]
    driver_log_template['B2'].font = Font(name='Calibri', color='FFFFFF', size=18, b=True)

    bold_cells = "G1 K1 A6 F6 N6 A8 B8 E8 H8 J8 K8 L8 N8 P8 A16 A17 A18 A20 A21 F16 J16 J17 J18 J20 J21".split()
    for cell in bold_cells:
        driver_log_template[cell].font = Font(name='Calibri', color='FFFFFF', size=11.5, b=True)

    df = pd.read_excel('../../Dump Trucking BookRecords - TEST.xlsx', sheet_name=sheet_name)

    df["DATE"] = pd.to_datetime(df["DATE"])

    # DEBUG: Print all unique project IDs with their repr() to see hidden characters
    print("\n=== DEBUG: All unique PROJECT IDs in dataframe ===")
    unique_ids = sorted(df["PROJECT ID"].unique(), key=str)
    for uid in unique_ids:
        print(f"ID: '{uid}' | repr: {repr(uid)} | type: {type(uid)}")

    print(f"\n=== DEBUG: Input project_id ===")
    print(f"ID: '{project_id}' | repr: {repr(project_id)} | type: {type(project_id)}")

    # Optional: Strip whitespace from both
    df["PROJECT ID"] = df["PROJECT ID"].astype(str).str.strip()
    project_id = str(project_id).strip()

    # Filter data to only include data for project ID that matches the date range
    data = df[df["PROJECT ID"] == project_id]
    print(f"\n=== DEBUG: Rows matched: {len(data)} ===")
    print(data)
    # Apply date filters if they are defined
    if pd.notna(start_date):
        data = data[data["DATE"] >= start_date]
    if pd.notna(end_date):
        data = data[data["DATE"] <= end_date]

    print("now data after date filter: ")
    print(data)
    validate_data(data, project_id)

    invoice_template = 'InvoiceASAP_Template_2025_NonTaxable.xlsx' if not taxable else 'InvoiceASAP_Template_2025.xlsx'
    invoice_wb = load_workbook(filename=invoice_template)
    invoice_wb["Blank_Template"].title = "Invoice"
    invoice_sheet = invoice_wb["Invoice"]
    min_date = data['DATE'].min()
    max_date = data['DATE'].max()
    invoice_sheet["D5"] = project_id
    invoice_sheet["G2"] = f"Start: {min_date.strftime('%m/%d/%y')}"
    invoice_sheet["H2"] = f"End: {max_date.strftime('%m/%d/%y')}"

    # Delete template sheet from final workbook file
    del driver_log_wb['2024 Version']

    return (project_id, driver_log_wb, invoice_wb, invoice_sheet, driver_log_template, data, taxable,
            should_create_driver_logs, should_create_invoice, min_date.strftime('%b %d').upper(),
            max_date.strftime('%b %d').upper())


# Function to check if BookRecords data exists and contains the correct columns
def validate_data(data, project_id):
    if data.empty:
        raise ValueError("No data found for project ID " + project_id)

    required_columns = ["PROJECT ID", "DATE", "TRUCK ID#", "PRODUCT", "LOAD QTY \n",
                        "CUSTOMER", "HAULING FROM", "HAULING TO"]
    missing_columns = [col for col in required_columns if col not in data.columns]
    if missing_columns:
        raise ValueError(f"Required data columns are missing: {', '.join(missing_columns)}")
