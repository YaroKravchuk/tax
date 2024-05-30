import pandas as pd
from utility import create_layout, create_images, create_workbook, validate_data, populate_data_sheet, populate_driver_log_sheet
from openpyxl.worksheet.table import Table, TableStyleInfo
import PySimpleGUI as sg

try:
    form, layout = create_layout()
    wb, ws, source = create_workbook()

    window = form.Layout(layout)
    button, values = window.Read()
    sheet_name, project_id, start_date, end_date = values[0], values[1], pd.to_datetime(values[2]), pd.to_datetime(values[3])
    df = pd.read_excel('../../Dump Trucking BookRecords - TEST.xlsx', sheet_name=sheet_name)
    data = df[df["PROJECT ID"] == project_id]

    validate_data(data)

    prev_truck = ""
    load_row_count = 0
    row_count = 2
    prev_date = ""
    prev_id = ""
    for index, row in data.iterrows():
        if row["DATE"] < start_date or row["DATE"] > end_date:
            continue

        populate_data_sheet(ws, row_count, row)
        row_count = row_count + 1

        # First check if previous sheet is same truck, same day
        if row["DATE"] == prev_date and row["TRUCK ID#"] == prev_truck:
            target = wb[prev_id]
            load_row_count = load_row_count + 1
        else:
            target = wb.copy_worksheet(source)
            target.title = "Form " + str(index + 1)
            prev_id = "Form " + str(index + 1)
            prev_date = row["DATE"]
            prev_truck = row["TRUCK ID#"]
            target['I2'] = row["DATE"]
            target['B6'] = row["CUSTOMER"]
            target['L6'] = row["TRUCK ID#"]
            logoImage, signatureImage = create_images()
            target.add_image(logoImage)
            target.add_image(signatureImage)
            load_row_count = 0

        populate_driver_log_sheet(target, load_row_count, row)

        if not pd.isnull(row["TIME IN"]):
            target['O17'] = row["TIME IN"]
        if not pd.isnull(row["TIME OUT"]):
            target['O18'] = row["TIME OUT"]
        if row["HOURS"] > 0:
            target['O19'] = row["HOURS"]

    del wb['Sheet1']
    tab = Table(displayName="Table1", ref="A1:D" + str(row_count - 1))
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    ws.add_table(tab)

    wb.save('../Driver_Logs_' + project_id + '.xlsx')
except Exception as e:
    sg.PopupError(f"An error occurred: {e}")
