from sheet_manager import SheetManager
from utility import collect_UI_input
import PySimpleGUI
import traceback
import os
import sys

try:
    # Create UI and collect user input from form.
    # Then, create materials such as template sheets and new files
    project_id, driver_log_wb, invoice_wb, invoice_sheet, driver_log_template, data, taxable = collect_UI_input()

    # Loop over data in BookRecords and add data to driver log sheets and invoice sheet
    sheet_manager = SheetManager(driver_log_wb, driver_log_template, invoice_sheet, taxable)
    for index, row in data.iterrows():
        sheet_manager.populate_invoice_sheet_row(row)
        sheet_manager.populate_driver_log_sheet(row)
    sheet_manager.merge_date_cells()
    sheet_manager.merge_truck_cells()

    # Save files
    driver_log_wb.save('../Driver_Logs_' + project_id + '.xlsx')
    invoice_wb.save('../Invoice_' + project_id + '.xlsx')
except Exception as e:
    tb = traceback.extract_tb(sys.exc_info()[2])
    for frame in reversed(tb):
        filename = os.path.basename(frame.filename)
        if filename in ['utility.py', 'tax.py', 'sheet_manager.py']:
            PySimpleGUI.PopupError(f"Error: {str(e)}\n\n\nFailed at this spot in the code: "
                                   f"\nFile: {filename} \nLine: {frame.lineno}")
            break
