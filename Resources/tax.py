from sheet_manager import SheetManager
from utility import collect_UI_input
import PySimpleGUI
import traceback
import os
import sys

try:
    # Create UI and collect user input from form.
    # Then, create materials such as template sheets and new files
    (project_id, driver_log_wb, invoice_wb, invoice_sheet, driver_log_template, data, taxable, should_create_driver_logs,
     should_create_invoice, min_date, max_date) = collect_UI_input()

    # Loop over data in BookRecords and add data to driver log sheets and invoice sheet
    sheet_manager = SheetManager(driver_log_wb, driver_log_template, invoice_sheet, taxable)
    for index, row in data.iterrows():
        if should_create_invoice:
            sheet_manager.populate_invoice_sheet_row(row)
        if should_create_driver_logs:
            sheet_manager.populate_driver_log_sheet(row)

    # Format and save invoice sheet
    if should_create_invoice:
        sheet_manager.merge_date_cells()
        sheet_manager.merge_truck_cells()
        invoice_wb.save(f'../INVOICE__{project_id}__{min_date} - {max_date}.xlsx')

    # Save driver logs sheet
    if should_create_driver_logs:
        driver_log_wb.save(f'../DRIVER LOGS__{project_id}__{min_date} - {max_date}.xlsx')

    if sheet_manager.row_count == 999:
        raise ValueError('There is too much data to fit on the invoice! '
                         '\n\nInvoice has been filled as much data as possible. The rest of the data is not included. '
                         '\n\nDecrease time range to avoid this issue...')

except Exception as e:
    tb = traceback.extract_tb(sys.exc_info()[2])
    for frame in reversed(tb):
        filename = os.path.basename(frame.filename)
        if filename in ['utility.py', 'tax.py', 'sheet_manager.py']:
            PySimpleGUI.PopupError(f"Error: {str(e)}\n\n\n\nFailed at this spot in the code: "
                                   f"\n\tFile: {filename} \n\tLine: {frame.lineno}")
            break
