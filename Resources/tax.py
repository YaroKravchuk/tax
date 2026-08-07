from sheet_manager import SheetManager
from pdf_exporter import export_to_pdf
from utility import collect_UI_input
import FreeSimpleGUI as PySimpleGUI
import traceback
import os
import sys

try:
    # Create UI and collect user input from form.
    # Then, create materials such as template sheets and new files
    (project_id, driver_log_wb, invoice_wb, invoice_sheet, driver_log_template, data, taxable, should_create_driver_logs,
     should_create_invoice, should_export_pdf, min_date, max_date) = collect_UI_input()

    # Loop over data in BookRecords and add data to driver log sheets and invoice sheet
    sheet_manager = SheetManager(driver_log_wb, driver_log_template, invoice_sheet, taxable)
    for index, row in data.iterrows():
        if should_create_invoice:
            sheet_manager.populate_invoice_sheet_row(row)
        if should_create_driver_logs:
            sheet_manager.populate_driver_log_sheet(row)

    saved_files = []

    # Format and save invoice sheet
    if should_create_invoice:
        sheet_manager.merge_date_cells()
        sheet_manager.merge_truck_cells()
        invoice_file = f'../INVOICE__{project_id}__{min_date} - {max_date}.xlsx'
        invoice_wb.save(invoice_file)
        saved_files.append(invoice_file)

    # Save driver logs sheet
    if should_create_driver_logs:
        driver_log_file = f'../DRIVER LOGS__{project_id}__{min_date} - {max_date}.xlsx'
        driver_log_wb.save(driver_log_file)
        saved_files.append(driver_log_file)

    # Save a PDF copy of each Excel file that was created. The Excel files are already saved at
    # this point, so a failed PDF export is reported on its own instead of failing the whole run
    if should_export_pdf:
        pdf_errors = []
        for saved_file in saved_files:
            try:
                export_to_pdf(saved_file)
            except Exception as pdf_error:
                pdf_errors.append(f'{os.path.basename(saved_file)}\n{str(pdf_error)}')
        if pdf_errors:
            PySimpleGUI.PopupError('The Excel files were created, but the PDF export failed:'
                                   '\n\n' + '\n\n'.join(pdf_errors))

    if sheet_manager.row_count == 999:
        raise ValueError('There is too much data to fit on the invoice! '
                         '\n\nInvoice has been filled as much data as possible. The rest of the data is not included. '
                         '\n\nDecrease time range to avoid this issue...')

except Exception as e:
    tb = traceback.extract_tb(sys.exc_info()[2])
    for frame in reversed(tb):
        filename = os.path.basename(frame.filename)
        if filename in ['utility.py', 'tax.py', 'sheet_manager.py', 'pdf_exporter.py']:
            PySimpleGUI.PopupError(f"Error: {str(e)}\n\n\n\nFailed at this spot in the code: "
                                   f"\n\tFile: {filename} \n\tLine: {frame.lineno}")
            break
