from form import StatusWindow, collect_UI_input, report_error, report_finished
from sheet_manager import SheetManager
from pdf_exporter import export_to_pdf
from utility import create_materials, open_output_folder
import traceback
import os
import sys

# Move the status window on every this many loads, so a long run keeps showing progress
# without being repainted for every single row
PROGRESS_STEP = 25

# The program's own files, used to point at where an error came from
CODE_FOLDER = os.path.dirname(os.path.abspath(__file__))

status = None

try:
    # Create UI and collect user input from form
    user_input = collect_UI_input()

    # Cancel, or closing the window, leaves quietly without creating anything
    if user_input is None:
        sys.exit()

    # From here on the form is gone and the work takes a while, so a status window
    # keeps saying what is happening
    status = StatusWindow('Preparing')

    # Create materials such as template sheets and new files
    materials = create_materials(user_input)

    # Loop over data in BookRecords and add data to driver log sheets and invoice sheet
    sheet_manager = SheetManager(materials.driver_log_wb, materials.driver_log_template,
                                 materials.invoice_sheet, user_input.taxable)
    load_count = len(materials.data)
    for position, (index, row) in enumerate(materials.data.iterrows(), start=1):
        if position == 1 or position % PROGRESS_STEP == 0:
            status.update('Adding loads', done=position, total=load_count)
        if user_input.should_create_invoice:
            sheet_manager.populate_invoice_sheet_row(row)
        if user_input.should_create_driver_logs:
            sheet_manager.populate_driver_log_sheet(row)

    saved_files = []

    # Format and save invoice sheet
    if user_input.should_create_invoice:
        status.update('Saving invoice')
        sheet_manager.merge_date_cells()
        sheet_manager.merge_truck_cells()
        invoice_file = f'../INVOICE__{user_input.project_id}__{materials.min_date} - {materials.max_date}.xlsx'
        materials.invoice_wb.save(invoice_file)
        saved_files.append(invoice_file)

    # Save driver logs sheet
    if user_input.should_create_driver_logs:
        status.update('Saving driver logs')
        driver_log_file = (f'../DRIVER LOGS__{user_input.project_id}'
                           f'__{materials.min_date} - {materials.max_date}.xlsx')
        materials.driver_log_wb.save(driver_log_file)
        saved_files.append(driver_log_file)

    # Save a PDF copy of each Excel file that was created. The Excel files are already saved at
    # this point, so a failed PDF export is reported on its own instead of failing the whole run
    created_files = list(saved_files)
    if user_input.should_export_pdf:
        pdf_errors = []
        for number, saved_file in enumerate(saved_files, start=1):
            # The bar counts the PDFs, so the slowest step is still the one that visibly
            # moves. It only moves twice, but a stalled window is what this is here to avoid
            status.update('Making PDFs', done=number - 1, total=len(saved_files))
            try:
                created_files.append(export_to_pdf(saved_file))
            except Exception as pdf_error:
                pdf_errors.append(f'{os.path.basename(saved_file)}\n{str(pdf_error)}')
        if pdf_errors:
            status.close()
            report_error('The Excel files were created, but the PDF export failed:'
                         '\n\n' + '\n\n'.join(pdf_errors))

    status.close()

    if sheet_manager.row_count == 999:
        raise ValueError('There is too much data to fit on the invoice! '
                         '\n\nInvoice has been filled as much data as possible. The rest of the data is not included. '
                         '\n\nDecrease time range to avoid this issue...')

    # Say what was made and where, so the files do not have to be hunted for
    output_folder = os.path.abspath('..')
    file_list = '\n'.join(os.path.basename(path) for path in created_files)
    if report_finished(f'Finished! {len(created_files)} files created for {user_input.project_id}:'
                       f'\n\n{file_list}\n\nSaved in:  {output_folder}'):
        open_output_folder(output_folder)

except Exception as e:
    if status is not None:
        status.close()

    # Point at the last line of this program's own code that the error passed through.
    # Anything unrecognised is still reported, rather than failing in silence
    tb = traceback.extract_tb(sys.exc_info()[2])
    where = next((frame for frame in reversed(tb)
                  if os.path.dirname(os.path.abspath(frame.filename)) == CODE_FOLDER), None)
    if where is not None:
        report_error(f"Error: {str(e)}\n\n\n\nFailed at this spot in the code: "
                     f"\n\tFile: {os.path.basename(where.filename)} \n\tLine: {where.lineno}")
    else:
        report_error(f"Error: {str(e)}")

finally:
    if status is not None:
        status.close()
