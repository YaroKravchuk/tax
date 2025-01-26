from openpyxl.drawing.image import Image
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
import pandas as pd
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles.colors import BLUE
from openpyxl.utils.units import pixels_to_EMU as p2e


class SheetManager:

    # Function to initialize SheetManager class
    def __init__(self, driver_log_wb, driver_log_template, invoice_sheet, taxable):

        # Initialize resources
        self.driver_log_wb = driver_log_wb
        self.driver_log_template = driver_log_template
        self.invoice_sheet = invoice_sheet
        self.taxable = taxable

        # Create global variables
        self.load_row_count = 0
        self.row_count = 0
        self.sheet_count = 0
        self.prev_date = None
        self.prev_truck = None
        self.driver_log_sheet = None
        self.curr_date_row = "8"

    # Creates a new driver log sheet
    def create_new_driver_log_sheet(self, row):
        self.sheet_count = self.sheet_count + 1
        self.load_row_count = 0
        self.prev_date = row["DATE"]
        self.prev_truck = row["TRUCK ID#"]
        self.driver_log_sheet = self.driver_log_wb.copy_worksheet(self.driver_log_template)
        self.driver_log_sheet.title = "Driver Log " + str(self.sheet_count)

        self.driver_log_sheet['A4'] = CellRichText(
            'P.O. Box 3571 Bellevue, WA 98009'
            '\n(425) 569-9910 / ',
            TextBlock(InlineFont(rFont='Calibri', b=True, sz=9), 'prospectllc@pauldimov.com'),
        )
        self.driver_log_sheet['G4'] = CellRichText(
            TextBlock(InlineFont(rFont='Calibri', i=True, sz=7),
                      'And whatever you do, do it heartily, as to the Lord and not to men, knowing that from the Lord '
                      'you will receive the reward of the inheritance; for you serve the Lord Christ.',),
            TextBlock(InlineFont(rFont='Calibri', b=True, sz=7), '\nColossians 3:23-24 NKJV'),
        )

        self.driver_log_sheet['H1'] = row["DATE"]
        self.driver_log_sheet['A7'] = row["CUSTOMER"]
        self.driver_log_sheet['F7'] = row["PROJECT ID"]
        self.driver_log_sheet['N7'] = row["TRUCK ID#"]

        self.add_images_to_driver_log()

    # Function to input a single row of driver log data
    def populate_driver_log_sheet(self, row):
        # First check if previous sheet is same truck, same day
        if row["DATE"] == self.prev_date and row["TRUCK ID#"] == self.prev_truck:
            self.load_row_count = self.load_row_count + 1
        else:
            # If data row contains a new truck or a new day, create a new sheet
            self.create_new_driver_log_sheet(row)

        driver_log_row = str(self.load_row_count + 9)
        self.driver_log_sheet['A' + driver_log_row] = row["HAULING FROM"]
        self.driver_log_sheet['B' + driver_log_row] = row["HAULING TO"]
        self.driver_log_sheet['E' + driver_log_row] = row["PRODUCT"]
        self.driver_log_sheet['H' + driver_log_row] = row["LOAD QTY \n"]
        self.driver_log_sheet['K' + driver_log_row] = row["DUMP FEE RATE"]
        self.driver_log_sheet['J' + driver_log_row] = row["MATERIAL COST"]
        self.driver_log_sheet['L' + driver_log_row] = row["TIME IN"]
        self.driver_log_sheet['N' + driver_log_row] = row["TIME OUT"]
        self.driver_log_sheet['P' + driver_log_row] = row["STAND-BY TIME"]

        if not pd.isnull(row["TIME IN"]):
            self.driver_log_sheet['M20'] = row["TIME IN"]
        if not pd.isnull(row["TIME OUT"]):
            self.driver_log_sheet['M21'] = row["TIME OUT"]
        if not pd.isnull(row["NOTES"]):
            # Store current "comment" text in current_value and add a new line (If the cell has any text in it)
            current_value = str(self.driver_log_sheet['F21'].value) + "\n" \
                if (self.driver_log_sheet['F21'].value is not None
                    and str(self.driver_log_sheet['F21'].value).strip() != "") \
                else ""
            # Format comment: Load 1: "COMMENT"
            self.driver_log_sheet['F21'] = (current_value + "Load " + str(self.load_row_count + 1) + ': "'
                                            + str(row["NOTES"]) + '"')
        if row["HOURS"] > 0:
            self.driver_log_sheet['M22'] = row["HOURS"]

    # Function to input a single row of invoice data
    def populate_invoice_sheet_row(self, row):
        self.invoice_sheet["D4"] = row["CUSTOMER"]
        invoice_row = str(self.row_count + 8)
        self.invoice_sheet["A" + invoice_row] = row["DATE"].date()
        self.invoice_sheet["B" + invoice_row] = row["TRUCK ID#"]
        self.invoice_sheet["C" + invoice_row] = CellRichText(
            TextBlock(InlineFont(rFont='Tahoma', color=BLUE, sz=9, b=True), f"{' ' * 20}{row['SERVICE TYPE']}:"),
            TextBlock(InlineFont(rFont='Tahoma', color=BLUE, sz=9), f" {row['PRODUCT']}")
        )
        self.invoice_sheet["E" + invoice_row] = row["LOAD QTY \n"]
        self.invoice_sheet["F" + invoice_row] = row["RATE PER LOAD"]
        if self.taxable:
            self.invoice_sheet["G" + invoice_row] = "X"
        self.invoice_sheet["H" + invoice_row] = "=E" + invoice_row + "*F" + invoice_row
        self.row_count = self.row_count + 1

        offset = 27
        if pd.notna(row["STAND-BY TIME"]):
            self.populate_invoice_sheet_row_subcategory(
                row,
                f"{' ' * offset}↳ Truck Standby Hours",
                True,
                row["STAND-BY TIME"],
                row["STAND-BY RATE"]
            )
        if pd.notna(row["TIME IN"]):
            self.populate_invoice_sheet_row_subcategory(
                row,
                f"{' ' * offset}↳ Truck Hours Worked",
                True,
                row["HOURS"],
                row["RATE PER HOUR"]
            )
        if pd.notna(row["DUMP FEE RATE"]):
            self.populate_invoice_sheet_row_subcategory(
                row,
                f"{' ' * offset}↳ Dump Fee",
                False,
                row["LOAD QTY \n"],
                row["DUMP FEE RATE"]
            )
        if pd.notna(row["MATERIAL COST"]):
            self.populate_invoice_sheet_row_subcategory(
                row,
                f"{' ' * offset}↳ Material Cost",
                False,
                row["LOAD QTY \n"],
                row["MATERIAL COST"]
            )

    def populate_invoice_sheet_row_subcategory(self, row, description, is_hours_unit, unit, rate):
        invoice_row = str(self.row_count + 8)
        self.invoice_sheet["A" + invoice_row] = row["DATE"].date()
        self.invoice_sheet["B" + invoice_row] = row["TRUCK ID#"]
        self.invoice_sheet["C" + invoice_row] = description.ljust(30)
        if is_hours_unit:
            self.invoice_sheet["E" + invoice_row].number_format = '0.00'
        self.invoice_sheet["E" + invoice_row] = unit
        self.invoice_sheet["F" + invoice_row] = rate
        self.invoice_sheet["H" + invoice_row] = "=E" + invoice_row + "*F" + invoice_row

        self.row_count = self.row_count + 1

    # Function to configure the Logo image and the Signature image for the driver log
    def add_images_to_driver_log(self):
        # Logo image setup
        logo_image = Image('logo.png')
        logo_image.width = 90
        logo_image.height = 65
        # Create offset. O1 is column 14, row 0
        logo_size = XDRPositiveSize2D(p2e(logo_image.width), p2e(logo_image.height))
        logo_marker = AnchorMarker(col=14, colOff=p2e(10), row=0, rowOff=p2e(10))
        logo_image.anchor = OneCellAnchor(_from=logo_marker, ext=logo_size)

        # Signature image setup
        signature_image = Image('sig.png')
        signature_image.width = 104
        signature_image.height = 40
        # Create offset. M21 is column 12, row 25
        sig_size = XDRPositiveSize2D(p2e(signature_image.width), p2e(signature_image.height))
        sig_marker = AnchorMarker(col=12, colOff=p2e(30), row=24, rowOff=p2e(2))
        signature_image.anchor = OneCellAnchor(_from=sig_marker, ext=sig_size)
        self.driver_log_sheet.add_image(logo_image)
        self.driver_log_sheet.add_image(signature_image)

    def merge_date_cells(self):
        current_val = None
        start_row = 8

        for row in range(8, self.row_count + 8):
            cell_value = self.invoice_sheet[f"A{row}"].value

            if cell_value != current_val:
                if current_val and start_row < row - 1:
                    self.invoice_sheet.merge_cells(f"A{start_row}:A{row-1}")
                current_val = cell_value
                start_row = row

            if current_val and row == self.row_count + 7 and start_row < self.row_count + 7:
                self.invoice_sheet.merge_cells(f"A{start_row}:A{self.row_count + 7}")

    def merge_truck_cells(self):
        start_row = 8

        for row in range(8, self.row_count + 8):
            cell_value = self.invoice_sheet[f"C{row}"].value

            if '↳' not in str(cell_value):
                if start_row < row - 1:
                    self.invoice_sheet.merge_cells(f"B{start_row}:B{row-1}")
                start_row = row
            elif row == self.row_count + 7:
                self.invoice_sheet.merge_cells(f"B{start_row}:B{self.row_count + 7}")
