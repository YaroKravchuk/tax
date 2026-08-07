import os
import platform
import shutil
import subprocess
import tempfile

# Excel's xlTypePDF constant, used to tell ExportAsFixedFormat to write a PDF
XL_TYPE_PDF = 0

# Where LibreOffice normally installs itself, checked when it is not on the PATH
LIBREOFFICE_PATHS = [
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    r'C:\Program Files\LibreOffice\program\soffice.exe',
    r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
]

INSTALL_PYWIN32_MESSAGE = (
    'PDF export on Windows uses Excel, which needs the pywin32 package.'
    '\n\nInstall it by running: pip install pywin32'
)

INSTALL_LIBREOFFICE_MESSAGE = (
    'PDF export on this computer uses LibreOffice, which could not be found.'
    '\n\nInstall it from https://www.libreoffice.org/download'
    '\nor, on a Mac, by running: brew install --cask libreoffice'
)


# Function to convert a saved workbook into a PDF sitting right next to it
def export_to_pdf(xlsx_path):
    xlsx_path = os.path.abspath(xlsx_path)
    pdf_path = os.path.splitext(xlsx_path)[0] + '.pdf'

    # Both Excel and LibreOffice can stall on a PDF that is already there
    if os.path.exists(pdf_path):
        os.remove(pdf_path)

    if platform.system() == 'Windows':
        export_with_excel(xlsx_path, pdf_path)
    else:
        export_with_libreoffice(xlsx_path, pdf_path)

    if not os.path.exists(pdf_path):
        raise RuntimeError(f'The PDF was not created: {pdf_path}')

    return pdf_path


# Function to export a workbook using Excel itself, which honors the print area,
# orientation and scaling saved in the templates exactly as a manual export does
def export_with_excel(xlsx_path, pdf_path):
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError(INSTALL_PYWIN32_MESSAGE)

    # DispatchEx starts a separate Excel so an Excel the user already has open is left alone
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    workbook = None
    try:
        workbook = excel.Workbooks.Open(xlsx_path, ReadOnly=True)
        workbook.ExportAsFixedFormat(XL_TYPE_PDF, pdf_path)
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()


# Function to export a workbook using LibreOffice, used where Excel cannot be scripted
def export_with_libreoffice(xlsx_path, pdf_path):
    soffice = find_libreoffice()
    if soffice is None:
        raise RuntimeError(INSTALL_LIBREOFFICE_MESSAGE)

    # Use a private profile so the export still runs when LibreOffice is already open
    profile_dir = os.path.join(tempfile.gettempdir(), 'prospect_llc_soffice_profile')
    result = subprocess.run(
        [soffice, f'-env:UserInstallation=file://{profile_dir}', '--headless', '--convert-to', 'pdf',
         '--outdir', os.path.dirname(pdf_path), xlsx_path],
        capture_output=True, text=True, timeout=300
    )

    if result.returncode != 0:
        raise RuntimeError(f'LibreOffice could not convert the file: {result.stderr.strip()}')


# Function to locate the LibreOffice command line program
def find_libreoffice():
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if soffice is not None:
        return soffice

    for path in LIBREOFFICE_PATHS:
        if os.path.exists(path):
            return path

    return None
