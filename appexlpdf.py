
import os
import win32com.client as win32

# File location
excel_file = r"C:\Users\LENOVO\Downloads\Excel.xlsx"
pdf_file = r"C:\Users\LENOVO\Downloads\Exceldata.pdf"

excel = None
workbook = None

try:
    # Initialize Excel
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    # Open workbook
    workbook = excel.Workbooks.Open(
        os.path.abspath(excel_file),
        ReadOnly=True
    )

    # Select worksheet
    worksheet = workbook.Worksheets(1)

    # Export to PDF
    worksheet.ExportAsFixedFormat(
        0,
        os.path.abspath(pdf_file)
    )

    print("Export PDF successfully!")
    print(f"PDF location: {pdf_file}")

except Exception as e:
    print(f"Export failed: {e}")

finally:
    if workbook is not None:
        workbook.Close(SaveChanges=False)

    if excel is not None:
        excel.Quit()
