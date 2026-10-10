import sys, os, time
sys.path.insert(0, os.path.abspath('.'))
import pythoncom
import win32com.client
import json

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.ScreenUpdating = False

test_copy = os.path.abspath('exports/test_altman.xlsx')
import shutil
shutil.copyfile('master_model_template.xlsx', test_copy)

wb = excel.Workbooks.Open(test_copy)
sheet_names = [s.Name for s in wb.Sheets]
print("Opened workbook, sheets count:", len(sheet_names))

from screener_client import fetch_company_data
data = fetch_company_data('NESTLEIND')
print("Fetched NESTLEIND data. Calling populate_dupont_altman_sheets...")
import excel_exporter
excel_exporter.populate_dupont_altman_sheets(wb, data)
print("Finished populate_dupont_altman_sheets successfully!")
wb.Close(SaveChanges=False)
excel.Quit()
pythoncom.CoUninitialize()
print("All clean!")
