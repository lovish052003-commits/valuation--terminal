import os
import win32com.client
import pythoncom

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

file_path = os.path.abspath('exports/TEST_RBLBANK_COM.xlsx')
print(f"Verifying normal open of {file_path}...")
try:
    # CorruptLoad=0 (xlNormalLoad)
    wb = excel.Workbooks.Open(file_path, UpdateLinks=0, ReadOnly=True, CorruptLoad=0)
    print(f"SUCCESS: Workbook opened cleanly with 0 errors! Sheet count: {wb.Sheets.Count}")
    for s in wb.Sheets:
        print(f" - Sheet: {s.Name}")
    wb.Close(SaveChanges=False)
except Exception as e:
    print(f"FAILED: {e}")
finally:
    excel.Quit()
    pythoncom.CoUninitialize()
