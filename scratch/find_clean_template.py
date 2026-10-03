import win32com.client, pythoncom, os

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

candidates = [
    'ITC Model.xlsx',
    'exports/ITC_Valuation_Model.xlsx',
    'exports/BRITANNIA_Valuation_Model.xlsx',
    'exports/NESTLEIND_Valuation_Model.xlsx',
    'exports/RELIANCE_Valuation_Model.xlsx',
    'exports/test_com_beta.xlsx',
    'exports/TITAN_Valuation_Model.xlsx',
]

for c in candidates:
    if os.path.exists(c):
        try:
            wb = excel.Workbooks.Open(os.path.abspath(c))
            print(f"CLEAN: {c} opened successfully! Sheet count: {wb.Sheets.Count}")
            wb.Close(SaveChanges=False)
        except Exception as e:
            print(f"FAILED: {c} -> {e}")

excel.Quit()
pythoncom.CoUninitialize()
