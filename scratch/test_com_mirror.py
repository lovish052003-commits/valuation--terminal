import os
import win32com.client
import pythoncom

path = os.path.abspath(r'SUNPHARMA_Valuation_Model_Corrected.xlsx')
print(f"Testing COM open for: {path}")

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(path, UpdateLinks=0, ReadOnly=False)
    print(f"Successfully opened: {wb.Name}")
    excel.CalculateFull()
    print("CalculateFull succeeded!")
    wb.Save()
    print("Save succeeded!")
    wb.Close()
except Exception as e:
    print(f"Failed with: {e}")
finally:
    excel.Quit()
    pythoncom.CoUninitialize()
