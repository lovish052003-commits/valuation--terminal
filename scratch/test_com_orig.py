import os
import win32com.client
import pythoncom

path = os.path.abspath(r'C:\Users\LENOVO\Downloads\test\SUNPHARMA_Valuation_Model (1).xlsx')
print(f"Testing COM open on original: {path}")

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(path, UpdateLinks=0, ReadOnly=False)
    print(f"Successfully opened original: {wb.Name}")
    wb.Close()
except Exception as e:
    print(f"Original failed with: {e}")
finally:
    excel.Quit()
    pythoncom.CoUninitialize()
