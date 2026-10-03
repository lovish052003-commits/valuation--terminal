import os
import win32com.client
import pythoncom

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('ITC Model.xlsx'), UpdateLinks=0, ReadOnly=False)
excel.Calculation = -4135  # xlCalculationManual
ws_data = wb.Sheets('Data Sheet')
ws_data.Range('B8').Value = 403.0
ws_data.Range('B9').Value = 62485.0
ws_data.Range('B6').Formula = '=IF(B9>0, B9/B8, 0)'
print('Before calculate, B6 value:', ws_data.Range('B6').Value)
ws_data.Calculate()
print('After ws_data.Calculate(), B6 value:', ws_data.Range('B6').Value)
ws_raw = wb.Sheets('Raw FS')
ws_raw.Cells(56, 14).Formula = "='Data Sheet'!B6"
print('Before ws_raw.Calculate(), Raw FS N56:', ws_raw.Cells(56, 14).Value)
ws_raw.Calculate()
print('After ws_raw.Calculate(), Raw FS N56:', ws_raw.Cells(56, 14).Value)
wb.Close(SaveChanges=False)
excel.Quit()
pythoncom.CoUninitialize()
