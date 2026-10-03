import win32com.client as win32, os, pythoncom, shutil

test_path = os.path.abspath('scratch/test_com_simple.xlsx')
shutil.copyfile('ITC Model.xlsx', test_path)

pythoncom.CoInitialize()
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(test_path)
ws_dcf = wb.Sheets('DCF')
eq_row = 39

print('Before insert, D39 formula:', ws_dcf.Range('D39').Formula)
ws_dcf.Rows(eq_row).Insert()
ws_dcf.Range('B39').Value = 'Less: Minority Interest'
ws_dcf.Range('D39').Formula = "='Data Sheet'!K73"
ws_dcf.Range('D40').Formula = "=D35+D37-D38-D39"
ws_dcf.Range('D43').Formula = "=D40/D41"

print('After insert, D39 formula:', ws_dcf.Range('D39').Formula)
print('After insert, D40 formula:', ws_dcf.Range('D40').Formula)
print('After insert, D43 formula:', ws_dcf.Range('D43').Formula)

wb.Save()
wb.Close(True)
excel.Quit()
pythoncom.CoUninitialize()

# Test opening with CorruptLoad=0!
pythoncom.CoInitialize()
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb_test = excel.Workbooks.Open(test_path, CorruptLoad=0)
    print('SUCCESS: Workbook opens cleanly with ZERO recovery warnings!')
    wb_test.Close(False)
except Exception as e:
    print('FAILED:', e)

excel.Quit()
pythoncom.CoUninitialize()
