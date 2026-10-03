import win32com.client, pythoncom, os, shutil

dest_path = os.path.abspath('exports/test_com_debug.xlsx')
shutil.copyfile('ITC Model.xlsx', dest_path)
pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(dest_path)
ws_beta = wb.Sheets('Beta-Regression')
max_pts = 247
end_beta_r = 10 + max_pts - 1
bcd_formulas = []
for i in range(max_pts):
    raw_r = 6 + i
    beta_r = 10 + i
    b_f = f"='Raw Data'!G{raw_r}"
    c_f = f"='Raw Data'!H{raw_r}"
    d_f = "-" if i == 0 else f"=C{beta_r-1}/C{beta_r}-1"
    bcd_formulas.append([b_f, c_f, d_f])

print('Attempting formula assignment...')
try:
    ws_beta.Range(ws_beta.Cells(10, 2), ws_beta.Cells(end_beta_r, 4)).Formula = bcd_formulas
    print('Formula assignment succeeded!')
except Exception as e:
    print('Formula assignment FAILED:', e)

print('C10 value:', ws_beta.Range('C10').Value)
print('C10 formula:', ws_beta.Range('C10').Formula)
print('D10 value:', ws_beta.Range('D10').Value)
print('D10 formula:', ws_beta.Range('D10').Formula)
print('C11 value:', ws_beta.Range('C11').Value)
print('C11 formula:', ws_beta.Range('C11').Formula)
wb.Close(SaveChanges=False)
excel.Quit()
pythoncom.CoUninitialize()
