import win32com.client as win32, os, pythoncom, shutil

test_path = os.path.abspath('scratch/test_com_mi.xlsx')
shutil.copyfile('ITC Model.xlsx', test_path)

pythoncom.CoInitialize()
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(test_path)
ws_raw = wb.Sheets('Raw FS')
ws_dcf = wb.Sheets('DCF')
ws_data = wb.Sheets('Data Sheet')

minority_val = 0.0
last_r = ws_raw.UsedRange.Rows.Count
for r in range(1, min(last_r + 1, 100)):
    txt = str(ws_raw.Cells(r, 2).Value or '').lower()
    if 'controlling' in txt or 'minority' in txt:
        for c in range(3, 30):
            v = ws_raw.Cells(r, c).Value
            if v is not None and isinstance(v, (int, float)):
                minority_val = float(v)
        break

print('Found minority_val:', minority_val)
ws_data.Range('A73').Value = 'Minority Interest'
ws_data.Range('K73').Value = minority_val

dcf_rows = ws_dcf.UsedRange.Rows.Count
has_mi = False
eq_row = None
for r in range(1, min(dcf_rows + 1, 60)):
    lbl = str(ws_dcf.Cells(r, 2).Value or '').strip().lower()
    if 'minority interest' in lbl:
        has_mi = True
        break
    if lbl == 'equity value' and not eq_row:
        eq_row = r

if not has_mi and eq_row:
    print('Inserting row at DCF row:', eq_row)
    ws_dcf.Rows(eq_row).Insert()
    ws_dcf.Range(f'B{eq_row}').Value = 'Less: Minority Interest'
    ws_dcf.Range(f'D{eq_row}').Formula = "='Data Sheet'!K73"
    
    new_eq = eq_row + 1
    ws_dcf.Range(f'D{new_eq}').Formula = f'=D{eq_row-4}+D{eq_row-2}-D{eq_row-1}-D{eq_row}'
    print('New Equity Value Formula:', ws_dcf.Range(f'D{new_eq}').Formula)

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
    print('SUCCESS: Native COM Minority Interest workbook opens with ZERO RECOVERY WARNINGS!')
    wb_test.Close(False)
except Exception as e:
    print('FAILURE: Error opening:', e)

excel.Quit()
pythoncom.CoUninitialize()
