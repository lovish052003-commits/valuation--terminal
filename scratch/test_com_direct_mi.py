import win32com.client as win32, os, pythoncom, shutil

test_path = os.path.abspath('scratch/test_com_mi_direct.xlsx')
shutil.copyfile('ITC Model.xlsx', test_path)

pythoncom.CoInitialize()
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(test_path)
ws_raw = wb.Sheets('Raw FS')
ws_dcf = wb.Sheets('DCF')
ws_data = wb.Sheets('Data Sheet')

# 1. Minority Interest lookup
minority_val = 0.0
for r in range(1, 100):
    val = str(ws_raw.Cells(r, 2).Value or '').lower()
    if 'controlling' in val or 'minority' in val:
        for c in range(3, 30):
            cv = ws_raw.Cells(r, c).Value
            if cv is not None and isinstance(cv, (int, float)):
                minority_val = float(cv)
        break

print('Found minority_val:', minority_val)
ws_data.Range('A73').Value = 'Minority Interest'
ws_data.Range('K73').Value = minority_val

# 2. DCF row insertion
eq_row = None
has_mi = False
for r in range(1, 60):
    lbl = str(ws_dcf.Cells(r, 2).Value or '').strip().lower()
    if 'minority interest' in lbl:
        has_mi = True
        break
    if lbl == 'equity value' and not eq_row:
        eq_row = r

print('eq_row:', eq_row, 'has_mi:', has_mi)
if not has_mi and eq_row:
    ws_dcf.Rows(eq_row).Insert()
    ws_dcf.Range(f'B{eq_row}').Value = 'Less: Minority Interest'
    ws_dcf.Range(f'D{eq_row}').Formula = "='Data Sheet'!K73"
    new_eq = eq_row + 1
    
    # Copy formatting from row above
    ws_dcf.Range(f'B{eq_row-1}').Copy()
    ws_dcf.Range(f'B{eq_row}').PasteSpecial(Paste=-4122) # xlPasteFormats
    ws_dcf.Range(f'D{eq_row-1}').Copy()
    ws_dcf.Range(f'D{eq_row}').PasteSpecial(Paste=-4122) # xlPasteFormats
    excel.CutCopyMode = False

    ws_dcf.Range(f'B{eq_row}').Value = 'Less: Minority Interest'
    ws_dcf.Range(f'D{eq_row}').Formula = "='Data Sheet'!K73"
    ws_dcf.Range(f'D{new_eq}').Formula = f'=D{eq_row-4}+D{eq_row-2}-D{eq_row-1}-D{eq_row}'
    ws_dcf.Range(f'D{new_eq+3}').Formula = f'=D{new_eq}/D{new_eq+1}'
    print(f'Formula at new_eq ({new_eq}):', ws_dcf.Range(f'D{new_eq}').Formula)
    print(f'Formula at per_share ({new_eq+3}):', ws_dcf.Range(f'D{new_eq+3}').Formula)

wb.Save()
wb.Close(True)
excel.Quit()
pythoncom.CoUninitialize()

# Verify with strict CorruptLoad=0
pythoncom.CoInitialize()
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb_test = excel.Workbooks.Open(test_path, CorruptLoad=0)
    print('VERIFICATION: Succeeded with ZERO recovery warnings!')
    wb_test.Close(False)
except Exception as e:
    print('VERIFICATION FAILED:', e)

excel.Quit()
pythoncom.CoUninitialize()
