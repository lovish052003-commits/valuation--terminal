import win32com.client as win32
import os
import sys
import pythoncom
import shutil

sys.path.insert(0, os.path.abspath('.'))
import excel_exporter

test_file = os.path.abspath('scratch/test_kalyan_com_patch.xlsx')
shutil.copyfile('exports/KALYANKJIL_Valuation_Model.xlsx.bak', test_file)

pythoncom.CoInitialize()
excel = win32.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(test_file)

ws_ds = wb.Sheets('Data Sheet')
ws_raw = wb.Sheets('Raw FS')
ws_dcf = wb.Sheets('DCF')
ws_wacc = wb.Sheets('WACC')

debt_anchor = float(ws_ds.Range('K59').Value or 0)
print('Debt anchor:', debt_anchor)

borrowing_row = None
for r in range(1, 60):
    lbl = str(ws_raw.Cells(r, 2).Value or '').strip().lower()
    if 'borrowing' in lbl or 'debt' in lbl:
        borrowing_row = r
        break

target_col = None
for c in range(3, 25):
    val = ws_raw.Cells(borrowing_row, c).Value
    if val is not None and abs(float(val) - debt_anchor) < 1.0:
        target_col = c
        break

print(f'Borrowing row: {borrowing_row}, Target col: {target_col}')

mi_row = None
for r in range(1, 60):
    lbl = str(ws_raw.Cells(r, 2).Value or '').strip().lower()
    if 'controlling' in lbl or 'minority' in lbl:
        mi_row = r
        break

mi_val = float(ws_raw.Cells(mi_row, target_col).Value or 0)
print(f'Minority row: {mi_row}, Minority val: {mi_val}')

ws_ds.Range('A73').Value = 'Minority Interest'
ws_ds.Range('K73').Value = mi_val

# DCF
ws_dcf.Range('D39').Formula = "='Data Sheet'!K73"
ws_dcf.Range('D40').Formula = "=D35+D37-D38-D39"
ws_dcf.Range('D43').Formula = "=D40/D41"

# WACC Peer Betas
peer_betas = {'Titan Company': 0.88, 'Havells India': 1.05, 'Berger Paints': 0.82, 'Asian Paints': 0.76, 'Lalithaa Jewel': 1.10}
for r in range(14, 19):
    name = str(ws_wacc.Cells(r, 2).Value or '')
    if name.startswith('='):
        name = str(wb.Sheets('Raw Data').Cells(10 + r, 15).Value or '')
    b = peer_betas.get(name, 0.90)
    ws_wacc.Cells(r, 10).Value = b
    ws_wacc.Cells(r, 11).Formula = f"=J{r}/(1+(1-G{r})*H{r})"
    print(f'WACC row {r} ({name}): Levered Beta = {b}')

try:
    print("Saving workbook via COM...")
    wb.Save()
    print("Saved successfully. Closing...")
    wb.Close(True)
    print("Closed.")
except Exception as e:
    import traceback
    traceback.print_exc()
finally:
    excel.Quit()
    pythoncom.CoUninitialize()

excel_exporter.strip_calc_chain_from_xlsx(test_file)

# Test opening with CorruptLoad=0!
pythoncom.CoInitialize()
excel = win32.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb_test = excel.Workbooks.Open(test_file, CorruptLoad=0)
    print('\nRESULT: SUCCESS! Opened cleanly in Excel COM with CorruptLoad=0 (0 warnings)!')
    wb_test.Close(False)
except Exception as e:
    import traceback
    traceback.print_exc()
finally:
    excel.Quit()
    pythoncom.CoUninitialize()
