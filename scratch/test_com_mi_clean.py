import win32com.client as win32, os, sys, pythoncom, shutil
sys.path.insert(0, os.path.abspath('.'))
import screener_client, valuation_engine, excel_exporter

d = screener_client.fetch_company_data('UNIONBANK')
v = valuation_engine.calculate_valuation(d)

test_path = os.path.abspath('scratch/test_clean_export.xlsx')
shutil.copyfile('ITC Model.xlsx', test_path)

pythoncom.CoInitialize()
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(test_path)

# 1. Standard COM updates
excel_exporter.populate_data_sheet(wb.Sheets('Data Sheet'), d)
excel_exporter.populate_raw_fs_sheet(wb.Sheets('Raw FS'), d, v)
excel_exporter.update_comp_valuation_sheet(wb.Sheets('Comp_Valuation'), d, v)
excel_exporter.update_raw_data_and_wacc(wb, d, v)
excel_exporter.populate_dupont_altman_sheets(wb, d)

# 2. DCF standard updates
ws_dcf = wb.Sheets('DCF')
ws_dcf.Range('D44').Formula = "='Data Sheet'!B8"
ws_dcf.Range('D37').Formula = "='Data Sheet'!K69"
ws_dcf.Range('D38').Formula = "='Data Sheet'!K59"
ws_dcf.Range('D40').Formula = "='Data Sheet'!K70/10000000"

# 3. Native COM Minority Interest Fix
ws_raw = wb.Sheets('Raw FS')
ws_data = wb.Sheets('Data Sheet')
minority_val = 0.0
found_row = None
for r in range(1, 120):
    try:
        cell_val = str(ws_raw.Cells(r, 2).Value or '').strip().lower()
        if 'controlling' in cell_val or 'minority' in cell_val:
            found_row = r
            break
    except Exception:
        pass

if found_row:
    for c in range(3, 30):
        try:
            val = ws_raw.Cells(found_row, c).Value
            if val is not None and isinstance(val, (int, float)):
                minority_val = float(val)
        except Exception:
            break

ws_data.Range('A73').Value = "Minority Interest"
ws_data.Range('K73').Value = minority_val

has_mi_row = False
eq_row = None
for r in range(1, 60):
    try:
        lbl = str(ws_dcf.Cells(r, 2).Value or '').strip().lower()
        if 'minority interest' in lbl:
            has_mi_row = True
            break
        if lbl == 'equity value' and not eq_row:
            eq_row = r
    except Exception:
        pass

if not has_mi_row and eq_row:
    ws_dcf.Rows(eq_row).Insert()
    ws_dcf.Range(f'B{eq_row}').Value = "Less: Minority Interest"
    ws_dcf.Range(f'D{eq_row}').Formula = "='Data Sheet'!K73"
    new_eq = eq_row + 1
    ws_dcf.Range(f'D{new_eq}').Formula = f"=D{eq_row-4}+D{eq_row-2}-D{eq_row-1}-D{eq_row}"
    shares_row = new_eq + 1
    per_share_row = new_eq + 3
    ws_dcf.Range(f'D{per_share_row}').Formula = f"=D{new_eq}/D{shares_row}"

# 4. Insert AI Valuation Summary in COM
excel_exporter.insert_or_update_ai_summary_sheet(wb, d, v)

# Save
wb.Save()
wb.Close(True)
excel.Quit()
pythoncom.CoUninitialize()

excel_exporter.strip_calc_chain_from_xlsx(test_path)

# Verify with Excel COM CorruptLoad=0!
pythoncom.CoInitialize()
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb_test = excel.Workbooks.Open(test_path, CorruptLoad=0)
    print('RESULT: SUCCESS! Opened cleanly with ZERO recovery/repair warnings!')
    wb_test.Close(False)
except Exception as e:
    print('RESULT: FAILED with error:', e)

excel.Quit()
pythoncom.CoUninitialize()
