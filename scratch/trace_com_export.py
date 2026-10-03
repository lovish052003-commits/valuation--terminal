import sys
import os
import shutil
import pythoncom
import win32com.client
sys.path.insert(0, '.')
import screener_client
import valuation_engine
import excel_exporter

print("[1] Fetching ADANIENT...", flush=True)
sd = screener_client.fetch_company_data("ADANIENT")
print("[2] Calculating valuation...", flush=True)
val = valuation_engine.calculate_valuation(sd)

dest_path = os.path.abspath("exports/TEST_ADANI_TRACE.xlsx")
shutil.copyfile(excel_exporter.TEMPLATE_PATH, dest_path)

print("[3] Launching Excel COM...", flush=True)
pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.ScreenUpdating = False

print("[4] Opening workbook...", flush=True)
wb = excel.Workbooks.Open(dest_path, UpdateLinks=0, ReadOnly=False)
try:
    excel.Calculation = -4135
except Exception:
    pass

sheet_names = [s.Name for s in wb.Sheets]
print(f"[5] Sheets in workbook: {len(sheet_names)}", flush=True)

print("[6] Populating Data Sheet...", flush=True)
excel_exporter.populate_data_sheet(wb.Sheets('Data Sheet'), sd)

print("[7] Populating Raw FS...", flush=True)
excel_exporter.populate_raw_fs_sheet(wb.Sheets('Raw FS'), sd, val)

print("[8] Populating Cash Flow Statement...", flush=True)
excel_exporter.populate_cash_flow_statement_sheet(wb.Sheets('Cash Flow Statement'), sd)

print("[9] Populating Raw Data prices...", flush=True)
ws_beta = wb.Sheets('Beta-Regression') if 'Beta-Regression' in sheet_names else None
excel_exporter.populate_raw_data_prices(wb.Sheets('Raw Data'), ws_beta, sd)

print("[10] Fixing Forecasting...", flush=True)
excel_exporter.fix_forecasting_sheet(wb.Sheets('Forecasting'))

print("[11] Updating DCF...", flush=True)
ws_dcf = wb.Sheets('DCF')
cmp_price = screener_client.clean_num(sd.get('current_price', 0))
if cmp_price > 0:
    ws_dcf.Range('D44').Value = cmp_price
ws_dcf.Range('D37').Formula = "='Data Sheet'!K69"
ws_dcf.Range('D38').Formula = "='Data Sheet'!K59"
ws_dcf.Range('D40').Formula = "='Data Sheet'!K70/10000000"

print("[12] Updating Comp_Valuation...", flush=True)
excel_exporter.update_comp_valuation_sheet(wb.Sheets('Comp_Valuation'), sd, val)

print("[13] Updating WACC & Raw Data...", flush=True)
excel_exporter.update_wacc_raw_data(wb.Sheets('Raw Data'), wb.Sheets('WACC'), sd, val)

print("[14] Populating DuPont & Altman...", flush=True)
excel_exporter.populate_dupont_altman_sheets(wb, sd)

print("[15] Populating AI Valuation Summary...", flush=True)
if 'AI Valuation Summary' in sheet_names:
    ws_sum = wb.Sheets('AI Valuation Summary')
else:
    ws_sum = wb.Sheets.Add(Before=wb.Sheets(1))
    ws_sum.Name = 'AI Valuation Summary'
excel_exporter.populate_ai_summary_sheet(ws_sum, sd, val)

print("[16] Recalculating...", flush=True)
try:
    excel.Calculation = -4105
except Exception:
    pass
excel.CalculateFull()

print("[17] Saving...", flush=True)
wb.Save()
wb.Close(SaveChanges=True)
excel.Quit()
pythoncom.CoUninitialize()

print(f"[18] SUCCESS! File saved to: {dest_path}", flush=True)
