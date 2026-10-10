import os, sys, time, shutil
sys.path.insert(0, os.path.abspath('.'))
import screener_client, valuation_engine
import win32com.client, pythoncom
import excel_exporter

print("Fetching data for ITC...", flush=True)
sd = screener_client.fetch_company_data('ITC')
val = valuation_engine.calculate_valuation(sd)
dest_path = os.path.abspath('exports/TEST_STEP_ITC.xlsx')
shutil.copyfile(excel_exporter.TEMPLATE_PATH, dest_path)

pythoncom.CoInitialize()
print("Dispatching Excel...", flush=True)
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.ScreenUpdating = False

print("Opening workbook...", flush=True)
wb = excel.Workbooks.Open(dest_path, UpdateLinks=0, ReadOnly=False)
sheet_names = [s.Name for s in wb.Sheets]
print("Workbook opened, sheets:", len(sheet_names), flush=True)

t0 = time.time()
if 'Data Sheet' in sheet_names:
    print("1. Populating Data Sheet...", flush=True)
    excel_exporter.populate_data_sheet(wb.Sheets('Data Sheet'), sd)
    print(f"Data Sheet done in {time.time()-t0:.2f}s", flush=True)

t0 = time.time()
if 'Raw FS' in sheet_names:
    print("2. Populating Raw FS...", flush=True)
    excel_exporter.populate_raw_fs_sheet(wb.Sheets('Raw FS'), sd, val)
    print(f"Raw FS done in {time.time()-t0:.2f}s", flush=True)

t0 = time.time()
if 'Cash Flow Statement' in sheet_names:
    print("2b. Populating Cash Flow Statement...", flush=True)
    excel_exporter.populate_cash_flow_statement_sheet(wb.Sheets('Cash Flow Statement'), sd)
    print(f"Cash Flow Statement done in {time.time()-t0:.2f}s", flush=True)

t0 = time.time()
if 'Raw Data' in sheet_names:
    print("2c. Populating Raw Data...", flush=True)
    ws_beta = wb.Sheets('Beta-Regression') if 'Beta-Regression' in sheet_names else None
    excel_exporter.populate_raw_data_prices(wb.Sheets('Raw Data'), ws_beta, sd)
    print(f"Raw Data done in {time.time()-t0:.2f}s", flush=True)

t0 = time.time()
if 'Comp_Valuation' in sheet_names:
    print("5. Updating Comp_Valuation...", flush=True)
    excel_exporter.update_comp_valuation_sheet(wb.Sheets('Comp_Valuation'), sd, val)
    print(f"Comp_Valuation done in {time.time()-t0:.2f}s", flush=True)

t0 = time.time()
if 'Raw Data' in sheet_names and 'WACC' in sheet_names:
    print("6. Updating WACC Raw Data...", flush=True)
    excel_exporter.update_wacc_raw_data(wb.Sheets('Raw Data'), wb.Sheets('WACC'), sd, val)
    print(f"WACC Raw Data done in {time.time()-t0:.2f}s", flush=True)

t0 = time.time()
print("8. Updating DuPont & Altman...", flush=True)
excel_exporter.populate_dupont_altman_sheets(wb, sd)
print(f"DuPont & Altman done in {time.time()-t0:.2f}s", flush=True)

t0 = time.time()
print("Populating AI Summary...", flush=True)
ws_sum = wb.Sheets('AI Valuation Summary') if 'AI Valuation Summary' in sheet_names else wb.Sheets.Add(Before=wb.Sheets(1))
ws_sum.Name = 'AI Valuation Summary'
excel_exporter.populate_ai_summary_sheet(ws_sum, sd, val)
print(f"AI Summary done in {time.time()-t0:.2f}s", flush=True)

t0 = time.time()
print("Calculating Full...", flush=True)
excel.Calculation = -4105
excel.CalculateFull()
print(f"CalculateFull done in {time.time()-t0:.2f}s", flush=True)

ws_sum_eval = wb.Sheets('AI Valuation Summary')
ws_dcf_eval = wb.Sheets('DCF')
print("AI Summary B5 (IV):", ws_sum_eval.Range('B5').Value)
print("AI Summary C5 (Upside):", ws_sum_eval.Range('C5').Value)
print("AI Summary D5 (MOS):", ws_sum_eval.Range('D5').Value)
print("AI Summary E5 (Verdict):", ws_sum_eval.Range('E5').Value)
print("AI Summary F5 (WACC):", ws_sum_eval.Range('F5').Value)
print("DCF D35 (EV):", ws_dcf_eval.Range('D35').Value)
print("DCF D37 (Cash):", ws_dcf_eval.Range('D37').Value)
print("DCF D38 (Debt):", ws_dcf_eval.Range('D38').Value)
print("DCF D39 (Minority):", ws_dcf_eval.Range('D39').Value)
print("DCF D40 (Equity):", ws_dcf_eval.Range('D40').Value)
print("DCF D41 (Shares):", ws_dcf_eval.Range('D41').Value)
print("DCF D43 (IV):", ws_dcf_eval.Range('D43').Value)
print("DCF D18 (Growth):", ws_dcf_eval.Range('D18').Value)
print("DCF D19 (TG):", ws_dcf_eval.Range('D19').Value)
print("DCF D20 (WACC):", ws_dcf_eval.Range('D20').Value)

wb.Save()
wb.Close(SaveChanges=True)
excel.Quit()
print("All complete successfully!", flush=True)
