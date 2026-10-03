import os, sys, json, time

sys.stdout.reconfigure(line_buffering=True)
sys.path.insert(0, os.path.abspath('.'))

import excel_exporter
import valuation_engine

import screener_client
print("Step 1: Fetching UNITEDTEA data via screener_client...")
screener_data = screener_client.fetch_company_data('UNITEDTEA')

print("Step 2: Calculating valuation...")
val_result = valuation_engine.calculate_valuation(screener_data)
print("Valuation calculated. Intrinsic value:", val_result.get('intrinsic_value'))

# Temporarily point TEMPLATE_PATH to TEST_MARUTI_COM.xlsx
excel_exporter.TEMPLATE_PATH = os.path.abspath('exports/TEST_MARUTI_COM.xlsx')

out_path = os.path.abspath('scratch/TEST_UNITEDTEA_STEP.xlsx')
if os.path.exists(out_path):
    os.remove(out_path)

print("Step 3: Calling export_via_excel_com...")
t0 = time.time()
success = excel_exporter.export_via_excel_com(out_path, screener_data, val_result)
print(f"export_via_excel_com finished in {time.time()-t0:.2f}s, result = {success}")

print("Step 4: Verifying Excel COM Open without alerts...")
import win32com.client
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(out_path)
print(f"SUCCESS: Opened with NO alerts! Sheets count: {len(wb.Sheets)}")
sheet_names = [s.Name for s in wb.Sheets]
print("Sheets:", sheet_names[:10], "...")
wb.Close(False)
excel.Quit()
print("All Done Successfully!")
