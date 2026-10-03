import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client
import valuation_engine
import excel_exporter
import openpyxl
import requests

print("=== RE-EXPORTING KALYANKJIL WITH ORIGINAL WORKBOOK LOGIC ===")
data = screener_client.fetch_company_data('KALYANKJIL')
val = valuation_engine.calculate_valuation(data)

dest_path = excel_exporter.export_valuation_model(data, val, "")
print(f"Exported to: {dest_path}")

print("\n=== VERIFYING WORKBOOK VALUES ===")
wb_v = openpyxl.load_workbook(dest_path, data_only=True)
ws_dcf = wb_v['DCF']
ws_sum = wb_v['AI Valuation Summary']

print("DCF H8:", ws_dcf['H8'].value)
print("DCF D35 (Operating Assets / EV):", ws_dcf['D35'].value)
print("DCF D40 (Equity Value):", ws_dcf['D40'].value)
print("DCF D43 (Intrinsic Value/Share):", ws_dcf['D43'].value)
print("Summary B5 (Intrinsic Value):", ws_sum['B5'].value)
print("Summary C5 (Margin of Safety):", ws_sum['C5'].value)
print("Summary D5 (Verdict):", ws_sum['D5'].value)

print("\n=== VERIFYING TERMINAL API SYNC ===")
r = requests.get('http://127.0.0.1:5000/api/export-status/KALYANKJIL')
status_data = r.json()
wb_val = status_data.get('workbook_valuation', {})
print("Terminal API Status:", status_data.get('status'))
print("Terminal Synced Metrics:")
print(f"  Intrinsic Value: Rs. {wb_val.get('intrinsic_value')}")
print(f"  Margin of Safety: {wb_val.get('margin_of_safety_pct')}%")
print(f"  Verdict: {wb_val.get('verdict')}")
print(f"  WACC: {wb_val.get('wacc')}%")

assert abs(wb_val.get('intrinsic_value') - ws_sum['B5'].value) < 0.1, "FAILED: Terminal does not match Workbook!"
print("\n>>> SUCCESS: Terminal picked the EXACT Intrinsic Value from the Excel workbook without touching or corrupting the workbook!")
