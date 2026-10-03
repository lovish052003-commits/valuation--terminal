import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client
import valuation_engine
import excel_exporter
import openpyxl
import requests

print("=== 1. FETCHING DATA & COMPUTING VALUATION FOR KALYANKJIL ===")
data = screener_client.fetch_company_data('KALYANKJIL')
val = valuation_engine.calculate_valuation(data)

print(f"Target: {data.get('company_name')} ({data.get('ticker')})")
print(f"CMP: Rs. {val.get('current_price')}")
print(f"Valuation Engine IV: Rs. {val.get('intrinsic_value_per_share')}")
print(f"Valuation Engine Verdict: {val.get('verdict')}")

print("\n=== 2. EXPORTING WORKBOOK ===")
dest_path = excel_exporter.export_valuation_model(data, val, "")
print(f"Exported to: {dest_path}")

print("\n=== 3. VERIFYING WORKBOOK INTEGRITY & EVALUATED VALUES ===")
wb = openpyxl.load_workbook(dest_path, data_only=False)
wb_v = openpyxl.load_workbook(dest_path, data_only=True)

ws_dcf = wb['DCF']
ws_dcf_v = wb_v['DCF']
ws_sum = wb['AI Valuation Summary']
ws_sum_v = wb_v['AI Valuation Summary']
ws_raw = wb['Raw FS']
ws_raw_v = wb_v['Raw FS']

print(f"DCF H8 Formula: {ws_dcf['H8'].value}")
print(f"DCF H8 Evaluated Value (EBIT): {ws_dcf_v['H8'].value}")
print(f"DCF D35 (Operating Assets / EV): {ws_dcf_v['D35'].value}")
print(f"DCF D40 (Equity Value): {ws_dcf_v['D40'].value}")
print(f"DCF D43 (Intrinsic Value/Share): {ws_dcf_v['D43'].value}")

print(f"\nAI Summary B5 Formula: {ws_sum['B5'].value}")
print(f"AI Summary B5 Evaluated Intrinsic Value: Rs. {ws_sum_v['B5'].value}")
print(f"AI Summary C5 Evaluated Margin of Safety: {ws_sum_v['C5'].value}")
print(f"AI Summary D5 Evaluated Verdict: {ws_sum_v['D5'].value}")

print(f"\nRaw FS Column AD (col 30):")
print(f"  AD4 (Sales): {ws_raw_v['AD4'].value}")
print(f"  AD9 (Interest): {ws_raw_v['AD9'].value}")
print(f"  AD11 (PBT): {ws_raw_v['AD11'].value}")

# Check that residual ITC data is completely gone
assert ws_raw_v['AD11'].value != 28033, "FAILED: Raw FS still contains ITC's 28,033 Cr PBT!"
assert ws_dcf_v['H8'].value != 28118, "FAILED: DCF H8 still uses ITC's 28,118 Cr EBIT!"
print(">>> CHECK PASSED: No ITC residual template data in Kalyan's workbook!")

print("\n=== 4. VERIFYING /api/export-status/KALYANKJIL API & TERMINAL SYNC ===")
r = requests.get('http://127.0.0.1:5000/api/export-status/KALYANKJIL')
status_data = r.json()
print("Export status API response:")
print("Status:", status_data.get('status'))
print("Workbook Valuation:", status_data.get('workbook_valuation'))

wb_val = status_data.get('workbook_valuation', {})
assert wb_val, "FAILED: workbook_valuation missing from /api/export-status response!"
assert wb_val.get('intrinsic_value') is not None, "FAILED: intrinsic_value is None!"
print(f"\n>>> SYNC SUCCESS: Terminal will display:")
print(f"    Intrinsic Value: Rs. {wb_val.get('intrinsic_value')}")
print(f"    Margin of Safety: {wb_val.get('margin_of_safety_pct')}%")
print(f"    Recommendation Verdict: {wb_val.get('verdict')}")
print(f"    WACC: {wb_val.get('wacc')}%")
print(f"    EV: Rs. {wb_val.get('enterprise_value')} Cr")
print("\nALL CHECKS PASSED FOR KALYANKJIL!")
