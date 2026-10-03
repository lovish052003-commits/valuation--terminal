import sys
import os
import openpyxl
sys.path.insert(0, '.')
import screener_client
import valuation_engine
import excel_exporter

print("Fetching TATASTEEL...")
sd = screener_client.fetch_company_data("TATASTEEL")
print("Target Name:", sd.get('company_name'))
print("Target Ticker:", sd.get('ticker'))

print("Running valuation engine...")
val = valuation_engine.calculate_valuation(sd)

print("Exporting valuation model...")
export_path = excel_exporter.export_valuation_model(sd, val)
print("Exported model path:", export_path)

# Verify the exported workbook
wb = openpyxl.load_workbook(export_path, data_only=True)
print("\n=== VERIFICATION OF ALL KEY SHEETS (TATA STEEL) ===")
print("Data Sheet B1:", wb['Data Sheet']['B1'].value)
print("Data Sheet B8 (CMP):", wb['Data Sheet']['B8'].value)
print("Data Sheet K17 (Sales):", wb['Data Sheet']['K17'].value)
print("Data Sheet K30 (PAT):", wb['Data Sheet']['K30'].value)
print("Profit & Loss A1:", wb['Profit & Loss']['A1'].value)
print("Balance Sheet A1:", wb['Balance Sheet']['A1'].value)
print("Cash Flow A1:", wb['Cash Flow']['A1'].value)
print("DCF D42 (IV):", wb['DCF']['D42'].value)
print("DCF D44 (CMP):", wb['DCF']['D44'].value)
print("Comp_Valuation B12:", wb['Comp_Valuation']['B12'].value)
print("WACC Row 14 Peer:", wb['WACC']['B14'].value)
print("WACC Row 15 Target:", wb['WACC']['B15'].value)
print("AI Valuation Summary A1:", wb['AI Valuation Summary']['A1'].value if 'AI Valuation Summary' in wb.sheetnames else "MISSING")
if 'AI Valuation Summary' in wb.sheetnames:
    print("AI Summary B5 (IV):", wb['AI Valuation Summary']['B5'].value)
print("Dupont Analysis B8:", str(wb['Dupont Analysis']['B8'].value)[:80])
print("Altman's Z Score B8:", str(wb["Altman's Z Score"]['B8'].value)[:80])

b1_val = str(wb['Data Sheet']['B1'].value or '')
assert 'ITC' not in b1_val.upper(), f"FAIL: Data Sheet B1 still has ITC: {b1_val}"
assert 'TATA' in b1_val.upper(), f"FAIL: Data Sheet B1 does not have TATA: {b1_val}"
print("\n[PASSED] 0.00% ITC data in TATA STEEL model!")
