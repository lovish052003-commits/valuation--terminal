import os, sys, zipfile, openpyxl
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__) + '/..'))

import screener_client
import valuation_engine
import excel_exporter

print("=== STEP 1: Fetching Screener Data for TATASTEEL ===")
data = screener_client.fetch_company_data("TATASTEEL")
print(f"Company: {data.get('company_name')} ({data.get('ticker')})")
print(f"Website: {data.get('company_website')}")
print(f"Peers Count: {len(data.get('peers_df', []))}")

print("\n=== STEP 2: Running Valuation Engine ===")
val = valuation_engine.calculate_valuation(data)
print(f"Default Tax Rate in Valuation: {val.get('tax_rate')}%")
print(f"Intrinsic Value: Rs. {val.get('intrinsic_value_per_share')}")
print(f"WACC: {val.get('wacc')}%")

print("\n=== STEP 3: Exporting Valuation Model ===")
export_path = excel_exporter.export_valuation_model(data, val)
print(f"Exported to: {export_path}")

print("\n=== STEP 4: Validating Exported Excel Workbook ===")
wb = openpyxl.load_workbook(export_path, data_only=False)

# A. Raw Data Sheet check
ws_raw = wb['Raw Data']
print("\n--- Raw Data Sheet (Rows 24 to 28, Cols O to T) ---")
for r in range(24, 29):
    c_name = ws_raw.cell(r, 15).value
    country = ws_raw.cell(r, 17).value
    debt = ws_raw.cell(r, 18).value
    mcap = ws_raw.cell(r, 19).value
    tax = ws_raw.cell(r, 20).value
    print(f"  Row {r}: Name='{c_name}', Country='{country}', Debt={debt}, Mcap={mcap}, Tax={tax}")

# B. WACC Sheet check
ws_wacc = wb['WACC']
print("\n--- WACC Sheet Key Cells ---")
print(f"  C34 (Target Debt): {ws_wacc['C34'].value}")
print(f"  C35 (Target Market Cap): {ws_wacc['C35'].value}")
print(f"  E27 (Target Tax Rate): {ws_wacc['E27'].value}")
print(f"  Row 14 (Peer 1): Name={ws_wacc['B14'].value}, Debt={ws_wacc['E14'].value}, Mcap={ws_wacc['F14'].value}, Tax={ws_wacc['G14'].value}")
print(f"  Row 15 (Peer 2): Name={ws_wacc['B15'].value}, Debt={ws_wacc['E15'].value}, Mcap={ws_wacc['F15'].value}, Tax={ws_wacc['G15'].value}")

# C. Zip images check
print("\n--- Zip Images Check ---")
with zipfile.ZipFile('ITC Model.xlsx', 'r') as z_tmpl:
    tmpl_img2 = z_tmpl.read('xl/media/image2.png')
    tmpl_img3 = z_tmpl.read('xl/media/image3.png')

with zipfile.ZipFile(export_path, 'r') as z_exp:
    exp_img2 = z_exp.read('xl/media/image2.png')
    exp_img3 = z_exp.read('xl/media/image3.png')

print(f"  Template image2 size: {len(tmpl_img2)}, Export image2 size: {len(exp_img2)} -> {'UNTOUCHED (MATCH!)' if tmpl_img2 == exp_img2 else 'ERROR: MODIFIED!'}")
print(f"  Template image3 size: {len(tmpl_img3)}, Export image3 size: {len(exp_img3)} -> {'UPDATED WITH REAL LOGO!' if tmpl_img3 != exp_img3 else 'NOT UPDATED'}")

wb.close()
print("\n=== ALL VALIDATION TESTS COMPLETED! ===")
