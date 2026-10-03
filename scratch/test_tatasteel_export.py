import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client, valuation_engine, excel_exporter, openpyxl

print("Fetching Screener data for TATASTEEL...")
data = screener_client.fetch_company_data('TATASTEEL')
print(f"Company: {data['company_name']} | Industry: {data['industry']}")
print(f"52W High/Low: {data['high_52w']} / {data['low_52w']}")
print(f"About (Wiki): {data['about_company'][:80]}...")
print(f"Updates count: {len(data['recent_updates'])}")
print("Peers (first 4):", data['peers_df']['Company'].head(4).tolist())

val_res = valuation_engine.calculate_valuation(data)
excel_path = excel_exporter.export_valuation_model(data, val_res, "")
print(f"Exported to: {excel_path}")

wb = openpyxl.load_workbook(excel_path, data_only=True)

# 1. Data Sheet
if 'Data Sheet' in wb.sheetnames:
    ws_data = wb['Data Sheet']
    print("\n=== Data Sheet ===")
    print("  B1 (Company Name):", repr(ws_data['B1'].value))
    print("  B8 (CMP):", ws_data['B8'].value)
    print("  B17 (Sales Y1):", ws_data['B17'].value, "| K17 (Sales Latest):", ws_data['K17'].value)
    print("  K30 (PAT Latest):", ws_data['K30'].value)
    print("  K57 (Equity Cap Latest):", ws_data['K57'].value)
    print("  K59 (Debt Latest):", ws_data['K59'].value)
    print("  K69 (Cash Latest):", ws_data['K69'].value)
    print("  K70 (Shares Latest):", ws_data['K70'].value)

# 2. Raw FS
if 'Raw FS' in wb.sheetnames:
    ws_raw = wb['Raw FS']
    print("\n=== Raw FS ===")
    print("  L56 (Target Name):", repr(ws_raw['L56'].value))
    print("  M56 (Target CMP):", ws_raw['M56'].value)
    print("  AR56 (Target Mcap):", ws_raw['AR56'].value)
    print("  L57 (Peer 1):", repr(ws_raw['L57'].value))
    print("  L58 (Peer 2):", repr(ws_raw['L58'].value))

# 3. Comp_Valuation
ws_comp = wb['Comp_Valuation']
print("\n=== Comp_Valuation ===")
print("  B30:", ws_comp['B30'].value)
for r in range(12, 16):
    print(f"  Row {r}: {ws_comp.cell(r, 2).value} | Price: {ws_comp.cell(r, 4).value}")

# 4. DCF
if 'DCF' in wb.sheetnames:
    ws_dcf = wb['DCF']
    print("\n=== DCF Sheet ===")
    print("  D44 (CMP):", ws_dcf['D44'].value)
    print("  D42 (Intrinsic Value):", ws_dcf['D42'].value)
    print("  D35 (EV):", ws_dcf['D35'].value)
    print("  D37 (Cash):", ws_dcf['D37'].value)
    print("  D38 (Debt):", ws_dcf['D38'].value)
    print("  D40 (Shares Cr):", ws_dcf['D40'].value)

# 5. Dupont & Altman
ws_dup = wb['Dupont Analysis']
print("\n=== Dupont Analysis ===")
print("  B5:", repr(ws_dup['B5'].value))
print("  B8:", repr(str(ws_dup['B8'].value)[:80]))
print("  B37:", repr(str(ws_dup['B37'].value)[:80]))

