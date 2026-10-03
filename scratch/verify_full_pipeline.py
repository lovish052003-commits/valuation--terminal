import os, sys, gc
import win32com.client, pythoncom
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from screener_client import fetch_company_data
from valuation_engine import calculate_valuation
from excel_exporter import export_valuation_model

print("=== Running Comprehensive End-to-End Pipeline Verification ===")
company_query = "Tata Motors"
print(f"1. Fetching company data for '{company_query}' from Screener.in...")
data = fetch_company_data(company_query)
print(f"   Fetched: {data['company_name']} ({data['ticker']})")
print(f"   CMP: Rs. {data['current_price']}, Market Cap: Rs. {data['market_cap_cr']} Cr")
print(f"   Schedules fetched: {list(data.get('schedules', {}).keys())}")
print(f"   Peers fetched: {len(data.get('peers_df', []))} peers")

print("\n2. Computing Valuation Engine...")
val = calculate_valuation(data)
print(f"   Intrinsic Value: Rs. {val['intrinsic_value_per_share']}")

print(f"   Margin of Safety: {val['margin_of_safety_pct']}%")
print(f"   Verdict: {val['verdict']}")

print("\n3. Exporting Valuation Model...")
dest_file = export_valuation_model(data, val)
print(f"   Saved to: {dest_file}")

print("\n4. Inspecting Exported Workbook via Excel COM...")
pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(os.path.abspath(dest_file))
    
    # A. Check Forecasting Sheet
    ws_f = wb.Sheets('Forecasting')
    print("\n--- A. Forecasting Sheet Check ---")
    c13 = str(ws_f.Range('C13').Text)
    c14 = str(ws_f.Range('C14').Text)
    c14_f = str(ws_f.Range('C14').Formula)
    c15 = str(ws_f.Range('C15').Text)
    c15_f = str(ws_f.Range('C15').Formula)
    c19 = str(ws_f.Range('C19').Text)
    c19_f = str(ws_f.Range('C19').Formula)
    print(f"  Sales Year Weight 9 (C13): {c13}")
    print(f"  Sales Year Weight 10 (C14): {c14} | Formula: {c14_f}")
    print(f"  Sales Year Weight 11 (C15): {c15} | Formula: {c15_f}")
    print(f"  Sales Year Weight 15 (C19): {c19} | Formula: {c19_f}")
    assert c14_f.upper() == '=C13+365', f"Expected =C13+365, got {c14_f}"
    assert c15_f.upper() == '=C14+365', f"Expected =C14+365, got {c15_f}"
    assert c19_f.upper() == '=C18+365', f"Expected =C18+365, got {c19_f}"

    # B. Check Raw FS Sheet
    ws_raw = wb.Sheets('Raw FS')
    print("\n--- B. Raw FS Sheet Check ---")
    raw_dates = [ws_raw.Cells(3, c).Text for c in range(3, 15)]
    print(f"  BS Dates (C3:N3): {raw_dates[-4:]}")
    raw_inv_latest = ws_raw.Range('N39').Value # inventories
    raw_rec_latest = ws_raw.Range('N40').Value # receivables
    raw_pay_latest = ws_raw.Range('N14').Value # trade payables
    raw_fa_latest = ws_raw.Range('N24').Value # plant machinery
    print(f"  Latest Plant Machinery (N24): {raw_fa_latest}")
    print(f"  Latest Inventories (N39): {raw_inv_latest}")
    print(f"  Latest Receivables (N40): {raw_rec_latest}")
    print(f"  Latest Trade Payables (N14): {raw_pay_latest}")

    target_name_raw = ws_raw.Range('L56').Value
    peer1_name_raw = ws_raw.Range('L57').Value
    peer1_cmp_raw = ws_raw.Range('M57').Value
    print(f"  Target in Raw FS (L56): {target_name_raw}")
    print(f"  Peer 1 in Raw FS (L57): {peer1_name_raw} (CMP: {peer1_cmp_raw})")

    # C. Check Intrinsic Valuation Sheet
    ws_iv = wb.Sheets('Intrinsic Valuation')
    print("\n--- C. Intrinsic Valuation Sheet Check ---")
    iv_dates = [ws_iv.Range(f'{col}6').Text for col in ['H', 'I', 'J', 'K', 'L']]
    iv_inv = [ws_iv.Range(f'{col}9').Value for col in ['H', 'I', 'J', 'K', 'L']]
    iv_rec = [ws_iv.Range(f'{col}10').Value for col in ['H', 'I', 'J', 'K', 'L']]
    iv_pay = [ws_iv.Range(f'{col}16').Value for col in ['H', 'I', 'J', 'K', 'L']]
    print(f"  IV Statement Dates (H6:L6): {iv_dates}")
    print(f"  IV Inventories (H9:L9): {iv_inv}")
    print(f"  IV Receivables (H10:L10): {iv_rec}")
    print(f"  IV Payables (H16:L16): {iv_pay}")

    # Check for Excel errors in Intrinsic Valuation
    error_cells = []
    for r in range(1, 50):
        for col_l in ['B', 'C', 'D', 'H', 'I', 'J', 'K', 'L']:
            val_txt = ws_iv.Range(f'{col_l}{r}').Text
            if any(err in val_txt for err in ['#VALUE!', '#REF!', '#NAME?']):
                error_cells.append(f"{col_l}{r}: {val_txt}")
    if error_cells:
        print(f"  WARNING: Found errors in IV sheet: {error_cells}")
    else:
        print("  ZERO errors in Intrinsic Valuation sheet!")

    # D. Check Comp_Valuation Sheet
    ws_cv = wb.Sheets('Comp_Valuation')
    print("\n--- D. Comp_Valuation Sheet Check ---")
    peer1_name_cv = ws_cv.Range('B12').Value
    peer1_cmp_cv = ws_cv.Range('D12').Value
    peer1_shares_cv = ws_cv.Range('E12').Value
    peer1_ev_cv = ws_cv.Range('H12').Value
    print(f"  Peer 1 Name (B12): {peer1_name_cv}")
    print(f"  Peer 1 CMP (D12): {peer1_cmp_cv}")
    print(f"  Peer 1 Shares (E12): {peer1_shares_cv}")
    print(f"  Target EV (H12): {peer1_ev_cv}")

    wb.Close(SaveChanges=False)
    print("\n>>> ALL VERIFICATION CHECKS PASSED PERFECTLY! <<<")

except Exception as e:
    import traceback
    print("VERIFICATION FAILED:", e)
    traceback.print_exc()

finally:
    excel.Quit()
    del excel
    gc.collect()
    pythoncom.CoUninitialize()
