import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import openpyxl
import excel_exporter
import screener_client
import valuation_engine

def test_peer_comps_standing_rules():
    print("=" * 80)
    print("TESTING PEER-COMPARABLES STANDING RULES (OPENPYXL & COM)")
    print("=" * 80)

    ticker = "SUZLON"
    print(f"\n1. Fetching live Screener.in data for {ticker}...")
    screener_data = screener_client.fetch_company_data(ticker)
    val_result = valuation_engine.calculate_valuation(screener_data)

    print(f"Target Company: {screener_data['company_name']} ({screener_data['ticker']})")
    print(f"Price: Rs. {screener_data['current_price']} | Shares: {screener_data['shares_in_cr']} Cr | MCap: Rs. {screener_data['market_cap_cr']} Cr")

    # Test openpyxl population and validation
    print("\n2. Testing OpenPyXL population of Raw FS & Comp_Valuation...")
    wb = openpyxl.load_workbook(excel_exporter.TEMPLATE_PATH)
    
    # Populate Data Sheet, Raw FS, Comp_Valuation, and AI Summary
    excel_exporter.populate_data_sheet_openpyxl(wb, screener_data)
    excel_exporter.populate_raw_fs_sheet_openpyxl(wb, screener_data, val_result)
    excel_exporter.update_comp_valuation_sheet_openpyxl(wb, screener_data, val_result)
    excel_exporter.populate_ai_summary_sheet_openpyxl(wb, screener_data, val_result)

    # Inspect Raw FS Row 56 (Target Company - Rule 2: Must trace back to Data Sheet)
    ws_raw = wb['Raw FS']
    ws_ds = wb['Data Sheet']
    r56_name_form = ws_raw['L56'].value
    r56_cmp_form = ws_raw['M56'].value
    r56_shares_form = ws_raw['N56'].value
    r56_mcap_form = ws_raw['AR56'].value
    print(f"\n[Raw FS Row 56 Formulas (Target)] Name: {r56_name_form} | CMP: {r56_cmp_form} | Shares: {r56_shares_form} | MCap: {r56_mcap_form}")
    assert r56_name_form == "='Data Sheet'!B1", f"L56 must link to 'Data Sheet'!B1, got {r56_name_form}"
    assert r56_cmp_form == "='Data Sheet'!B8", f"M56 must link to 'Data Sheet'!B8, got {r56_cmp_form}"
    assert r56_shares_form == "='Data Sheet'!B6", f"N56 must link to 'Data Sheet'!B6, got {r56_shares_form}"
    assert r56_mcap_form == "='Data Sheet'!B9", f"AR56 must link to 'Data Sheet'!B9, got {r56_mcap_form}"

    # Verify Data Sheet values
    ds_name = ws_ds['B1'].value
    ds_cmp = float(ws_ds['B8'].value)
    ds_mcap = float(ws_ds['B9'].value)
    ds_shares = float(screener_data['shares_in_cr'])
    print(f"[Data Sheet Source] Name: {ds_name} | CMP: {ds_cmp} | MCap: {ds_mcap} | Shares: {ds_shares}")
    calc_56 = round(ds_cmp * ds_shares, 2)
    diff_56 = abs(calc_56 - ds_mcap) / max(ds_mcap, 1.0) * 100
    print(f"Row 56 % Diff: {diff_56:.2f}% (Tolerance <= 2.0%)")
    assert diff_56 <= 2.0, f"Row 56 diff {diff_56:.2f}% > 2.0%"

    # Inspect Comp_Valuation Row 12 (Peer 1)
    ws_comp = wb['Comp_Valuation']
    b12 = ws_comp['B12'].value
    d12 = ws_comp['D12'].value
    e12 = ws_comp['E12'].value
    f12 = ws_comp['F12'].value
    g12 = ws_comp['G12'].value
    h12 = ws_comp['H12'].value
    i12 = ws_comp['I12'].value
    j12 = ws_comp['J12'].value
    print(f"\n[Comp_Valuation Row 12 (Peer 1)]")
    print(f"  B12: {b12}")
    print(f"  D12: {d12}")
    print(f"  E12: {e12}")
    print(f"  F12: {f12}")
    print(f"  G12: {g12}")
    print(f"  H12: {h12}")
    print(f"  I12: {i12}")
    print(f"  J12: {j12}")
    assert b12 == "='Raw FS'!L57", f"B12 must link to Raw FS Row 57, got {b12}"
    assert g12 == "='Raw FS'!AS57", f"G12 must link to Raw FS Row 57, got {g12}"
    assert h12 == "='Raw FS'!AT57", f"H12 must link to Raw FS Row 57, got {h12}"

    # Inspect Target Valuation Rows (Rows 32, 33, 35)
    o32 = ws_comp['O32'].value
    o33 = ws_comp['O33'].value
    o35 = ws_comp['O35'].value
    print(f"\n[Comp_Valuation Target Valuation (Rows 32, 33, 35)]")
    print(f"  O32: {o32}")
    print(f"  O33: {o33}")
    print(f"  O35: {o35}")
    assert "56" in str(o32), f"O32 must link to Raw FS Row 56, got {o32}"
    assert "56" in str(o33), f"O33 must link to Raw FS Row 56, got {o33}"
    assert "56" in str(o35), f"O35 must link to Raw FS Row 56, got {o35}"

    # Save to temp file and run full post-generation validation
    test_out = os.path.join(excel_exporter.EXPORT_DIR, "test_suzlon_standing_rules.xlsx")
    wb.save(test_out)
    wb.close()

    print("\n3. Running validate_generated_workbook...")
    excel_exporter.validate_generated_workbook(test_out, ticker, screener_data['company_name'], is_financial=False)
    print("\nALL STANDING RULES VERIFIED SUCCESSFULLY!")

if __name__ == '__main__':
    test_peer_comps_standing_rules()
