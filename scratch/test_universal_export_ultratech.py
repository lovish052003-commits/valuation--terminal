import os
import sys
import openpyxl

sys.path.insert(0, os.path.abspath('.'))

import screener_client
import valuation_engine
import excel_exporter

def test_export():
    ticker = 'ULTRACEMCO'
    print(f"1. Fetching screener data for {ticker}...")
    sd = screener_client.fetch_company_data(ticker)
    print(f"2. Calculating valuation for {sd.get('company_name')}...")
    val = valuation_engine.calculate_valuation(sd)
    print("3. Exporting valuation model...")
    exported_path = excel_exporter.export_valuation_model(sd, val)
    print(f"Exported to: {exported_path}")

    # Inspect the exported workbook
    wb = openpyxl.load_workbook(exported_path, data_only=False)
    
    print("\n--- Verifying Exported Sheets ---")
    sheet_names = wb.sheetnames
    print(f"Total Sheets: {len(sheet_names)}")
    assert 'Control' in sheet_names, "Control sheet missing!"
    assert 'Checks' in sheet_names, "Checks sheet missing!"
    assert 'AI Valuation Summary' in sheet_names, "AI Valuation Summary missing!"
    assert 'DCF' in sheet_names, "DCF sheet missing!"

    # 1. Control Sheet Section 4
    ws_ctrl = wb['Control']
    print(f"Control B32: {ws_ctrl['B32'].value}")
    print(f"Control B33: {ws_ctrl['B33'].value} | C33: {ws_ctrl['C33'].value}")
    print(f"Control B34: {ws_ctrl['B34'].value} | C34: {ws_ctrl['C34'].value}")
    print(f"Control B35: {ws_ctrl['B35'].value} | C35: {ws_ctrl['C35'].value}")
    print(f"Control B36: {ws_ctrl['B36'].value} | C36: {ws_ctrl['C36'].value}")
    assert ws_ctrl['B34'].value == "Exit EV/EBIT Multiple (x)", "Control B34 mismatch!"
    assert ws_ctrl['C35'].value == "Gordon Growth", "Control C35 mismatch!"

    # 2. Checks Sheet Check 8 and Advisory Checks
    ws_chk = wb['Checks']
    print(f"Checks D13: {ws_chk['D13'].value}")
    print(f"Checks B17: {ws_chk['B17'].value}")
    print(f"Checks B18: {ws_chk['B18'].value} | D18: {ws_chk['D18'].value}")
    print(f"Checks B19: {ws_chk['B19'].value} | D19: {ws_chk['D19'].value}")
    print(f"Checks B20: {ws_chk['B20'].value} | D20: {ws_chk['D20'].value}")
    print(f"Checks B21: {ws_chk['B21'].value} | D21: {ws_chk['D21'].value}")
    assert "DATEVALUE" in ws_chk['D13'].value, "Check 8 DATEVALUE formula missing!"
    assert ws_chk['B18'].value == "A1. DCF vs Market Divergence", "Advisory check A1 missing!"

    # 3. DCF 10-Year Alternate DCF
    ws_dcf = wb['DCF']
    print(f"DCF B58: {ws_dcf['B58'].value}")
    print(f"DCF B65: {ws_dcf['B65'].value} | D65: {ws_dcf['D65'].value}")
    print(f"DCF B81: {ws_dcf['B81'].value} | D81: {ws_dcf['D81'].value}")
    assert ws_dcf['B58'].value == "ALTERNATE 10-YEAR DCF (parallel view - does not feed the Summary verdict; inputs on Control rows 33-36)", "DCF B58 mismatch!"
    assert ws_dcf['B81'].value == "Alt DCF value per share", "DCF B81 mismatch!"

    # 4. Intrinsic Valuation Organic Cross-Check
    ws_iv = wb['Intrinsic Valuation']
    print(f"Intrinsic Valuation B67: {ws_iv['B67'].value}")
    print(f"Intrinsic Valuation B70: {ws_iv['B70'].value} | L70: {ws_iv['L70'].value}")
    assert ws_iv['B67'].value == "ORGANIC REINVESTMENT CROSS-CHECK", "IV B67 mismatch!"

    # 5. AI Valuation Summary
    ws_ai = wb['AI Valuation Summary']
    print(f"AI Summary I4: {ws_ai['I4'].value} | I5: {ws_ai['I5'].value}")
    print(f"AI Summary D10: {ws_ai['D10'].value}")
    print(f"AI Summary D12: {ws_ai['D12'].value}")
    print(f"AI Summary B41: {ws_ai['B41'].value} | C41: {ws_ai['C41'].value}")
    print(f"AI Summary B42: {ws_ai['B42'].value} | C42: {ws_ai['C42'].value}")
    assert ws_ai['I4'].value == "Verdict Confidence", "AI Summary I4 mismatch!"
    assert ws_ai['B41'].value == "Intrinsic Value rolled to Valuation Date (info only)", "AI Summary B41 mismatch!"
    assert ws_ai['B42'].value == "Alt 10-Year DCF Value (info only; set on Control rows 33-36)", "AI Summary B42 mismatch!"

    wb.close()
    print("\nALL VERIFICATIONS PASSED UNIVERSALLY!")

if __name__ == '__main__':
    test_export()
