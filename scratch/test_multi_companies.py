import sys, os
sys.path.insert(0, os.path.abspath('.'))
import screener_client
import valuation_engine
import excel_exporter
import openpyxl

for ticker in ['ABCAPITAL', 'SUZLON']:
    print(f"\n--- Testing {ticker} ---")
    sd = screener_client.fetch_company_data(ticker)
    val = valuation_engine.calculate_valuation(sd)
    out_path = excel_exporter.export_valuation_model(sd, val)
    print(f"Exported to: {out_path}")
    
    wb = openpyxl.load_workbook(out_path, data_only=False)
    ws_beta = wb['Beta-Regression']
    ws_raw = wb['Raw Data']
    
    print(f"Raw Data H6: {ws_raw['H6'].value}")
    print(f"Beta C10: {ws_beta['C10'].value} (Expected: ='Raw Data'!H6)")
    print(f"Beta C11: {ws_beta['C11'].value} (Expected: ='Raw Data'!H7)")
    print(f"Beta B10: {ws_beta['B10'].value} (Expected: ='Raw Data'!G6)")
    print(f"Beta G10: {ws_beta['G10'].value} (Expected: ='Raw Data'!J6)")
    print(f"Beta B7: {ws_beta['B7'].value}")
    
    assert ws_beta['C10'].value == "='Raw Data'!H6", f"Failed for {ticker}: C10 is not linked to Raw Data H6!"
    assert ws_beta['C11'].value == "='Raw Data'!H7", f"Failed for {ticker}: C11 is not linked to Raw Data H7!"
    print(f"PASS: {ticker} Beta-Regression is 100% connected to Raw Data!")
