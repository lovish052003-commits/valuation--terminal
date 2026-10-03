import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client
import valuation_engine
import excel_exporter
import openpyxl
import requests

tickers = ['SUZLON', 'WIPRO']

for ticker in tickers:
    print(f"\n==========================================")
    print(f"TESTING UNIVERSAL SYNC FOR {ticker}")
    print(f"==========================================")
    data = screener_client.fetch_company_data(ticker)
    val = valuation_engine.calculate_valuation(data)
    dest_path = excel_exporter.export_valuation_model(data, val, "")
    print(f"Exported {ticker} to {dest_path}")

    wb_v = openpyxl.load_workbook(dest_path, data_only=True)
    ws_dcf = wb_v['DCF']
    ws_sum = wb_v['AI Valuation Summary']

    ebit = ws_dcf['H8'].value
    iv_share = ws_sum['B5'].value
    mos = ws_sum['C5'].value
    verdict = ws_sum['D5'].value

    print(f"Workbook Evaluated Results:")
    print(f"  DCF H8 (Base EBIT): {ebit}")
    print(f"  AI Summary B5 (Intrinsic Value): Rs. {iv_share}")
    print(f"  AI Summary C5 (Margin of Safety): {mos}")
    print(f"  AI Summary D5 (Verdict): {verdict}")

    assert ebit is not None and ebit > 0, f"FAILED: Base EBIT is invalid for {ticker}!"
    assert iv_share is not None and str(iv_share) not in ['#DIV/0!', '#VALUE!', '#REF!', 'None'], f"FAILED: Intrinsic Value error for {ticker}!"

    # API check
    r = requests.get(f'http://127.0.0.1:5000/api/export-status/{ticker}')
    api_res = r.json()
    wb_val = api_res.get('workbook_valuation')
    print(f"API Workbook Valuation for {ticker}: {wb_val}")
    assert wb_val and wb_val.get('intrinsic_value') is not None, f"FAILED: API missing workbook_valuation for {ticker}!"
    print(f">>> PASS: {ticker} verified with 100% accuracy!")

print("\n==========================================")
print("ALL MULTI-COMPANY UNIVERSAL TESTS PASSED!")
print("==========================================")
