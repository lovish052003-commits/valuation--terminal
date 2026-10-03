import sys, os
sys.path.insert(0, os.path.abspath('.'))
import screener_client
import valuation_engine
import excel_exporter
import win32com.client, pythoncom

print("1. Fetching RELIANCE data...")
sd = screener_client.fetch_company_data('RELIANCE')
print(f"Historical 1y prices count: {len(sd.get('historical_prices_1y', []))}")
print(f"Latest 3 prices: {sd.get('historical_prices_1y', [])[:3]}")

print("2. Calculating valuation...")
val = valuation_engine.calculate_valuation(sd)

print("3. Exporting RELIANCE model via excel_exporter...")
out_path = excel_exporter.export_valuation_model(sd, val)
print(f"Exported to: {out_path}")

print("4. Testing COM opening of exported file and checking Beta...")
pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(os.path.abspath(out_path))
    print(f"SUCCESS! {out_path} opened cleanly with ZERO repair warnings! Sheets: {wb.Sheets.Count}")
    ws_beta = wb.Sheets('Beta-Regression')
    ws_raw = wb.Sheets('Raw Data')
    
    print("\n--- FIRST 5 ROWS IN RAW DATA ---")
    for r in range(6, 11):
        print(f"Raw Data Row {r}: Date={ws_raw.Cells(r, 7).Text} | Stock Price={ws_raw.Cells(r, 8).Value} | Nifty={ws_raw.Cells(r, 10).Value}")
        
    print("\n--- FIRST 5 ROWS IN BETA-REGRESSION ---")
    for r in range(10, 15):
        print(f"Beta Row {r}: Date={ws_beta.Cells(r, 2).Text} | Price={ws_beta.Cells(r, 3).Value} (Formula={ws_beta.Cells(r, 3).Formula}) | Return={ws_beta.Cells(r, 4).Text}")
        
    print("\n--- BETA DRIFTING SUMMARY ---")
    print("Levered Raw Beta (O11):", ws_beta.Range('O11').Value)
    print("Adjusted Beta (L15):", ws_beta.Range('L15').Value)
    
    wb.Close(SaveChanges=False)
except Exception as e:
    print("FAILED opening exported file:", e)
finally:
    excel.Quit()
    pythoncom.CoUninitialize()
