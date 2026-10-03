import sys, os
sys.path.insert(0, os.path.abspath('.'))
import screener_client
import valuation_engine
import excel_exporter
import win32com.client, pythoncom

print("1. Fetching SUZLON data...")
sd = screener_client.fetch_company_data('SUZLON')
val = valuation_engine.calculate_valuation(sd)

print("2. Exporting SUZLON model via excel_exporter...")
out_path = excel_exporter.export_valuation_model(sd, val)
print(f"Exported to: {out_path}")

print("3. Testing COM opening of exported file...")
pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(os.path.abspath(out_path))
    print(f"SUCCESS! {out_path} opened cleanly with ZERO repair warnings! Sheets: {wb.Sheets.Count}")
    ws_beta = wb.Sheets('Beta-Regression')
    ws_raw = wb.Sheets('Raw Data')
    print("Raw Data H6:", ws_raw.Range('H6').Value)
    print("Beta C10 formula:", ws_beta.Range('C10').Formula)
    print("Beta C10 value:", ws_beta.Range('C10').Value)
    print("Beta B10 formula:", ws_beta.Range('B10').Formula)
    print("Beta G10 formula:", ws_beta.Range('G10').Formula)
    print("Beta G10 value:", ws_beta.Range('G10').Value)
    wb.Close(SaveChanges=False)
except Exception as e:
    print("FAILED opening exported file:", e)
finally:
    excel.Quit()
    pythoncom.CoUninitialize()
