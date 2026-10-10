import sys, os, time
sys.path.insert(0, os.path.abspath('.'))
import pythoncom
import win32com.client
import screener_client
import valuation_engine
import excel_exporter

ticker = 'NESTLEIND'
data = screener_client.fetch_company_data(ticker)
val = valuation_engine.calculate_valuation(data)
test_out = os.path.abspath('exports/test_com_nestle.xlsx')
if os.path.exists(test_out):
    try: os.remove(test_out)
    except: pass

print("Testing export_via_excel_com...")
start = time.time()
try:
    res = excel_exporter.export_via_excel_com(test_out, data, val, "")
    print("export_via_excel_com returned:", res, f"in {time.time()-start:.2f}s")
except Exception as e:
    import traceback
    traceback.print_exc()
    print("export_via_excel_com failed:", e)
