import os, sys, time

sys.stdout.reconfigure(line_buffering=True)
sys.path.insert(0, os.path.abspath('.'))

import screener_client
import valuation_engine
import excel_exporter

ticker = 'UNITEDTEA'
print(f"--- Testing End-to-End Export for {ticker} ---")
data = screener_client.fetch_company_data(ticker)
print("Data fetched:", data.get('company_name'))

val = valuation_engine.calculate_valuation(data)
print("Valuation calculated. Intrinsic Value:", val.get('intrinsic_value'))

t0 = time.time()
exported_file = excel_exporter.export_valuation_model(data, val)
print(f"Export completed in {time.time()-t0:.2f}s -> {exported_file}")

# Verify file exists and its size
sz = os.path.getsize(exported_file)
print(f"Exported file size: {sz} bytes")

# Verify relationship sanity
import zipfile, re
with zipfile.ZipFile(exported_file, 'r') as z:
    bad_targets = []
    for n in z.namelist():
        if n.endswith('.rels'):
            d = z.read(n).decode('utf-8', errors='ignore')
            targets = re.findall(r'Target="([^"]+)"', d)
            bad_targets.extend([t for t in targets if t.startswith('/xl/') or t.startswith('/')])
    print(f"Bad targets count in {exported_file}: {len(bad_targets)}")
    ext_links = [n for n in z.namelist() if 'external' in n.lower()]
    print(f"External links in {exported_file}: {ext_links}")
    charts = [n for n in z.namelist() if 'chart' in n.lower()]
    print(f"Chart files count in {exported_file}: {len(charts)}")

# Test opening in Excel COM
import pythoncom, win32com.client
pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath(exported_file), UpdateLinks=0, ReadOnly=False)
print("NATIVE EXCEL COM OPEN SUCCESS!")
print(f"Sheets count: {len(wb.Sheets)}")
wb.Close(False)
excel.Quit()
print("TEST COMPLETED WITH ZERO ERRORS!")
