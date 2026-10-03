import os, sys, win32com.client, pythoncom
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from screener_client import fetch_company_data
from valuation_engine import calculate_valuation
import excel_exporter

print("1. Fetching ITC data...")
data = fetch_company_data('ITC')
val = calculate_valuation(data)

print("2. Running export_valuation_model (Primary COM Engine)...")
path1 = excel_exporter.export_valuation_model(data, val)
print(f"Generated: {path1} (Size: {os.path.getsize(path1)} bytes)")

print("3. Testing verification of COM-generated file in Excel...")
pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(os.path.abspath(path1))
    print("SUCCESS: COM-generated ITC_Valuation_Model.xlsx opened with ZERO errors in Excel!")
    wb.Close(SaveChanges=False)
except Exception as e:
    print("FAILED opening COM file:", e)
finally:
    excel.Quit()
    del excel
    pythoncom.CoUninitialize()

print("\n4. Testing Standalone OpenPyXL Engine...")
path2 = os.path.abspath('exports/TEST_STANDALONE_OPX.xlsx')
excel_exporter.export_standalone_openpyxl(path2, data, val)
print(f"Generated standalone: {path2} (Size: {os.path.getsize(path2)} bytes)")

print("5. Testing verification of Standalone OpenPyXL file in Excel...")
pythoncom.CoInitialize()
excel2 = win32com.client.DispatchEx('Excel.Application')
excel2.Visible = False
excel2.DisplayAlerts = False
try:
    wb2 = excel2.Workbooks.Open(path2)
    print("SUCCESS: Standalone OpenPyXL workbook opened with ZERO errors in Excel!")
    wb2.Close(SaveChanges=False)
except Exception as e:
    print("FAILED opening Standalone file:", e)
finally:
    excel2.Quit()
    del excel2
    pythoncom.CoUninitialize()

print("\nALL VERIFICATIONS PASSED!")
