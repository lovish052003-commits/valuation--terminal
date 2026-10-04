import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import screener_client
import valuation_engine
import excel_exporter
import openpyxl

data = screener_client.load_offline_company_data('ITC')
val = valuation_engine.calculate_valuation(data)

dest_path = os.path.join(excel_exporter.EXPORT_DIR, "TEST_ITC_Valuation_Model.xlsx")
excel_exporter.export_via_excel_com(dest_path, data, val)
print('Exported to:', dest_path)

wb = openpyxl.load_workbook(dest_path, data_only=False)
ws_iv = wb['Intrinsic Valuation']
ws_dcf = wb['DCF']

print('--- Intrinsic Valuation Formulas ---')
print('H51:', ws_iv['H51'].value)
print('H52:', ws_iv['H52'].value)
print('I52:', ws_iv['I52'].value)
print('L52:', ws_iv['L52'].value)

print('--- DCF Formulas ---')
print('D18:', ws_dcf['D18'].value)

h52 = str(ws_iv['H52'].value)
l52 = str(ws_iv['L52'].value)
d18 = str(ws_dcf['D18'].value)

assert 'H51/H49' in h52, f"Expected H51/H49, got {h52}"
assert 'L51/L49' in l52, f"Expected L51/L49, got {l52}"
assert 'L62' in d18, f"Expected L62 in D18, got {d18}"
print("ALL FORMULA CHECKS VERIFIED SUCCESSFULLY!")
