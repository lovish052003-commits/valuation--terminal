import openpyxl

wb = openpyxl.load_workbook('exports/ADANIENT_Valuation_Model.xlsx', data_only=False)
ws_dcf = wb['DCF']

print("=== DCF SHEET FORMULAS ===")
for r in range(7, 45):
    for col in ['D', 'I', 'J', 'K', 'L', 'M']:
        c_ref = f"{col}{r}"
        v = ws_dcf[c_ref].value
        if v is not None and str(v).startswith('='):
            print(f"  {c_ref}: {v}")

wb_val = openpyxl.load_workbook('exports/ADANIENT_Valuation_Model.xlsx', data_only=True)
ws_dcf_val = wb_val['DCF']
print("\n=== DCF SHEET VALUES ===")
for r in [8, 10, 11, 12, 14, 16, 20, 21, 29, 33, 34, 35, 37, 38, 39, 40, 42, 44]:
    vals = [f"{c}{r}={ws_dcf_val[f'{c}{r}'].value}" for c in ['D', 'I', 'J', 'K', 'L', 'M'] if ws_dcf_val[f'{c}{r}'].value is not None]
    if vals:
        print(f"  Row {r:2d}:", " | ".join(vals))
