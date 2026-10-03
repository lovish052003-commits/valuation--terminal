import openpyxl, sys

sys.stdout.reconfigure(encoding='utf-8')
wb = openpyxl.load_workbook(r'C:\Users\LENOVO\Downloads\test\ADANIENT_Valuation_Model (8).xlsx', data_only=False)
ws_ai = wb['AI Valuation Summary']
for r in range(40, 100):
    row_strs = []
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
        val = ws_ai[f'{col}{r}'].value
        if val is not None:
            row_strs.append(f"{col}{r}: {val}")
    if row_strs:
        print(f"Row {r} -> " + " | ".join(row_strs))
