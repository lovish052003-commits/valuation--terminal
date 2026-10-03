import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
wb = openpyxl.load_workbook("ITC Model.xlsx", data_only=False)
ws = wb['AI Valuation Summary']

print("=== ITC Model.xlsx AI Valuation Summary ===")
for r in range(1, 40):
    row_strs = []
    for c in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
        v = ws[f'{c}{r}'].value
        if v is not None:
            row_strs.append(f"{c}{r}: {repr(v)}")
    if row_strs:
        print(" | ".join(row_strs))
