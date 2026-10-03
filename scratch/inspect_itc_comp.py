import openpyxl

wb = openpyxl.load_workbook('ITC Model.xlsx', data_only=False)
ws = wb['Comp_Valuation']
print("=== ITC Model.xlsx Comp_Valuation Rows 30-40 ===")
for r in range(30, 41):
    vals = []
    for col in ['B', 'N', 'O', 'P', 'Q']:
        v = ws[f'{col}{r}'].value
        if v is not None:
            vals.append(f"{col}{r}: {v}")
    if vals:
        print(" | ".join(vals))
