import openpyxl

wb = openpyxl.load_workbook('ITC Model.xlsx', data_only=False)
ws = wb['WACC']
print("=== WACC Sheet Inspection (ITC Model.xlsx) ===")
for r in range(10, 50):
    row_vals = []
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']:
        v = ws[f'{col}{r}'].value
        if v is not None:
            row_vals.append(f"{col}{r}: {repr(v)}")
    if row_vals:
        print(f"Row {r:2d}: " + " | ".join(row_vals))
