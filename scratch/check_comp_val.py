import openpyxl

wb = openpyxl.load_workbook('ITC Model.xlsx', data_only=False)
ws = wb['Comp_Valuation']
for r in range(10, 25):
    row_vals = [ws.cell(r, c).value for c in range(1, 12)]
    if any(x is not None for x in row_vals):
        print(f"Row {r}: {row_vals}")
wb.close()
