import openpyxl

wb = openpyxl.load_workbook('ITC Model.xlsx', data_only=False)
ws = wb['Historical FS']
print("=== Historical FS Rows 1 to 45 ===")
for r in range(1, 46):
    lbl = ws.cell(r, 2).value
    c_form = ws.cell(r, 3).value
    k_form = ws.cell(r, 11).value
    if lbl or c_form:
        print(f"Row {r:2d} | Label: {str(lbl):35s} | C3: {str(c_form):30s} | K3: {str(k_form):30s}")
