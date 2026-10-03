import openpyxl

target_path = r'C:\Users\LENOVO\Downloads\test\SUNPHARMA_Valuation_Model (1).xlsx'
wb = openpyxl.load_workbook(target_path, data_only=False)
wb_d = openpyxl.load_workbook(target_path, data_only=True)

ws = wb['Data Sheet']
ws_d = wb_d['Data Sheet']

print("=== DATA SHEET ROWS 60 TO 100 ===")
for r in range(56, 100):
    a = ws.cell(row=r, column=1).value
    b = ws.cell(row=r, column=2).value
    b_v = ws_d.cell(row=r, column=2).value
    k = ws.cell(row=r, column=11).value
    k_v = ws_d.cell(row=r, column=11).value
    if a or b or k:
        print(f"Row {r:2d} | A: {str(a):<30} | B Form: {str(b):<20} | B Val: {str(b_v):<15} | K Form: {str(k):<20} | K Val: {str(k_v):<15}")
