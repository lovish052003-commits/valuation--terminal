import openpyxl

target_path = r'C:\Users\LENOVO\Downloads\test\SUNPHARMA_Valuation_Model (1).xlsx'
wb_f = openpyxl.load_workbook(target_path, data_only=False)
wb_d = openpyxl.load_workbook(target_path, data_only=True)

ws_f = wb_f['Beta-Regression']
ws_d = wb_d['Beta-Regression']

out_file = r'C:\Users\LENOVO\Downloads\Advance Financial Project\scratch\beta_details.txt'

with open(out_file, 'w', encoding='utf-8') as f:
    f.write("=== BETA REGRESSION COLS O TO V ===\n")
    for r in range(1, 25):
        row_f = {openpyxl.utils.get_column_letter(c): ws_f.cell(row=r, column=c).value for c in range(14, 22)}
        row_v = {openpyxl.utils.get_column_letter(c): ws_d.cell(row=r, column=c).value for c in range(14, 22)}
        f.write(f"Row {r:2d} Form: {row_f}\n")
        f.write(f"       Val : {row_v}\n")

print("Beta details written to scratch/beta_details.txt")
