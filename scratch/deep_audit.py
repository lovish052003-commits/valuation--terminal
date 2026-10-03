import openpyxl

target_path = r'C:\Users\LENOVO\Downloads\test\SUNPHARMA_Valuation_Model (1).xlsx'
wb = openpyxl.load_workbook(target_path, data_only=False)
wb_d = openpyxl.load_workbook(target_path, data_only=True)

ws = wb['Data Sheet']
ws_d = wb_d['Data Sheet']

out_file = r'C:\Users\LENOVO\Downloads\Advance Financial Project\scratch\datasheet_deep_audit.txt'

with open(out_file, 'w', encoding='utf-8') as f:
    f.write("=== DATA SHEET ROWS 1 TO 55 ===\n")
    for r in range(1, 56):
        a_val = ws.cell(row=r, column=1).value
        b_val = ws.cell(row=r, column=2).value
        k_val = ws.cell(row=r, column=11).value
        k_d = ws_d.cell(row=r, column=11).value
        b_d = ws_d.cell(row=r, column=2).value
        f.write(f"Row {r:2d} | A: {str(a_val):<30} | B Form: {str(b_val):<20} | B Val: {str(b_d):<15} | K Form: {str(k_val):<20} | K Val: {str(k_d):<15}\n")

    f.write("\n=== HISTORICAL FS ROWS 1 TO 45 FORMULAS & VALUES ===\n")
    ws_h = wb['Historical FS']
    ws_hd = wb_d['Historical FS']
    for r in range(1, 45):
        lbl = ws_h.cell(row=r, column=2).value
        if lbl:
            f.write(f"Row {r:2d} | Label: {str(lbl):<35}\n")
            f.write(f"   Cols C..K Form: {[ws_h.cell(row=r, column=c).value for c in range(3, 12)]}\n")
            f.write(f"   Cols C..K Val : {[ws_hd.cell(row=r, column=c).value for c in range(3, 12)]}\n")

    f.write("\n=== FORECASTING SHEET AUDIT ===\n")
    if 'Forecasting' in wb.sheetnames:
        ws_fc = wb['Forecasting']
        ws_fcd = wb_d['Forecasting']
        for r in range(1, 60):
            lbl = ws_fc.cell(row=r, column=2).value
            if lbl:
                f.write(f"Row {r:2d} | Label: {str(lbl):<35}\n")
                f.write(f"   Cols C..K Form: {[ws_fc.cell(row=r, column=c).value for c in range(3, 12)]}\n")
                f.write(f"   Cols C..K Val : {[ws_fcd.cell(row=r, column=c).value for c in range(3, 12)]}\n")

print("Deep audit written to scratch/datasheet_deep_audit.txt")
