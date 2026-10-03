import openpyxl

wb = openpyxl.load_workbook('exports/NESTLEIND_Valuation_Model.xlsx', data_only=False)
ws = wb['DCF']
wb_d = openpyxl.load_workbook('exports/NESTLEIND_Valuation_Model.xlsx', data_only=True)
ws_d = wb_d['DCF']

for r in range(35, 47):
    b = ws.cell(row=r, column=2).value
    d_f = ws.cell(row=r, column=4).value
    d_v = ws_d.cell(row=r, column=4).value
    print(f"R{r:02d}: B='{b}' | D_form='{d_f}' | D_val='{d_v}'")
