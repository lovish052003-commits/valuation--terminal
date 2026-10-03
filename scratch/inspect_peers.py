import openpyxl

wb_t = openpyxl.load_workbook('Tata Steel Final Model.xlsx', data_only=False)
ws_t = wb_t['Comp_Valuation']

wb_m = openpyxl.load_workbook('master_model_template.xlsx', data_only=False)
ws_m = wb_m['Comp_Valuation']

print('--- TATA STEEL PEER ROWS (12-22) ---')
for r in range(12, 22):
    b = ws_t[f'B{r}'].value
    if b:
        print(f"Row {r}: B={b}, C={ws_t[f'C{r}'].value}, G={ws_t[f'G{r}'].value}, K={ws_t[f'K{r}'].value}")

print('--- MASTER PEER ROWS (12-22) ---')
for r in range(12, 22):
    b = ws_m[f'B{r}'].value
    print(f"Row {r}: B={b}, C={ws_m[f'C{r}'].value}, G={ws_m[f'G{r}'].value}, K={ws_m[f'K{r}'].value}")
