import openpyxl

wb_formula = openpyxl.load_workbook('exports/ADANIENT_Valuation_Model.xlsx', data_only=False)
wb_val = openpyxl.load_workbook('exports/ADANIENT_Valuation_Model.xlsx', data_only=True)
ws_f = wb_formula['DCF']
ws_v = wb_val['DCF']

print("--- DCF Rows 13 to 17 ---")
for r in range(13, 18):
    for c in ['H', 'I', 'J', 'K', 'L', 'M']:
        print(f"{c}{r}: formula={ws_f[f'{c}{r}'].value} | value={ws_v[f'{c}{r}'].value}")

print("\n--- DCF Column D Rows 18 to 44 ---")
for r in range(18, 45):
    label = ws_v[f'C{r}'].value
    cf = ws_f[f'D{r}'].value
    cv = ws_v[f'D{r}'].value
    print(f"D{r} ({label}): formula={cf} | value={cv}")
