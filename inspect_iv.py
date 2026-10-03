import openpyxl

wb = openpyxl.load_workbook(r'C:\Users\LENOVO\Downloads\test\ADANIENT_Valuation_Model (8).xlsx', data_only=False)
ws_iv = wb['Intrinsic Valuation']
for r in range(45, 76):
    b_val = ws_iv[f'B{r}'].value
    l_val = ws_iv[f'L{r}'].value
    if b_val is not None or l_val is not None:
        print(f"Row {r}: B={b_val} | L={l_val}")
