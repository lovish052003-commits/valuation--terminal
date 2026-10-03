import openpyxl

wb = openpyxl.load_workbook(r'../test/ADANIPOWER_Valuation_Model.xlsx', data_only=True)
ws_iv = wb['Intrinsic Valuation']
print("Data only values in Intrinsic Valuation:")
for r in [51, 52, 54, 55, 60, 67, 68, 69, 70, 71]:
    print(f"Row {r}: B={ws_iv[f'B{r}'].value}, I={ws_iv[f'I{r}'].value}, J={ws_iv[f'J{r}'].value}, K={ws_iv[f'K{r}'].value}, L={ws_iv[f'L{r}'].value}")

ws_dcf = wb['DCF']
print("\nData only values in DCF:")
for r in [11, 18, 19, 20, 21]:
    print(f"DCF Row {r}: D={ws_dcf[f'D{r}'].value}, H={ws_dcf[f'H{r}'].value}, I={ws_dcf[f'I{r}'].value}, M={ws_dcf[f'M{r}'].value}")

wb_form = openpyxl.load_workbook(r'../test/ADANIPOWER_Valuation_Model.xlsx', data_only=False)
ws_iv_f = wb_form['Intrinsic Valuation']
print("\nFormulas in Intrinsic Valuation:")
for r in [52, 54, 55, 60, 67, 68, 69, 70, 71]:
    print(f"Row {r}: B={ws_iv_f[f'B{r}'].value}, L={ws_iv_f[f'L{r}'].value}")
