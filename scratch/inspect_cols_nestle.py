import openpyxl

wb = openpyxl.load_workbook(r'C:\Users\LENOVO\Downloads\Nestle India (2).xlsx', data_only=False)
ws = wb['Data Sheet']

cols = [chr(65+i) for i in range(1, 12)] # A to K
for r in [56, 57, 67, 68, 69, 70, 71, 72, 75, 81, 90, 93]:
    lbl = ws[f'A{r}'].value
    vals = [f"{c}: {ws[f'{c}{r}'].value}" for c in cols[1:]]
    print(f"Row {r:2d} ({lbl}):")
    print("  " + " | ".join(vals))
