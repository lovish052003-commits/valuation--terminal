import openpyxl
import re

wb = openpyxl.load_workbook('exports/ADANIPOWER_Valuation_Model.xlsx', data_only=False)
ws_comp = wb['Comp_Valuation']
for r in range(12, 22):
    val_b = ws_comp[f'B{r}'].value
    val_c = ws_comp[f'C{r}'].value
    print(f"Row {r}: B={val_b} | C={val_c}")
    if val_b and str(val_b).startswith('='):
        m = re.match(r"='?([^'!]+)'?!([A-Z0-9]+)", str(val_b))
        if m:
            sheet_n, cell_ref = m.groups()
            if sheet_n in wb.sheetnames:
                print(f"   -> Resolved from {sheet_n}!{cell_ref}: {wb[sheet_n][cell_ref].value}")
