import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

wb = openpyxl.load_workbook(r"C:\Users\LENOVO\Downloads\test\ADANIENT_Valuation_Model (8).xlsx", data_only=False)
wb_d = openpyxl.load_workbook(r"C:\Users\LENOVO\Downloads\test\ADANIENT_Valuation_Model (8).xlsx", data_only=True)

print("=== ADANIENT (8) DCF ROWS 30 TO 48 ===")
ws = wb['DCF']
ws_d = wb_d['DCF']
for r in range(30, 48):
    lbl = ws[f'B{r}'].value
    f = ws[f'D{r}'].value
    v = ws_d[f'D{r}'].value
    if lbl or f or v:
        print(f"Row {r:2d}: B='{lbl}' | D_formula='{f}' | D_eval={v}")

print("\n=== ADANIENT (8) COMP_VALUATION ROWS 22 TO 42 ===")
ws_c = wb['Comp_Valuation']
ws_cd = wb_d['Comp_Valuation']
for r in range(22, 43):
    lbl = ws_c[f'B{r}'].value
    print(f"Row {r:2d}: B='{lbl}'")
    for col in ['O', 'P', 'Q']:
        f = ws_c[f'{col}{r}'].value
        v = ws_cd[f'{col}{r}'].value
        print(f"    {col}: form='{f}' | eval={v}")

print("\n=== ADANIENT (8) WACC SHEET KEY CELLS ===")
ws_w = wb['WACC']
ws_wd = wb_d['WACC']
for r in range(1, 45):
    for col in ['B', 'C', 'D', 'E']:
        f = ws_w[f'{col}{r}'].value
        v = ws_wd[f'{col}{r}'].value
        if f is not None:
            print(f"  {col}{r}: form='{f}' | eval={v}")
