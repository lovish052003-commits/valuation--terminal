import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
fresh_oil_path = r"C:\Users\LENOVO\Downloads\Advance Financial Project\exports\OIL_Valuation_Model.xlsx"

wb = openpyxl.load_workbook(fresh_oil_path, data_only=False)
wb_d = openpyxl.load_workbook(fresh_oil_path, data_only=True)

print("=== FRESH OIL_VALUATION_MODEL.XLSX VERIFICATION ===")

print("\n1. AI Valuation Summary Key Cells:")
ws_s = wb['AI Valuation Summary']
ws_sd = wb_d['AI Valuation Summary']
for coord in ['A4', 'B4', 'C4', 'D4', 'E4', 'F4', 'G4',
              'A5', 'B5', 'C5', 'D5', 'E5', 'F5', 'G5']:
    print(f"  {coord}: form='{ws_s[coord].value}' | eval={ws_sd[coord].value}")

print("\n2. DCF Sheet Bridge:")
ws_dcf = wb['DCF']
ws_dcfd = wb_d['DCF']
for r in [33, 34, 35, 37, 38, 39, 40, 42, 44, 45]:
    lbl = ws_dcf[f'B{r}'].value
    f = ws_dcf[f'D{r}'].value
    v = ws_dcfd[f'D{r}'].value
    print(f"  Row {r:2d} ({lbl}): form='{f}' | eval={v}")

print("\n3. Comp_Valuation Rows 30 to 42:")
ws_c = wb['Comp_Valuation']
ws_cd = wb_d['Comp_Valuation']
for r in range(30, 43):
    lbl = ws_c[f'B{r}'].value
    o = ws_c[f'O{r}'].value
    p = ws_c[f'P{r}'].value
    q = ws_c[f'Q{r}'].value
    if lbl or o or p or q:
        print(f"  Row {r:2d} ({lbl}): O='{o}' | P='{p}' | Q='{q}'")
    else:
        print(f"  Row {r:2d}: [Clean / Empty as expected in Gold Standard]")
