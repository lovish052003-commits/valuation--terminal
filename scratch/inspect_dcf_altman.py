import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('master_model_template.xlsx', data_only=False)

print("=== DCF ROW 15-22 ===")
ws_dcf = wb['DCF']
for r in range(15, 23):
    print(f"DCF Row {r}: B={repr(ws_dcf[f'B{r}'].value)} | C={repr(ws_dcf[f'C{r}'].value)} | D={repr(ws_dcf[f'D{r}'].value)}")

print("\n=== INTRINSIC VALUATION ROW 60-68 ===")
ws_iv = wb['Intrinsic Valuation']
for r in range(60, 68):
    row_vals = [f"{c}{r}={repr(ws_iv[f'{c}{r}'].value)}" for c in ['A', 'B', 'K', 'L', 'M'] if ws_iv[f'{c}{r}'].value is not None]
    if row_vals:
        print(f"IV Row {r}: " + " | ".join(row_vals))

print("\n=== ALTMANS Z SCORE ROWS 55-85 ===")
ws_alt = wb["Altman's Z Score"]
for r in range(55, 85):
    row_vals = []
    for c in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
        val = ws_alt[f'{c}{r}'].value
        if val is not None:
            row_vals.append(f"{c}{r}={repr(val)}")
    if row_vals:
        print(f"Altman Row {r}: " + " | ".join(row_vals))

print("\n=== DATA SHEET ROWS 55-70 ===")
ws_ds = wb["Data Sheet"]
for r in [57, 58, 61, 69]:
    row_vals = []
    for c in ['A', 'B', 'G', 'H', 'I', 'J', 'K']:
        val = ws_ds[f'{c}{r}'].value
        if val is not None:
            row_vals.append(f"{c}{r}={repr(val)}")
    print(f"DS Row {r}: " + " | ".join(row_vals))

print("\n=== INTRINSIC VALUATION L13, L19 ===")
print("IV L13:", ws_iv['L13'].value)
print("IV L19:", ws_iv['L19'].value)
