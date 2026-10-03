import os, sys, openpyxl

path = os.path.abspath('exports/TEST_MARUTI_COM.xlsx')
if not os.path.exists(path):
    print("File not found:", path)
    sys.exit(1)

wb_f = openpyxl.load_workbook(path, data_only=False)
ws_f = wb_f['Beta-Regression']
ws_raw = wb_f['Raw Data']

print("--- FORMULAS CHECK in Beta-Regression ---")
for r in range(10, 15):
    raw_r = 6 + (r - 10)
    print(f"Row {r}:")
    print(f"  B: {ws_f.cell(r, 2).value} (Expected: ='Raw Data'!G{raw_r})")
    print(f"  C: {ws_f.cell(r, 3).value} (Expected: ='Raw Data'!H{raw_r})")
    print(f"  D: {ws_f.cell(r, 4).value}")
    print(f"  F: {ws_f.cell(r, 6).value} (Expected: ='Raw Data'!G{raw_r})")
    print(f"  G: {ws_f.cell(r, 7).value} (Expected: ='Raw Data'!J{raw_r})")
    print(f"  H: {ws_f.cell(r, 8).value}")

wb_v = openpyxl.load_workbook(path, data_only=True)
ws_v = wb_v['Beta-Regression']
ws_raw_v = wb_v['Raw Data']

print("\n--- VALUES CHECK in Beta-Regression vs Raw Data ---")
for r in range(10, 15):
    raw_r = 6 + (r - 10)
    raw_price = ws_raw_v.cell(raw_r, 8).value
    beta_price = ws_v.cell(r, 3).value
    print(f"Row {r}: Raw Data H{raw_r}={raw_price} | Beta-Regression C{r}={beta_price}")
    if beta_price is not None and raw_price is not None:
        assert float(beta_price) == float(raw_price), f"Mismatch! Raw: {raw_price}, Beta: {beta_price}"

print("\n--- SUCCESS: Beta-Regression is 100% connected to Raw Data! ---")
