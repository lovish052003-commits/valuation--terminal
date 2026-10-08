import openpyxl

wb_f = openpyxl.load_workbook("FORCEMOT_Valuation_Model_FIXED.xlsx", data_only=False)
wb_o = openpyxl.load_workbook(r"exports\FORCEMOT_Valuation_Model.xlsx", data_only=False)

ws_f = wb_f['DCF']
ws_o = wb_o['DCF']

print("=== DCF D35:D50 IN FIXED ===")
for r in range(35, 50):
    print(f"Row {r}: B={ws_f.cell(r, 2).value!r} | C={ws_f.cell(r, 3).value!r} | D={ws_f.cell(r, 4).value!r}")

print("\n=== DCF D35:D50 IN ORIG ===")
for r in range(35, 50):
    print(f"Row {r}: B={ws_o.cell(r, 2).value!r} | C={ws_o.cell(r, 3).value!r} | D={ws_o.cell(r, 4).value!r}")
