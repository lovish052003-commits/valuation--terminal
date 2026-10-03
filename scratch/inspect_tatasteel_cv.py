import openpyxl

wb = openpyxl.load_workbook("Tata Steel Final Model.xlsx", data_only=False)
ws = wb["Comp_Valuation"]
for r in [10, 12, 23, 24, 25, 26, 27, 28, 30, 32, 33, 34, 35, 37, 39]:
    vals = [f"{col}: {ws[col + str(r)].value}" for col in ["I", "J", "O", "P", "Q"]]
    print(f"Row {r:2d}: " + " | ".join(vals))
