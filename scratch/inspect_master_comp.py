import openpyxl

wb_m = openpyxl.load_workbook('master_model_template.xlsx', data_only=False)
ws_m = wb_m['Comp_Valuation']

for r in range(1, 42):
    row_vals = []
    for c in range(1, 19):
        col_letter = openpyxl.utils.get_column_letter(c)
        val = ws_m[f'{col_letter}{r}'].value
        if val is not None:
            row_vals.append(f"{col_letter}{r}: {val}")
    if row_vals:
        print(f"--- ROW {r} ---")
        print(" | ".join(row_vals))
