import openpyxl

wb_t = openpyxl.load_workbook('Tata Steel Final Model.xlsx', data_only=False)
ws_t = wb_t['Comp_Valuation']

for col_idx in range(1, 20):
    col_letter = openpyxl.utils.get_column_letter(col_idx)
    cd = ws_t.column_dimensions.get(col_letter)
    width = cd.width if cd else None
    print(f"Col {col_letter}: width = {width}")

print("\n--- CELL STYLES IN ROW 10, 12, 23, 30, 32, 37, 39 ---")
for r in [10, 12, 23, 30, 32, 37, 39]:
    for col in ['B', 'D', 'H', 'I', 'K', 'O', 'Q']:
        cell = ws_t[f'{col}{r}']
        print(f"{col}{r}: val={cell.value}, num_fmt={cell.number_format}, fill={cell.fill.fill_type if cell.fill else None}, font_bold={cell.font.bold if cell.font else None}")
