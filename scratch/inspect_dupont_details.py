import openpyxl

wb = openpyxl.load_workbook('ITC Model.xlsx', data_only=False)

for sname in ['Dupont Analysis', "Altman's Z Score"]:
    ws = wb[sname]
    print(f"\n==================== Merged ranges in {sname} ====================")
    for rng in ws.merged_cells.ranges:
        coord = str(rng)
        if any(f"B{r}" in coord for r in [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44]):
            print(coord)

    print(f"\n--- Specific cell contents in {sname} ---")
    for r in [2, 3, 4, 5, 6, 8, 14, 15, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44]:
        val = ws.cell(r, 2).value
        if val is not None:
            print(f"Row {r:2d} (Col B): {repr(val)[:100]}")
