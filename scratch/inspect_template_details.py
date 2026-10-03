import openpyxl
from openpyxl.utils import get_column_letter

wb = openpyxl.load_workbook('master_model_template.xlsx', data_only=False)

print('=== DATA SHEET ROWS 65-75 ===')
ws_ds = wb['Data Sheet']
for r in range(65, 75):
    print(f"Row {r}: A={ws_ds.cell(r, 1).value} | B={ws_ds.cell(r, 2).value} | K={ws_ds.cell(r, 11).value}")

print('\n=== DATA SHEET ROW 93 ===')
for c in range(1, 12):
    print(f"Col {get_column_letter(c)}: {ws_ds.cell(93, c).value}", end=" | ")
print()

print('\n=== ALTMAN Z ROWS 70-80 ===')
ws_az = wb["Altman's Z Score"]
for r in range(70, 80):
    vals = [f"{get_column_letter(c)}={ws_az.cell(r, c).value}" for c in range(1, 8)]
    print(f"Row {r}: " + " | ".join(vals))

print('\n=== DCF ROW 6 DATES ===')
ws_dcf = wb['DCF']
for c in range(2, 14):
    print(f"{get_column_letter(c)}6: {ws_dcf.cell(6, c).value}", end=" | ")
print()

print('\n=== RAW DATA PEERS (Rows 1-20, Cols A-F and AZ-BB) ===')
ws_rd = wb['Raw Data']
for r in range(1, 20):
    print(f"Row {r}: A={ws_rd.cell(r, 1).value} | B={ws_rd.cell(r, 2).value} | C={ws_rd.cell(r, 3).value} | AZ={ws_rd.cell(r, 52).value} | BA={ws_rd.cell(r, 53).value}")

print('\n=== COMP VALUATION ===')
ws_comp = wb['Comp_Valuation']
for r in range(1, 20):
    row_vals = [f"{get_column_letter(c)}={ws_comp.cell(r, c).value}" for c in range(1, 18) if ws_comp.cell(r, c).value is not None]
    if row_vals:
        print(f"Row {r}: " + " | ".join(row_vals[:8]))
