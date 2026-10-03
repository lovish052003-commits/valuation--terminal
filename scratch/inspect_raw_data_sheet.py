import openpyxl

for fname in ['ITC Model.xlsx', 'exports/TATASTEEL_Valuation_Model.xlsx']:
    wb = openpyxl.load_workbook(fname, data_only=True)
    if 'Raw Data' in wb.sheetnames:
        ws = wb['Raw Data']
        print(f"\n=== {fname} -> Raw Data ===")
        # Print key cells
        print("  A1:D5:", [[ws.cell(r, c).value for c in range(1, 6)] for r in range(1, 6)])
        print("  Rows 20-30 (Cols A-E):", [[ws.cell(r, c).value for c in range(1, 6)] for r in range(20, 31)])
        print("  Cols G, H, J (Row 5-10):", [[ws.cell(r, c).value for c in [7, 8, 10]] for r in range(5, 11)])
    if 'List of Stocks' in wb.sheetnames:
        ws_list = wb['List of Stocks']
        print(f"=== {fname} -> List of Stocks ===")
        print("  A1:C5:", [[ws_list.cell(r, c).value for c in range(1, 4)] for r in range(1, 6)])
