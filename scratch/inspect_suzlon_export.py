import openpyxl

wb = openpyxl.load_workbook('exports/SUZLON_Valuation_Model.xlsx', data_only=True)
ws_ds = wb['Data Sheet']
print("=== SUZLON EXPORTS/ DATA SHEET ===")
print("B1 (Company):", ws_ds['B1'].value)
print("B6 (Shares):", ws_ds['B6'].value)
print("B8 (Price):", ws_ds['B8'].value)
print("B9 (Market Cap):", ws_ds['B9'].value)
print("B17 (Sales FY17):", ws_ds['B17'].value)
print("K17 (Sales FY26):", ws_ds['K17'].value)

ws_ai = wb['AI Valuation Summary']
print("\n=== AI VALUATION SUMMARY ===")
print("B8 (WACC Label):", ws_ai['B8'].value, "| C8 (WACC Val):", ws_ai['C8'].value)
print("B37 (Intrinsic Value):", ws_ai['B37'].value, "| C37:", ws_ai['C37'].value)
print("B38 (CMP):", ws_ai['B38'].value, "| C38:", ws_ai['C38'].value)

ws_comp = wb['Comp_Valuation']
print("\n=== COMP_VALUATION ===")
print("B12 (Peer 1):", ws_comp['B12'].value)
print("C12 (Ticker 1):", ws_comp['C12'].value)
print("O25 (Median EV/Rev):", ws_comp['O25'].value)
print("P25 (Median EV/EBITDA):", ws_comp['P25'].value)
