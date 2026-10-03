import openpyxl

wb = openpyxl.load_workbook(r'C:\Users\LENOVO\Downloads\test\ADANIENT_Valuation_Model (8).xlsx', data_only=False)
ws_wacc = wb['WACC']
for r in [26, 27, 28, 29, 33, 34, 35, 36, 37, 38, 40, 41, 43, 44, 46]:
    for col in ['C', 'E', 'J', 'K']:
        val = ws_wacc[f'{col}{r}'].value
        if val is not None:
            print(f"{col}{r}: {val}")
