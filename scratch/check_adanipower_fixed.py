import openpyxl

wb = openpyxl.load_workbook('exports/ADANIPOWER_Valuation_Model.xlsx', data_only=False)
print('--- DATA SHEET ---')
ws_ds = wb['Data Sheet']
print('K61 Total Liab:', ws_ds['K61'].value)
print('K66 Total Assets:', ws_ds['K66'].value)
print('K67 Rec:', ws_ds['K67'].value)
print('K68 Inv:', ws_ds['K68'].value)
print('K69 Cash:', ws_ds['K69'].value)

print('\n--- FORECASTING ---')
ws_fc = wb['Forecasting']
print('C14:', ws_fc['C14'].value)
print('C15:', ws_fc['C15'].value)

print('\n--- ALTMAN Z ---')
ws_alt = wb["Altman's Z Score"]
print('C61:', ws_alt['C61'].value)
print('I89:', ws_alt['I89'].value)

print('\n--- DUPONT ---')
ws_dup = wb['Dupont Analysis']
print('C72:', ws_dup['C72'].value)
print('I78:', ws_dup['I78'].value)

print('\n--- COMMON SIZE ---')
ws_cs = wb['Common Size Statement']
print('C7:', ws_cs['C7'].value)
print('C28:', ws_cs['C28'].value)

print('\n--- COMP VALUATION ---')
ws_cv = wb['Comp_Valuation']
for r in range(38, 43):
    print(f'Row {r}:', ws_cv.cell(row=r, column=2).value)

print('\n--- AI SUMMARY ---')
ws_ai = wb['AI Valuation Summary']
for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
    print(f'{col}5:', ws_ai[f'{col}5'].value)
