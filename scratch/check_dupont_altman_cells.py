import openpyxl

wb = openpyxl.load_workbook('exports/ADANIENT_Valuation_Model.xlsx', data_only=True)
ws_dup = wb['Dupont Analysis']
ws_alt = wb["Altman's Z Score"]

print("--- Dupont Analysis all cells ---")
for r in range(1, 40):
    for c in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        v = ws_dup[f'{c}{r}'].value
        if v is not None and ('roe' in str(v).lower() or isinstance(v, (int, float))):
            print(f'Dupont {c}{r}: {v}')

print("\n--- Altman all cells ---")
for r in range(1, 40):
    for c in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        v = ws_alt[f'{c}{r}'].value
        if v is not None and ('altman' in str(v).lower() or 'score' in str(v).lower() or isinstance(v, (int, float))):
            print(f'Altman {c}{r}: {v}')
