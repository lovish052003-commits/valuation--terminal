import openpyxl

wb = openpyxl.load_workbook('ITC Model.xlsx', data_only=False)
ws = wb['Raw FS']
for r in [56, 57]:
    for col in ['L', 'M', 'N', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AZ', 'BA']:
        print(f'{col}{r}: {ws[f"{col}{r}"].value}', end=' | ')
    print()
