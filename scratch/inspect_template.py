import openpyxl

wb = openpyxl.load_workbook('master_model_template.xlsx', data_only=False)
print('Sheet names:', wb.sheetnames)
if 'DCF' in wb.sheetnames:
    ws = wb['DCF']
    print('DCF D18:', ws['D18'].value)
    print('DCF D19:', ws['D19'].value)
    print('DCF D20:', ws['D20'].value)
    print('DCF D21:', ws['D21'].value)
    print('DCF I11:', ws['I11'].value)
if 'Intrinsic Valuation' in wb.sheetnames:
    iv = wb['Intrinsic Valuation']
    print('Intrinsic Valuation L55:', iv['L55'].value)
    for r in range(50, 78):
        row_vals = [f'{c}{r}: {iv[f"{c}{r}"].value}' for c in ['A', 'B', 'K', 'L'] if iv[f'{c}{r}'].value is not None]
        if row_vals:
            print(' | '.join(row_vals))
