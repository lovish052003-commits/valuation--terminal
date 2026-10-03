import openpyxl

wb_path = r'C:\Users\LENOVO\Downloads\test\ADANIPOWER_Valuation_Model (2).xlsx'
wb_v = openpyxl.load_workbook(wb_path, data_only=True)
wb_f = openpyxl.load_workbook(wb_path, data_only=False)

ds_v = wb_v['Data Sheet']
ds_f = wb_f['Data Sheet']

print('--- DATA SHEET ROWS 55 TO 75 ---')
for r in range(55, 75):
    k = ds_v.cell(row=r, column=1).value
    vals = [ds_v.cell(row=r, column=c).value for c in range(1, 10)]
    forms = [ds_f.cell(row=r, column=c).value for c in range(1, 10)]
    print(f'Row {r} | Label: {k}')
    print(f'   Vals: {vals[1:6]}')
    print(f'   Forms: {forms[1:6]}')
