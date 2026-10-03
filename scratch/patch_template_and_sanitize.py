import openpyxl

wb = openpyxl.load_workbook('master_model_template.xlsx')

# Update DCF sheet formulas
ws_dcf = wb['DCF']
ws_dcf['I11'] = "=MIN(0.85, 'Intrinsic Valuation'!$L$55)"
ws_dcf['J11'] = "=MIN(0.85, $I$11+($M$11-$I$11)/4*1)"
ws_dcf['K11'] = "=MIN(0.85, $I$11+($M$11-$I$11)/4*2)"
ws_dcf['L11'] = "=MIN(0.85, $I$11+($M$11-$I$11)/4*3)"
ws_dcf['M11'] = "=MIN(0.85, D21)"

# Clear N24 in Raw FS
ws_raw = wb['Raw FS']
print('Old N24 value:', ws_raw['N24'].value)
ws_raw['N24'].value = 0.0
print('New N24 value:', ws_raw['N24'].value)

wb.save('master_model_template.xlsx')
wb.close()
print('master_model_template.xlsx successfully updated!')
