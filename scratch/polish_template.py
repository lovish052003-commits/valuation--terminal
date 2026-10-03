import openpyxl

wb = openpyxl.load_workbook('master_model_template.xlsx')
wb['DCF']['B37'] = 'Add: Cash & Liquid Investments'
wb['Intrinsic Valuation']['H38'] = "='Raw FS'!Z6-'Raw FS'!Z10"
wb['Intrinsic Valuation']['H44'] = "=-SUM('Raw FS'!Z27:Z28)-'Raw FS'!Z10"
wb.save('master_model_template.xlsx')
print('Template polished successfully!')
