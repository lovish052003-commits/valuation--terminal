import openpyxl

wb_t = openpyxl.load_workbook('Tata Steel Final Model.xlsx', data_only=False)
ws_t = wb_t['Comp_Valuation']

wb_m = openpyxl.load_workbook('master_model_template.xlsx', data_only=False)
ws_m = wb_m['Comp_Valuation']

print("Tata B1:", ws_t['B1'].value)
print("Tata B7:", ws_t['B7'].value)
print("Tata B30:", ws_t['B30'].value)

print("Master B1:", ws_m['B1'].value)
print("Master B7:", ws_m['B7'].value)
print("Master B30:", ws_m['B30'].value)
