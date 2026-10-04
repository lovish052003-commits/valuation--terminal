import openpyxl

wb = openpyxl.load_workbook('master_model_template.xlsx', data_only=False)
ws_iv = wb['Intrinsic Valuation']
ws_dcf = wb['DCF']

print('--- Master Template Formulas ---')
print('H51:', ws_iv['H51'].value)
print('H52:', ws_iv['H52'].value)
print('I52:', ws_iv['I52'].value)
print('J52:', ws_iv['J52'].value)
print('K52:', ws_iv['K52'].value)
print('L52:', ws_iv['L52'].value)
print('H60:', ws_iv['H60'].value)
print('L60:', ws_iv['L60'].value)
print('H62:', ws_iv['H62'].value)
print('L62:', ws_iv['L62'].value)
print('DCF D18:', ws_dcf['D18'].value)

assert ws_iv['H52'].value == '=H51/H49', f"H52 is {ws_iv['H52'].value}"
assert ws_iv['I52'].value == '=I51/I49', f"I52 is {ws_iv['I52'].value}"
assert ws_iv['L52'].value == '=L51/L49', f"L52 is {ws_iv['L52'].value}"
assert ws_dcf['D18'].value == "='Intrinsic Valuation'!L62", f"D18 is {ws_dcf['D18'].value}"
print('\n[SUCCESS] Master template formulas verified: H52 = =H51/H49 and DCF!D18 = =\'Intrinsic Valuation\'!L62')
