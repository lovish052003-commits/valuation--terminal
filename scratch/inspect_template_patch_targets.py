import openpyxl

wb = openpyxl.load_workbook('master_model_template.xlsx', data_only=False)

print('=== DCF Sheet Cells ===')
ws_dcf = wb['DCF']
for coord in ['H6', 'I6', 'J6', 'K6', 'L6', 'M6', 'D18', 'D19', 'D20', 'D21', 'B22', 'D22', 'D25', 'D29', 'D33', 'D34', 'D35', 'D37', 'D38', 'D39', 'D40', 'D42', 'D44', 'B45', 'D45']:
    print(f"{coord}: {ws_dcf[coord].value}")

print('\n=== Intrinsic Valuation Rows 38 & 44 ===')
ws_iv = wb['Intrinsic Valuation']
for r in [38, 40, 44, 60, 62, 65]:
    for col_let in ['I', 'J', 'K', 'L']:
        print(f"{col_let}{r}={ws_iv[f'{col_let}{r}'].value}", end=' | ')
    print()

print('\n=== Raw Data T24 & WACC J16 ===')
ws_rd = wb['Raw Data']
print(f"Raw Data T24: {ws_rd['T24'].value}")
ws_wacc = wb['WACC']
print(f"WACC J16: {ws_wacc['J16'].value}")
print(f"WACC K16: {ws_wacc['K16'].value}")
print(f"WACC E27: {ws_wacc['E27'].value}")
print(f"WACC K26 (Rf): {ws_wacc['K26'].value}")
print(f"WACC K27 (ERP): {ws_wacc['K27'].value}")

print('\n=== Dupont Analysis B3 & Altman Z B3 ===')
ws_dup = wb['Dupont Analysis']
print(f"Dupont B3: {ws_dup['B3'].value}")
ws_alt = wb["Altman's Z Score"]
print(f"Altman B3: {ws_alt['B3'].value}")
