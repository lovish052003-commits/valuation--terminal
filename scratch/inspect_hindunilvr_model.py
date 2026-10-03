import openpyxl

wb = openpyxl.load_workbook(r'exports\HINDUNILVR_Valuation_Model.xlsx', data_only=False)

print("=== 1. DCF SHEET ===")
ws_dcf = wb['DCF']
print("B37:", ws_dcf['B37'].value, "| D37:", ws_dcf['D37'].value)
print("B38:", ws_dcf['B38'].value, "| D38:", ws_dcf['D38'].value)
print("B45:", ws_dcf['B45'].value, "| D45:", ws_dcf['D45'].value)
print("B22:", ws_dcf['B22'].value, "| D22:", ws_dcf['D22'].value)
print("D21:", ws_dcf['D21'].value)
print("D18:", ws_dcf['D18'].value)
print("Forecast Dates (H6:M6):", [ws_dcf.cell(row=6, column=c).value for c in range(8, 14)])

print("\n=== 2. DATA SHEET ===")
ws_data = wb['Data Sheet']
print("K69 (Cash):", ws_data['K69'].value)
print("K64 (Investments):", ws_data['K64'].value)
print("K93 (Shares):", ws_data['K93'].value)
print("B1 (Company):", ws_data['B1'].value)
print("B8 (Price):", ws_data['B8'].value)

print("\n=== 3. AI VALUATION SUMMARY ===")
ws_sum = wb['AI Valuation Summary']
for r in range(4, 6):
    print(f"Row {r}:", [ws_sum.cell(row=r, column=c).value for c in range(1, 9)])
print("Params (Rows 8 to 17, Col B, C, D):")
for r in range(8, 18):
    print(f"  {ws_sum.cell(row=r, column=2).value}: {ws_sum.cell(row=r, column=3).value} | Unit: {repr(str(ws_sum.cell(row=r, column=4).value).encode('ascii', 'replace').decode('ascii'))}")
print("Relative Multiples Section (Rows 47 to 51):")
for r in range(47, 52):
    lbl = ws_sum.cell(row=r, column=2).value
    if lbl:
        print(f"  {lbl}: {ws_sum.cell(row=r, column=3).value} | Unit: {repr(str(ws_sum.cell(row=r, column=4).value).encode('ascii', 'replace').decode('ascii'))}")

print("\n=== 4. DUPONT & ALTMAN ===")
print("Dupont Analysis B3:", wb['Dupont Analysis']['B3'].value)
print("Altman's Z Score B3:", wb["Altman's Z Score"]['B3'].value)

print("\n=== 5. WACC SHEET ===")
ws_wacc = wb['WACC']
print("WACC J16 (Beta):", ws_wacc['J16'].value)
print("WACC E27 (Tax Rate):", ws_wacc['E27'].value)
print("WACC Peer Betas (J14:J18):", [ws_wacc.cell(row=r, column=10).value for r in range(14, 19)])

print("\n=== 6. RAW DATA PEERS (Rows 12 to 17) ===")
ws_rd = wb['Raw Data']
print("Raw Data T24 (Tax Rate):", ws_rd['T24'].value)
for r in range(12, 17):
    print(f"Row {r}: Name={ws_rd.cell(row=r, column=12).value}, Price={ws_rd.cell(row=r, column=16).value}, Shares(Q)={ws_rd.cell(row=r, column=17).value}, Mcap(R)={ws_rd.cell(row=r, column=18).value}")
