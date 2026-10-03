import os
import openpyxl

for fname in ['Tata Steel Ltd.xlsx', 'exports/TATASTEEL_Valuation_Model.xlsx']:
    if not os.path.exists(fname):
        continue
    wb = openpyxl.load_workbook(fname, data_only=False)
    dcf = wb['DCF']
    print(f"\n=== Checking {fname} DCF Row 11 ===")
    for col in ['H', 'I', 'J', 'K', 'L', 'M']:
        print(f"  {col}11: {dcf[f'{col}11'].value}")
    print(f"  D18: {dcf['D18'].value}")
    print(f"  D21: {dcf['D21'].value}")
    print(f"  B45: {dcf['B45'].value}")
    print(f"  D45: {dcf['D45'].value}")
    
    iv = wb['Intrinsic Valuation']
    print(f"\n=== Checking {fname} Intrinsic Valuation Rows 67-74 ===")
    for row in range(67, 75):
        print(f"  B{row}: {iv[f'B{row}'].value} | L{row}: {iv[f'L{row}'].value}")
