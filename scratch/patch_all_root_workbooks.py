import openpyxl
import glob
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from excel_exporter import strip_calc_chain_from_xlsx

files = [f for f in glob.glob('*.xlsx') if not f.startswith('~$')]
print(f"Found {len(files)} files to examine.")

for f in files:
    try:
        wb = openpyxl.load_workbook(f, data_only=False)
        changed = False
        sheets = wb.sheetnames
        
        # 1. WACC!E26
        if 'WACC' in sheets:
            ws_wacc = wb['WACC']
            target_formula = "='Data Sheet'!K27/AVERAGE('Data Sheet'!J59:K59)"
            if ws_wacc['E26'].value != target_formula:
                ws_wacc['E26'].value = target_formula
                changed = True
                
        # 2. DCF!D22
        if 'DCF' in sheets:
            ws_dcf = wb['DCF']
            target_d22 = "=MAX('Intrinsic Valuation'!L40,D20)"
            if ws_dcf['D22'].value != target_d22:
                ws_dcf['D22'].value = target_d22
                changed = True
                
        # 3. Raw Data!W12:W21
        if 'Raw Data' in sheets:
            ws_raw = wb['Raw Data']
            for r in range(12, 22):
                target_w = f'=IF(X{r}>0,U{r}/X{r},"N/A")'
                if ws_raw[f'W{r}'].value != target_w:
                    ws_raw[f'W{r}'].value = target_w
                    changed = True
                    
        # Apply Control & Checks and Institutional Wiring
        from institutional_control_checks import apply_institutional_wiring
        apply_institutional_wiring(wb)
        wb.save(f)
        wb.close()
        strip_calc_chain_from_xlsx(f)
        print(f"Patched: {f}")
    except Exception as e:
        print(f"Skipping {f} due to: {e}")

print("All workbooks checked and synchronized!")
