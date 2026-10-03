"""
fix_jswsteel.py
Fixes three specific mechanical and logic errors in JSWSTEEL_Valuation_Model.xlsx:
1. Historical Tax Data (Data Sheet, Profit & Loss, Quarters):
   - Sets Tax to flat 25% of 'Profit before tax' for all historical periods.
   - Updates Net Profit = Profit before tax - Tax.
2. Peer Betas (WACC Sheet):
   - Updates peer Levered Betas (Rows 14 to 18) with realistic Indian steel market betas:
     Tata Steel: 1.45, JSW Steel: 1.25, Jindal Steel: 1.50, SAIL: 1.60, Jindal Stainless: 1.35.
   - Ensures Unlevered Beta formulas dynamically calculate based on these inputs.
3. Remove Arbitrary Reinvestment Caps (DCF Sheet):
   - Overwrites Reinvestment Rate (Row 11) to remove MIN and MAX functions.
   - Updates Free Cash Flow to Firm (Row 12) so unconstrained rates flow through.

Saves the corrected file as 'JSWSTEEL_Valuation_Model_Fixed.xlsx'.
"""

import os
import re
import openpyxl

def fix_jswsteel_valuation_model(input_path, output_path):
    print(f"Loading workbook: {input_path}")
    wb = openpyxl.load_workbook(input_path, data_only=False)

    # ---------------------------------------------------------
    # 1. Fix Historical Tax Data (Data Sheet & Profit & Loss)
    # ---------------------------------------------------------
    print("\n--- 1. Fixing Historical Tax Data ---")
    
    # 1a. Data Sheet (Annual: Rows 28-30, Quarterly: Rows 47-49)
    if 'Data Sheet' in wb.sheetnames:
        ws_data = wb['Data Sheet']
        
        # Annual Section (PBT=28, Tax=29, Net Profit=30)
        annual_pbt_row = 28
        annual_tax_row = 29
        annual_np_row = 30
        
        # Verify row labels
        for r in range(25, 35):
            lbl = str(ws_data.cell(r, 1).value or '').strip().lower()
            if 'profit before tax' in lbl:
                annual_pbt_row = r
            elif lbl == 'tax':
                annual_tax_row = r
            elif 'net profit' in lbl:
                annual_np_row = r
                
        print(f"Data Sheet (Annual): PBT Row={annual_pbt_row}, Tax Row={annual_tax_row}, Net Profit Row={annual_np_row}")
        
        # Apply 25% Tax and updated Net Profit across historical years (Cols B to K)
        for col_idx in range(2, 12):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            pbt_val = ws_data.cell(annual_pbt_row, col_idx).value
            try:
                pbt_num = float(pbt_val) if pbt_val is not None else 0.0
            except (ValueError, TypeError):
                pbt_num = 0.0
                
            tax_val = round(pbt_num * 0.25, 2)
            np_val = round(pbt_num - tax_val, 2)
            
            ws_data.cell(annual_tax_row, col_idx).value = tax_val
            ws_data.cell(annual_np_row, col_idx).value = np_val
            print(f"  Annual Col {col_letter}: PBT={pbt_num:,.1f} -> Tax (25%)={tax_val:,.1f} | Net Profit={np_val:,.1f}")

        # Quarterly Section (PBT=47, Tax=48, Net Profit=49)
        qtr_pbt_row = 47
        qtr_tax_row = 48
        qtr_np_row = 49
        for r in range(40, 55):
            lbl = str(ws_data.cell(r, 1).value or '').strip().lower()
            if 'profit before tax' in lbl:
                qtr_pbt_row = r
            elif lbl == 'tax':
                qtr_tax_row = r
            elif 'net profit' in lbl:
                qtr_np_row = r
                
        print(f"Data Sheet (Quarterly): PBT Row={qtr_pbt_row}, Tax Row={qtr_tax_row}, Net Profit Row={qtr_np_row}")
        for col_idx in range(2, 12):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            pbt_val = ws_data.cell(qtr_pbt_row, col_idx).value
            try:
                pbt_num = float(pbt_val) if pbt_val is not None else 0.0
            except (ValueError, TypeError):
                pbt_num = 0.0
                
            tax_val = round(pbt_num * 0.25, 2)
            np_val = round(pbt_num - tax_val, 2)
            ws_data.cell(qtr_tax_row, col_idx).value = tax_val
            ws_data.cell(qtr_np_row, col_idx).value = np_val

    # 1b. Profit & Loss Sheet
    if 'Profit & Loss' in wb.sheetnames:
        ws_pl = wb['Profit & Loss']
        pbt_row_pl = 10
        tax_row_pl = 11
        np_row_pl = 12
        for r in range(8, 15):
            lbl = str(ws_pl.cell(r, 1).value or '').strip().lower()
            if 'profit before tax' in lbl:
                pbt_row_pl = r
            elif lbl == 'tax':
                tax_row_pl = r
            elif 'net profit' in lbl:
                np_row_pl = r
                
        print(f"Profit & Loss Sheet: PBT Row={pbt_row_pl}, Tax Row={tax_row_pl}, Net Profit Row={np_row_pl}")
        for col_idx in range(2, 12):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            ws_pl.cell(tax_row_pl, col_idx).value = f"='Data Sheet'!{col_letter}{annual_tax_row}"
            ws_pl.cell(np_row_pl, col_idx).value = f"='Data Sheet'!{col_letter}{annual_np_row}"

    # 1c. Quarters Sheet
    if 'Quarters' in wb.sheetnames:
        ws_q = wb['Quarters']
        for col_idx in range(2, 12):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            ws_q.cell(11, col_idx).value = f"='Data Sheet'!{col_letter}{qtr_tax_row}"
            ws_q.cell(12, col_idx).value = f"='Data Sheet'!{col_letter}{qtr_np_row}"

    # ---------------------------------------------------------
    # 2. Replace Hardcoded Peer Betas (WACC Sheet)
    # ---------------------------------------------------------
    print("\n--- 2. Replacing Hardcoded Peer Betas in WACC Sheet ---")
    if 'WACC' in wb.sheetnames:
        ws_wacc = wb['WACC']
        
        # Exact market levered betas specified by user
        betas = [
            ("Tata Steel", 1.45),
            ("JSW Steel", 1.25),
            ("Jindal Steel", 1.50),
            ("SAIL", 1.60),
            ("Jindal Stainless", 1.35)
        ]
        
        for idx, (peer_label, beta_val) in enumerate(betas):
            r = 14 + idx  # Rows 14 to 18
            # Update Levered Beta (Column J)
            ws_wacc.cell(r, 10).value = beta_val
            # Ensure Unlevered Beta formula (Column K) is dynamic:
            # = J / (1 + (1 - TaxRate) * (Debt/Equity))
            ws_wacc.cell(r, 11).value = f"=J{r}/(1+(1-G{r})*H{r})"
            peer_curr = ws_wacc.cell(r, 2).value
            print(f"  Row {r}: Peer '{peer_label}' (Ref: {peer_curr}) -> Levered Beta = {beta_val} | Unlevered Formula = '{ws_wacc.cell(r, 11).value}'")

    # ---------------------------------------------------------
    # 3. Remove Arbitrary Reinvestment Caps (DCF Sheet)
    # ---------------------------------------------------------
    print("\n--- 3. Removing Arbitrary Reinvestment Caps in DCF Sheet ---")
    if 'DCF' in wb.sheetnames:
        ws_dcf = wb['DCF']
        
        # Row 11: Reinvestment Rate
        # Cell I11 currently has: =MIN(0.85, MAX(0.15, 'Intrinsic Valuation'!$L$55))
        # Remove MIN and MAX functions so the actual rate flows through:
        ws_dcf.cell(11, 9).value = "='Intrinsic Valuation'!$L$55"
        print(f"  DCF Cell I11 Reinvestment Rate updated to: {ws_dcf.cell(11, 9).value}")
        
        # Row 12: Free Cash Flow to Firm (FCFF)
        # Overwrite formulas in Columns I to M to remove MIN(0.85, ...) caps
        for col_idx in range(9, 14): # Columns I, J, K, L, M
            col_l = openpyxl.utils.get_column_letter(col_idx)
            # True FCFF formula: =EBIT(1-T) * (1 - Reinvestment Rate)
            ws_dcf.cell(12, col_idx).value = f"={col_l}10*(1-{col_l}11)"
            print(f"  DCF Cell {col_l}12 FCFF formula updated to: {ws_dcf.cell(12, col_idx).value}")

    # ---------------------------------------------------------
    # Save Output
    # ---------------------------------------------------------
    print(f"\nSaving corrected workbook to: {output_path}")
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    print("Workbook successfully saved!")

if __name__ == '__main__':
    candidate_paths = [
        r"C:\Users\LENOVO\Downloads\Advance Financial Project\exports\JSWSTEEL_Valuation_Model.xlsx",
        r"C:\Users\LENOVO\Downloads\test\JSWSTEEL_Valuation_Model.xlsx",
        r"C:\Users\LENOVO\Downloads\JSWSTEEL_Valuation_Model.xlsx",
        "JSWSTEEL_Valuation_Model.xlsx"
    ]
    
    in_file = None
    for p in candidate_paths:
        if os.path.exists(p):
            in_file = p
            break
            
    if not in_file:
        raise FileNotFoundError("Could not locate JSWSTEEL_Valuation_Model.xlsx")
        
    out_file = os.path.join(os.path.dirname(in_file), "JSWSTEEL_Valuation_Model_Fixed.xlsx")
    fix_jswsteel_valuation_model(in_file, out_file)
    
    # Save a copy directly in the workspace root as well
    root_out = os.path.join(os.getcwd(), "JSWSTEEL_Valuation_Model_Fixed.xlsx")
    if os.path.abspath(root_out) != os.path.abspath(out_file):
        import shutil
        shutil.copy2(out_file, root_out)
        print(f"Also copied to: {root_out}")
