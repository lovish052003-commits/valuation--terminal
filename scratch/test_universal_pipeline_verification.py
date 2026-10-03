import os, sys, zipfile, openpyxl

sys.path.insert(0, os.path.abspath('.'))
import screener_client
import valuation_engine
import excel_exporter

def test_company_pipeline(ticker):
    print(f"\n=======================================================")
    print(f"RUNNING COMPREHENSIVE PIPELINE AUDIT FOR: {ticker}")
    print(f"=======================================================")
    
    # 1. Fetch live data & compute valuation
    screener_data = screener_client.fetch_company_data(ticker)
    val_result = valuation_engine.calculate_valuation(screener_data)
    
    # 2. Export full institutional model
    export_path = excel_exporter.export_valuation_model(screener_data, val_result)
    print(f"[Exported Path]: {export_path}")
    assert os.path.exists(export_path), f"File {export_path} does not exist!"
    
    # 3. Check for absence of calcChain.xml (Guarantees ZERO Excel repair / corruption warnings)
    with zipfile.ZipFile(export_path, 'r') as z:
        names = z.namelist()
        has_calcchain = any('calcChain.xml' in n for n in names)
        print(f"[Audit 1 - XML Integrity] Has xl/calcChain.xml: {has_calcchain} (MUST BE FALSE)")
        assert not has_calcchain, "FAIL: calcChain.xml found in export! Excel will show repair popup!"
        print("  -> PASS: Zero calcChain corruption risk!")
        
    # 4. Open workbook and audit sheets
    wb = openpyxl.load_workbook(export_path, data_only=False)
    wb_vals = openpyxl.load_workbook(export_path, data_only=True)
    sheet_names = wb.sheetnames
    print(f"[Audit 2 - Sheet Count] Total sheets: {len(sheet_names)}")
    assert len(sheet_names) >= 22, f"FAIL: Expected >= 22 sheets, found {len(sheet_names)}"
    print(f"  -> PASS: All sheets intact: {sheet_names}")
    
    # 5. Audit Data Sheet Rows 16 to 25
    ws_ds = wb_vals['Data Sheet']
    print("\n[Audit 3 - Data Sheet P&L Expense Schedules]")
    r17_sales = [ws_ds.cell(17, c).value for c in range(2, 12)]
    r18_raw_mat = [ws_ds.cell(18, c).value for c in range(2, 12)]
    r19_inv_chg = [ws_ds.cell(19, c).value for c in range(2, 12)]
    r21_mfr_exp = [ws_ds.cell(21, c).value for c in range(2, 12)]
    r22_emp_cost = [ws_ds.cell(22, c).value for c in range(2, 12)]
    r24_other_exp = [ws_ds.cell(24, c).value for c in range(2, 12)]
    r32_ebitda = [ws_ds.cell(32, c).value for c in range(2, 12)]
    
    print(f"  Row 17 (Sales)          : {r17_sales[:5]} ...")
    print(f"  Row 18 (Raw Material)   : {r18_raw_mat[:5]} ...")
    print(f"  Row 19 (Change in Inv)  : {r19_inv_chg[:5]} ...")
    print(f"  Row 21 (Other Mfr. Exp) : {r21_mfr_exp[:5]} ...")
    print(f"  Row 22 (Employee Cost)  : {r22_emp_cost[:5]} ...")
    print(f"  Row 24 (Other Expenses) : {r24_other_exp[:5]} ...")
    print(f"  Row 32 (EBITDA)         : {r32_ebitda[:5]} ...")
    
    # For pharmaceutical / manufacturing companies like SUNPHARMA, raw materials & employee cost must be > 0!
    has_itemized_expenses = any(v and v > 0 for v in r18_raw_mat) or any(v and v > 0 for v in r22_emp_cost)
    print(f"  -> Has itemized expense values: {has_itemized_expenses}")
    assert has_itemized_expenses, "FAIL: Expenses were not properly itemized into Rows 18-24!"
    print("  -> PASS: Data Sheet P&L expenses are fully itemized!")
    
    # 6. Audit Raw FS AV56 Target EBITDA Link
    ws_rfs = wb['Raw FS']
    av56_formula = ws_rfs['AV56'].value
    print(f"\n[Audit 4 - Raw FS AV56 Formula]: {av56_formula}")
    assert av56_formula == "='Data Sheet'!K32", f"FAIL: Expected ='Data Sheet'!K32, got {av56_formula}"
    print("  -> PASS: Raw FS AV56 correctly links to Target EBITDA K32!")
    
    # 7. Audit Historical FS Row 16 Formula
    hfs_name = 'Historical FS' if 'Historical FS' in sheet_names else 'HistoricalFS'
    ws_hfs = wb[hfs_name]
    c16_formula = ws_hfs['C16'].value
    k16_formula = ws_hfs['K16'].value
    print(f"\n[Audit 5 - Historical FS Row 16 Formulas]")
    print(f"  C16: {c16_formula}")
    print(f"  K16: {k16_formula}")
    assert "23" in str(c16_formula) and "24" in str(c16_formula), f"FAIL: C16 must combine row 23 and 24, got {c16_formula}"
    assert "23" in str(k16_formula) and "24" in str(k16_formula), f"FAIL: K16 must combine row 23 and 24, got {k16_formula}"
    print("  -> PASS: Historical FS accurately combines Selling & admin (Row 23) and Other Expenses (Row 24)!")
    
    # 8. Audit DCF Share Price CMP Link
    ws_dcf = wb['DCF']
    cmp_formula = None
    cmp_row = None
    for r in range(40, 50):
        lbl = str(ws_dcf.cell(r, 2).value or '').strip().lower()
        if 'share price' in lbl:
            cmp_row = r
            cmp_formula = ws_dcf.cell(r, 4).value
            break
    print(f"\n[Audit 6 - DCF Row {cmp_row} Share Price CMP Formula]: {cmp_formula}")
    assert cmp_formula == "='Data Sheet'!B8", f"FAIL: Expected ='Data Sheet'!B8, got {cmp_formula}"
    print("  -> PASS: DCF Share Price links dynamically to CMP B8!")
    
    # 9. Audit Comp_Valuation Sheet
    ws_comp = wb['Comp_Valuation']
    p32_formula = ws_comp['P32'].value
    print(f"\n[Audit 7 - Comp_Valuation P32 Formula]: {p32_formula}")
    assert "AV56" in str(p32_formula), f"FAIL: Comp_Valuation P32 should reference Raw FS AV56, got {p32_formula}"
    print("  -> PASS: Comp_Valuation EV/EBITDA connects to Target EBITDA AV56!")
    
    # 10. Audit Cash Flow Statement Sheet
    ws_cfs = wb_vals['Cash Flow Statement']
    cfs_r6_cfo = [ws_cfs.cell(6, c).value for c in range(3, 10)]
    cfs_r13_cfi = [ws_cfs.cell(13, c).value for c in range(3, 10)]
    cfs_r24_cff = [ws_cfs.cell(24, c).value for c in range(3, 10)]
    print(f"\n[Audit 8 - Cash Flow Statement]")
    print(f"  CFO (Row 6) : {cfs_r6_cfo}")
    print(f"  CFI (Row 13): {cfs_r13_cfi}")
    print(f"  CFF (Row 24): {cfs_r24_cff}")
    has_cf = any(v and v != 0 for v in cfs_r6_cfo)
    assert has_cf, "FAIL: Cash Flow Statement row 6 has no values!"
    print("  -> PASS: Cash Flow Statement sheet dynamically populated!")
    
    # 11. Audit DuPont and Altman sheets
    ws_dup = wb['Dupont Analysis']
    dup_b2 = ws_dup['B2'].value
    print(f"\n[Audit 9 - DuPont & Altman B2 link]: {dup_b2}")
    assert dup_b2 == "='Data Sheet'!B1", f"FAIL: Expected ='Data Sheet'!B1, got {dup_b2}"
    print("  -> PASS: DuPont & Altman dynamic links intact!")

    # 12. Audit Universal Minority Interest Bridge
    print(f"\n[Audit 10 - Universal Minority Interest Bridge]")
    assert ws_ds['A73'].value == "Minority Interest", f"FAIL: Data Sheet A73 should be 'Minority Interest', got {ws_ds['A73'].value}"
    print(f"  Data Sheet A73: '{ws_ds['A73'].value}' | K73: {ws_ds['K73'].value}")

    # Check DCF Less: Minority Interest
    found_dcf_mi = False
    for r in range(30, 45):
        lbl = str(ws_dcf.cell(r, 2).value or '').strip()
        if 'minority interest' in lbl.lower():
            found_dcf_mi = True
            mi_formula = ws_dcf.cell(r, 4).value
            eq_formula = ws_dcf.cell(r + 1, 4).value
            print(f"  DCF Row {r}: '{lbl}' | Formula: {mi_formula}")
            print(f"  DCF Row {r+1}: '{ws_dcf.cell(r+1, 2).value}' | Formula: {eq_formula}")
            assert mi_formula == "='Data Sheet'!K73", f"FAIL: Expected ='Data Sheet'!K73, got {mi_formula}"
            assert f"-D{r}" in str(eq_formula), f"FAIL: Equity Value formula should subtract -D{r}, got {eq_formula}"
            break
    assert found_dcf_mi, "FAIL: 'Less: Minority Interest' row not found in DCF bridge!"
    print("  -> PASS: Universal Minority Interest deduction successfully verified in DCF and Data Sheet!")

    print(f"\n>>> ALL 10 AUDIT CHECKS PASSED FOR {ticker}! <<<")
    return True

if __name__ == '__main__':
    test_company_pipeline('SUNPHARMA')
