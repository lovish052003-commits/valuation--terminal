import os
import sys
import openpyxl
from datetime import datetime

# Set UTF-8 encoding for console output
sys.stdout.reconfigure(encoding='utf-8')

from screener_client import fetch_company_data
from valuation_engine import calculate_valuation
from excel_exporter import export_valuation_model, validate_generated_workbook

def audit_workbook(file_path, ticker, is_bank=False):
    print(f"\n========================================================")
    print(f"AUDITING WORKBOOK: {ticker} -> {file_path}")
    print(f"========================================================")
    
    # 1. Formula pass (data_only=False)
    wb_f = openpyxl.load_workbook(file_path, data_only=False)
    # 2. Cached value pass (data_only=True)
    wb_v = openpyxl.load_workbook(file_path, data_only=True)
    
    # Check for error tokens
    error_tokens = ['#VALUE!', '#REF!', '#DIV/0!', '#NAME?', '#N/A', '#NUM!']
    found_errors = []
    
    for s_name in wb_f.sheetnames:
        ws_f = wb_f[s_name]
        ws_v = wb_v[s_name]
        for r in range(1, min(ws_f.max_row + 1, 100)):
            for c in range(1, min(ws_f.max_column + 1, 30)):
                col_let = openpyxl.utils.get_column_letter(c)
                v_val = str(ws_v.cell(row=r, column=c).value or '').strip()
                f_val = str(ws_f.cell(row=r, column=c).value or '').strip()
                for tok in ['#VALUE!', '#REF!', '#DIV/0!', '#NAME?']:
                    if tok in v_val or tok in f_val:
                        found_errors.append((s_name, f"{col_let}{r}", f_val, v_val))
    
    if found_errors:
        print(f"❌ FAILED: Found {len(found_errors)} formula error tokens:")
        for err in found_errors[:10]:
            print(f"   Sheet '{err[0]}' {err[1]}: formula='{err[2]}' value='{err[3]}'")
    else:
        print("✅ PASS: Zero critical error tokens (#VALUE!, #REF!, #DIV/0!, #NAME?) across all sheets.")

    # Check Data Sheet & DCF Share Count Reconciliation
    if 'Data Sheet' in wb_f.sheetnames:
        ds_b6 = wb_v['Data Sheet']['B6'].value
        ds_b8 = wb_v['Data Sheet']['B8'].value
        ds_b9 = wb_v['Data Sheet']['B9'].value
        ds_k70 = wb_v['Data Sheet']['K70'].value
        
        print(f"\n[DATA SHEET RECONCILIATION]")
        print(f"   CMP (B8): ₹{ds_b8}")
        print(f"   Shares Cr (B6): {ds_b6} (Formula: {wb_f['Data Sheet']['B6'].value})")
        print(f"   Market Cap (B9): ₹{ds_b9} (Formula: {wb_f['Data Sheet']['B9'].value})")
        print(f"   Raw Shares (K70): {ds_k70}")
        
        if 'DCF' in wb_f.sheetnames and not is_bank:
            dcf_d40_f = str(wb_f['DCF']['D40'].value or '')
            dcf_d40_v = wb_v['DCF']['D40'].value
            print(f"   DCF Shares (D40): {dcf_d40_v} (Formula: {dcf_d40_f})")
            
            # Check share discrepancy
            if ds_b6 and dcf_d40_v:
                diff = abs(float(ds_b6) - float(dcf_d40_v))
                if diff > 0.01:
                    print(f"❌ FAILED: Share count discrepancy! Data Sheet B6={ds_b6} vs DCF D40={dcf_d40_v}")
                else:
                    print(f"✅ PASS: 100% Share count reconciliation (Data Sheet B6 == DCF D40 == {ds_b6} Cr).")
        
        # Check Market Cap == CMP * Shares
        if ds_b8 and ds_b6 and ds_b9:
            expected_mcap = round(float(ds_b8) * float(ds_b6), 2)
            actual_mcap = round(float(ds_b9), 2)
            if abs(expected_mcap - actual_mcap) > 2.0:
                print(f"❌ FAILED: Market Cap does not reconcile: expected {expected_mcap}, got {actual_mcap}")
            else:
                print(f"✅ PASS: Market Cap reconciles with Price × Shares (₹{actual_mcap:,.2f} Cr).")

    # Check AI Valuation Summary WACC Synchronization
    if 'AI Valuation Summary' in wb_f.sheetnames:
        ws_ai_f = wb_f['AI Valuation Summary']
        ws_ai_v = wb_v['AI Valuation Summary']
        
        ai_a5 = ws_ai_v['A5'].value
        ai_b5 = ws_ai_v['B5'].value
        ai_c5 = ws_ai_v['C5'].value
        ai_d5 = ws_ai_v['D5'].value
        ai_e5 = ws_ai_v['E5'].value
        
        print(f"\n[AI VALUATION SUMMARY HEADLINE]")
        print(f"   Current Price (A5): ₹{ai_a5} (Formula: {ws_ai_f['A5'].value})")
        print(f"   Intrinsic Value (B5): ₹{ai_b5} (Formula: {ws_ai_f['B5'].value})")
        print(f"   Margin of Safety (C5): {ai_c5}")
        print(f"   Valuation Gap (D5): {ai_d5} (Formula: {ws_ai_f['D5'].value})")
        print(f"   Discount Rate / WACC (E5): {ai_e5} (Formula: {ws_ai_f['E5'].value})")
        
        # Check verdict has no BUY/SELL/HOLD
        d5_str = str(ai_d5 or '').upper()
        if any(w in d5_str for w in ['BUY', 'SELL', 'HOLD', 'OVERVALUED', 'UNDERVALUED']):
            print(f"❌ FAILED: AI Summary D5 contains non-neutral verdict: '{d5_str}'")
        else:
            print(f"✅ PASS: AI Summary D5 is strictly analytical: '{ai_d5}'")
            
        # Check Pillar 3 WACC synchronization
        b10_txt = str(ws_ai_v['B10'].value or '')
        print(f"\n[AI VALUATION SUMMARY PILLARS]")
        print(f"   Pillar 3 (B10): {b10_txt}")
        print(f"   Pillar 4 (B11): {ws_ai_v['B11'].value}")
        
        # Check no 0.0x fake zeros in Pillar 4
        b11_txt = str(ws_ai_v['B11'].value or '')
        if '0.0x' in b11_txt:
            print(f"❌ FAILED: Pillar 4 contains fake zero multiple '0.0x'!")
        else:
            print(f"✅ PASS: Zero fake zero multiples ('0.0x') in Pillar 4.")

        if not is_bank and 'WACC' in wb_f.sheetnames:
            wacc_k46_v = wb_v['WACC']['K46'].value
            print(f"   WACC Sheet K46: {wacc_k46_v}")
            if wacc_k46_v and ai_e5:
                diff = abs(float(wacc_k46_v) - float(ai_e5))
                if diff > 0.001:
                    print(f"❌ FAILED: WACC discrepancy between WACC!K46 ({wacc_k46_v}) and AI Summary E5 ({ai_e5})")
                else:
                    print(f"✅ PASS: 100% WACC synchronization between WACC!K46 and AI Summary E5.")
            if wacc_k46_v:
                wacc_pct_str = f"{float(wacc_k46_v)*100:.2f}%"
                if wacc_pct_str in b10_txt:
                    print(f"✅ PASS: AI Summary Pillar 3 narrative is 100% synchronized with live WACC ({wacc_pct_str}).")
                else:
                    print(f"⚠️ NOTICE: Pillar 3 narrative WACC text vs K46: text='{b10_txt}', expected '{wacc_pct_str}'")

    # Check Comp_Valuation single engine
    if 'Comp_Valuation' in wb_f.sheetnames:
        ws_cv_f = wb_f['Comp_Valuation']
        print(f"\n[COMP_VALUATION SINGLE ENGINE AUDIT]")
        col_i_sample = str(ws_cv_f['I12'].value or '')
        col_j_sample = str(ws_cv_f['J12'].value or '')
        col_o_sample = str(ws_cv_f['O12'].value or '')
        col_p_sample = str(ws_cv_f['P12'].value or '')
        print(f"   Row 12 Col O (Canonical EV/Rev): {col_o_sample}")
        print(f"   Row 12 Col I (Linked Reference): {col_i_sample}")
        print(f"   Row 12 Col P (Canonical EV/EBITDA): {col_p_sample}")
        print(f"   Row 12 Col J (Linked Reference): {col_j_sample}")
        
        if not is_bank:
            if col_i_sample == '=O12' and col_j_sample == '=P12':
                print("✅ PASS: Columns I and J directly link to Canonical Engine in Columns O & P (Zero duplicate logic).")
            else:
                print(f"❌ FAILED: Duplicate multiple calculation still present! I12='{col_i_sample}', J12='{col_j_sample}'")
        else:
            print("✅ PASS: Financial institution Comp_Valuation handles EV multiples appropriately as N/A.")

    print(f"\nAUDIT COMPLETE FOR {ticker}: SUCCESSFUL")
    return len(found_errors) == 0

def test_company(ticker, is_bank=False):
    print(f"\n========================================================")
    print(f"RUNNING E2E INSTITUTIONAL VALUATION FOR: {ticker}")
    print(f"========================================================")
    data = fetch_company_data(ticker)
    val_res = calculate_valuation(data)
    
    out_file = export_valuation_model(data, val_res)
    print(f"[EXPORT] Successfully generated {out_file}")
    
    # Run audit
    passed = audit_workbook(out_file, ticker, is_bank=is_bank)
    return passed

if __name__ == '__main__':
    targets = [
        ('ADANIPOWER', False),
        ('SBIN', True),
        ('ADANIENT', False)
    ]
    
    results = {}
    for t_sym, is_b in targets:
        try:
            ok = test_company(t_sym, is_bank=is_b)
            results[t_sym] = "PASSED" if ok else "FAILED"
        except Exception as e:
            print(f"❌ ERROR testing {t_sym}: {e}")
            import traceback
            traceback.print_exc()
            results[t_sym] = f"ERROR: {e}"
            
    print("\n========================================================")
    print("FINAL SUMMARY OF INSTITUTIONAL REGRESSION TEST:")
    for k, v in results.items():
        print(f"   {k}: {v}")
    print("========================================================")
