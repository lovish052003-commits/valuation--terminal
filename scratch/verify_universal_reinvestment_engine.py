import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import openpyxl
from screener_client import fetch_company_data
from valuation_engine import calculate_valuation, compute_fundamental_reinvestment_engine
from excel_exporter import export_valuation_model

def audit_workbook(xlsx_path, target_ticker, target_name):
    wb = openpyxl.load_workbook(xlsx_path, data_only=False)
    wb_val = openpyxl.load_workbook(xlsx_path, data_only=True)
    
    # 1. DCF Sheet checks
    dcf = wb['DCF']
    dcf_v = wb_val['DCF']
    iv = wb['Intrinsic Valuation']
    iv_v = wb_val['Intrinsic Valuation']
    
    h11_formula = str(dcf['H11'].value)
    i11_formula = str(dcf['I11'].value)
    j11_formula = str(dcf['J11'].value)
    m11_formula = str(dcf['M11'].value)
    d18_formula = str(dcf['D18'].value)
    d21_formula = str(dcf['D21'].value)
    d45_formula = str(dcf['D45'].value)
    b45_label = str(dcf['B45'].value)
    
    # Intrinsic Valuation checks
    l67_formula = str(iv['L67'].value)
    l68_val = iv['L68'].value
    l69_formula = str(iv['L69'].value)
    l70_formula = str(iv['L70'].value)
    l71_formula = str(iv['L71'].value)
    l72_formula = str(iv['L72'].value)
    l73_val = str(iv['L73'].value)
    l74_val = str(iv['L74'].value)
    
    # Check no forecast cells in DCF reference L55
    l55_in_dcf = []
    for col in ['H', 'I', 'J', 'K', 'L', 'M']:
        c_val = str(dcf[f'{col}11'].value)
        if 'L55' in c_val:
            l55_in_dcf.append(f'{col}11')
            
    # Check WACC peer set strictly excludes target
    wacc_target_found = False
    if 'WACC' in wb.sheetnames and 'Raw Data' in wb.sheetnames:
        raw = wb['Raw Data']
        for r in range(24, 29):
            p_name = str(raw.cell(r, 15).value or '').strip()
            if target_ticker.upper() in p_name.upper() or (len(target_name) > 4 and target_name.upper() in p_name.upper()):
                wacc_target_found = True

    return {
        'h11_formula': h11_formula,
        'i11_formula': i11_formula,
        'j11_formula': j11_formula,
        'm11_formula': m11_formula,
        'd18_formula': d18_formula,
        'd21_formula': d21_formula,
        'd45_formula': d45_formula,
        'b45_label': b45_label,
        'l67_formula': l67_formula,
        'l68_val': l68_val,
        'l69_formula': l69_formula,
        'l70_formula': l70_formula,
        'l71_formula': l71_formula,
        'l72_formula': l72_formula,
        'l73_growth_source': l73_val,
        'l74_confidence': l74_val,
        'l55_in_dcf': l55_in_dcf,
        'wacc_target_found': wacc_target_found,
        'dcf_value_share': dcf_v['D42'].value if dcf_v else None,
        'year_1_reinvest_val': dcf_v['I11'].value if dcf_v else None,
        'year_5_reinvest_val': dcf_v['M11'].value if dcf_v else None,
    }

def run_tests():
    print("=== RUNNING UNIVERSAL REINVESTMENT ENGINE AUDIT ===")
    tickers = [
        ('TATASTEEL', 'Tata Steel Ltd', 'cyclical industrial'),
        ('NESTLEIND', 'Nestle India Ltd', 'mature consumer'),
        ('INFY', 'Infosys Ltd', 'IT services'),
        ('SUNPHARMA', 'Sun Pharmaceutical Industries Ltd', 'pharma'),
        ('TRENT', 'Trent Ltd', 'high growth'),
        ('SUZLON', 'Suzlon Energy Ltd', 'negative / volatile ROIC turnaround'),
        ('SBIN', 'State Bank of India', 'financial institution')
    ]
    
    for ticker, name, archetype in tickers:
        print(f"\n--- Testing {ticker} ({name}) [{archetype}] ---")
        try:
            data = fetch_company_data(ticker)
            val = calculate_valuation(data)
            fund_res = compute_fundamental_reinvestment_engine(data)
            
            print(f"Target Company: {name} ({ticker})")
            print(f"Historical Median Reinvestment: {fund_res['historical_median_reinvestment_rate']*100:.2f}%" if fund_res['historical_median_reinvestment_rate'] else "Historical Median Reinvestment: N/A")
            print(f"Normalized ROIC: {fund_res['normalized_roic']*100:.2f}%")
            print(f"Expected Growth: {fund_res['expected_growth_rate']*100:.2f}% (Source: {fund_res['growth_source']})")
            print(f"Fundamental Reinvestment: {fund_res['fundamental_reinvestment_rate']*100:.2f}%")
            print(f"Terminal ROIC: {fund_res['terminal_roic']*100:.2f}%")
            print(f"Terminal Reinvestment: {fund_res['terminal_reinvestment_rate']*100:.2f}%")
            print(f"Year 1 DCF Reinvestment: {fund_res['forecast_reinvest_rates'][0]:.2f}%")
            print(f"Year 5 DCF Reinvestment: {fund_res['forecast_reinvest_rates'][4]:.2f}%")
            print(f"DCF Value/Share: Rs. {val.get('target_price', 0):.2f}")
            print(f"Growth-ROIC Check: {fund_res['growth_roic_consistency']}")
            print(f"Reinvestment Confidence: {fund_res['reinvestment_confidence']}")
            print(f"Warnings: {fund_res['reinvestment_warnings']}")
            
            # Export to Excel and audit formulas
            dest_file = export_valuation_model(data, val)
            audit = audit_workbook(dest_file, ticker, name)
            
            print(f"Excel DCF!H11 Formula: {audit['h11_formula']}")
            print(f"Excel DCF!I11 Formula: {audit['i11_formula']}")
            print(f"Excel DCF!D18 Formula: {audit['d18_formula']}")
            print(f"Excel DCF!D45 Formula: {audit['d45_formula']} (Label: {audit['b45_label']})")
            print(f"Excel Intrinsic Valuation!L67: {audit['l67_formula']}")
            print(f"Excel Intrinsic Valuation!L70: {audit['l70_formula']}")
            print(f"Excel DCF L55 references: {audit['l55_in_dcf']}")
            print(f"WACC Target in Peers: {audit['wacc_target_found']}")
            
            # Assertions
            assert audit['h11_formula'] == "='Intrinsic Valuation'!$L$69", f"H11 failure: {audit['h11_formula']}"
            assert audit['i11_formula'] == "='Intrinsic Valuation'!$L$69", f"I11 failure: {audit['i11_formula']}"
            assert audit['d18_formula'] == "='Intrinsic Valuation'!$L$68", f"D18 failure: {audit['d18_formula']}"
            assert audit['d45_formula'] == "=(D42-D44)/D44", f"D45 failure: {audit['d45_formula']}"
            assert "0.0815" not in str(audit['l67_formula']), "Hardcoded 0.0815 found in L67!"
            assert "MAX(DCF!D20" not in str(audit['l70_formula']), "Artificial WACC floor found in L70!"
            assert len(audit['l55_in_dcf']) == 0, f"Obsolete L55 found in DCF forecast: {audit['l55_in_dcf']}"
            assert not audit['wacc_target_found'], f"Target found in WACC peer set!"
            print(f"PASSED all 10 universal criteria for {ticker}!")
            
        except Exception as e:
            print(f"ERROR testing {ticker}: {e}")
            import traceback
            traceback.print_exc()

if __name__ == '__main__':
    run_tests()
