import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import openpyxl
from screener_client import fetch_company_data
from valuation_engine import calculate_valuation, compute_fundamental_reinvestment_engine
from excel_exporter import export_valuation_model

def test_tata_steel():
    print("Fetching TATASTEEL data...", flush=True)
    data = fetch_company_data('TATASTEEL')
    print("Calculating valuation...", flush=True)
    val = calculate_valuation(data)
    fund_res = compute_fundamental_reinvestment_engine(data)
    
    print("\n================ TATA STEEL RESULTS ================", flush=True)
    print(f"Target Company: {data.get('company_name')} (TATASTEEL)", flush=True)
    print(f"Historical Median Reinvestment: {fund_res['historical_median_reinvestment_rate']*100:.2f}%", flush=True)
    print(f"Normalized ROIC: {fund_res['normalized_roic']*100:.2f}%", flush=True)
    print(f"Expected Growth: {fund_res['expected_growth_rate']*100:.2f}%", flush=True)
    print(f"Growth Source: {fund_res['growth_source']}", flush=True)
    print(f"Fundamental Reinvestment: {fund_res['fundamental_reinvestment_rate']*100:.2f}%", flush=True)
    print(f"Terminal ROIC: {fund_res['terminal_roic']*100:.2f}%", flush=True)
    print(f"Terminal Reinvestment: {fund_res['terminal_reinvestment_rate']*100:.2f}%", flush=True)
    print(f"Year 1 DCF Reinvestment: {fund_res['forecast_reinvest_rates'][0]:.2f}%", flush=True)
    print(f"Year 5 DCF Reinvestment: {fund_res['forecast_reinvest_rates'][4]:.2f}%", flush=True)
    print(f"DCF Value/Share: Rs. {val.get('target_price', 0):.2f}", flush=True)
    print(f"Growth-ROIC Check: {fund_res['growth_roic_consistency']}", flush=True)
    print(f"Reinvestment Confidence: {fund_res['reinvestment_confidence']}", flush=True)
    print(f"Warnings: {fund_res['reinvestment_warnings']}", flush=True)
    
    print("\nExporting model to Excel...", flush=True)
    dest_file = export_valuation_model(data, val)
    print(f"Exported to: {dest_file}", flush=True)
    
    wb = openpyxl.load_workbook(dest_file, data_only=False)
    dcf = wb['DCF']
    iv = wb['Intrinsic Valuation']
    
    print("\n--- Excel Workbook Formulas Audit ---", flush=True)
    print(f"DCF!H11: {dcf['H11'].value}", flush=True)
    print(f"DCF!I11: {dcf['I11'].value}", flush=True)
    print(f"DCF!J11: {dcf['J11'].value}", flush=True)
    print(f"DCF!M11: {dcf['M11'].value}", flush=True)
    print(f"DCF!D18: {dcf['D18'].value}", flush=True)
    print(f"DCF!D21: {dcf['D21'].value}", flush=True)
    print(f"DCF!B45: {dcf['B45'].value}", flush=True)
    print(f"DCF!D45: {dcf['D45'].value}", flush=True)
    print(f"Intrinsic Valuation!L67: {iv['L67'].value}", flush=True)
    print(f"Intrinsic Valuation!L68: {iv['L68'].value}", flush=True)
    print(f"Intrinsic Valuation!L69: {iv['L69'].value}", flush=True)
    print(f"Intrinsic Valuation!L70: {iv['L70'].value}", flush=True)
    print(f"Intrinsic Valuation!L71: {iv['L71'].value}", flush=True)
    print(f"Intrinsic Valuation!L72: {iv['L72'].value}", flush=True)
    print(f"Intrinsic Valuation!L73: {iv['L73'].value}", flush=True)
    print(f"Intrinsic Valuation!L74: {iv['L74'].value}", flush=True)
    
    # Audit assertions
    assert dcf['H11'].value == "='Intrinsic Valuation'!$L$69", f"Expected H11 to be ='Intrinsic Valuation'!$L$69, got {dcf['H11'].value}"
    assert dcf['I11'].value == "='Intrinsic Valuation'!$L$69", f"Expected I11 to be ='Intrinsic Valuation'!$L$69, got {dcf['I11'].value}"
    assert dcf['D18'].value == "='Intrinsic Valuation'!$L$68", f"Expected D18 to be ='Intrinsic Valuation'!$L$68, got {dcf['D18'].value}"
    assert dcf['D45'].value == "=(D42-D44)/D44", f"Expected D45 to be =(D42-D44)/D44, got {dcf['D45'].value}"
    assert "0.0815" not in str(iv['L67'].value), f"Hardcoded 0.0815 in L67: {iv['L67'].value}"
    assert "MAX(DCF!D20" not in str(iv['L70'].value), f"Artificial WACC floor in L70: {iv['L70'].value}"
    assert iv['L74'].value != "HIGH", f"Expected confidence not to be HIGH for cyclical Tata Steel, got {iv['L74'].value}"
    
    # Check no forecast cells in DCF reference L55
    for col in ['H', 'I', 'J', 'K', 'L', 'M']:
        assert 'L55' not in str(dcf[f'{col}11'].value), f"DCF cell {col}11 still references L55: {dcf[f'{col}11'].value}"
        
    print("\nALL TATA STEEL AUDIT CHECKS PASSED PERFECTLY!", flush=True)

if __name__ == '__main__':
    test_tata_steel()
