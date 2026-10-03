import os
import sys

project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

import screener_client
import valuation_engine
import excel_exporter

def test_company(symbol):
    print(f"==================================================")
    print(f"Testing valuation & Excel export for: {symbol}")
    print(f"==================================================")
    try:
        screener_data = screener_client.fetch_company_data(symbol)
        print(f"Found: {screener_data['company_name']} ({screener_data.get('ticker')})")
        print(f"Sector: {screener_data.get('sector')}, Price: {screener_data.get('current_price')}")
        
        val = valuation_engine.calculate_valuation(screener_data)
        print(f"Valuation math complete:")
        print(f"  Intrinsic Value: Rs. {val.get('intrinsic_value')}")
        print(f"  Current Price: Rs. {val.get('current_price')}")
        print(f"  WACC: {val.get('wacc')}")
        
        export_path = excel_exporter.export_valuation_model(screener_data, val)
        print(f"Exported workbook to: {export_path}")
        
        target_ticker = screener_data.get('ticker', '')
        target_name = screener_data.get('company_name', '')
        sec_k = excel_exporter.get_sector_key(screener_data)
        is_fin = excel_exporter.is_financial_sector(sec_k)
        
        is_valid = excel_exporter.validate_generated_workbook(export_path, target_ticker, target_name, is_financial=is_fin)
        print(f"\nWorkbook Validation Report: Passed={is_valid}")
        return is_valid
    except Exception as e:
        import traceback
        traceback.print_exc()
        return False

if __name__ == '__main__':
    ticker = sys.argv[1] if len(sys.argv) > 1 else 'ADANIPORTS'
    ok = test_company(ticker)
    if ok:
        print(f"\nSUCCESS: {ticker} passed validation!")
        sys.exit(0)
    else:
        print(f"\nFAILED: {ticker} validation failed!")
        sys.exit(1)
