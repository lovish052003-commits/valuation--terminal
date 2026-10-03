import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import traceback
import screener_client
import valuation_engine
import excel_exporter

print("=" * 80)
print("TESTING EXCEL EXPORT FOR RELIGARE")
print("=" * 80)

ticker = "RELIGARE"
try:
    screener_data = screener_client.fetch_company_data(ticker)
    print(f"Target: {screener_data['company_name']} ({screener_data['ticker']}) | Sector: {screener_data['sector']}")
    val = valuation_engine.calculate_valuation(screener_data)
    print(f"Valuation calculated: CMP={val['current_price']}, IV={val['intrinsic_value_per_share']}")
    
    print("Calling export_valuation_model...")
    p = excel_exporter.export_valuation_model(screener_data, val)
    print(f"SUCCESS: Exported to {p}")
except Exception as e:
    print(f"FAILED with error: {e}")
    traceback.print_exc()
