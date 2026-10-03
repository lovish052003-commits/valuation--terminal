import sys, os
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__) + '/..'))
import screener_client, valuation_engine, excel_exporter

print("1. Fetching Wipro data...")
d = screener_client.fetch_company_data('Wipro')
print("Company:", d.get('company_name'), "| Ticker:", d.get('ticker'))
val = valuation_engine.calculate_valuation(d)

print("2. Exporting Wipro...")
try:
    path = excel_exporter.export_valuation_model(d, val, "")
    print("Exported path:", path)
except Exception as e:
    import traceback
    traceback.print_exc()
