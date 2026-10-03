import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client
import valuation_engine
import excel_exporter

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

screener_data = screener_client.fetch_company_data('OIL')
val_result = valuation_engine.calculate_valuation(screener_data)

test_com_path = os.path.join(excel_exporter.EXPORT_DIR, "test_com_oil.xlsx")
print(f"Testing export_via_excel_com directly to: {test_com_path}...")
try:
    ok = excel_exporter.export_via_excel_com(test_com_path, screener_data, val_result)
    print("export_via_excel_com result:", ok)
except Exception as e:
    import traceback
    traceback.print_exc()
