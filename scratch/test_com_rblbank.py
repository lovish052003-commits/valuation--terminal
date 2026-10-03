import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import traceback
import screener_client
import valuation_engine
import excel_exporter

print("Testing export_via_excel_com for RBLBANK...")
ticker = "RBLBANK"
screener_data = screener_client.fetch_company_data(ticker)
val = valuation_engine.calculate_valuation(screener_data)

dest = os.path.join(excel_exporter.EXPORT_DIR, "TEST_RBLBANK_COM.xlsx")
try:
    success = excel_exporter.export_via_excel_com(dest, screener_data, val)
    print("export_via_excel_com success:", success)
except Exception as e:
    print("export_via_excel_com FAILED:", e)
    traceback.print_exc()
