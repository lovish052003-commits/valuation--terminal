import sys, os
sys.path.insert(0, os.path.abspath('.'))
import screener_client
import valuation_engine
import excel_exporter

screener_data = screener_client.fetch_company_data('SUNPHARMA')
val_result = valuation_engine.calculate_valuation(screener_data)

test_out = os.path.abspath('scratch/test_sunpharma_com_direct.xlsx')
try:
    success = excel_exporter.export_via_excel_com(test_out, screener_data, val_result)
    print("export_via_excel_com returned:", success)
except Exception as e:
    import traceback
    traceback.print_exc()
