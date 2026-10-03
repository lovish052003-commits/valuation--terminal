import os, sys, shutil
sys.path.insert(0, os.path.abspath('.'))

import screener_client
import valuation_engine
import excel_exporter

# Temporarily point TEMPLATE_PATH to TEST_MARUTI_COM.xlsx
orig_template = excel_exporter.TEMPLATE_PATH
excel_exporter.TEMPLATE_PATH = os.path.abspath('exports/TEST_MARUTI_COM.xlsx')

screener_data = screener_client.fetch_company_data('UNITEDTEA')
val_result = valuation_engine.calculate_valuation(screener_data)

out_path = os.path.abspath('scratch/UNITEDTEA_COM_TEST.xlsx')
try:
    print("Calling export_via_excel_com with clean template...")
    success = excel_exporter.export_via_excel_com(out_path, screener_data, val_result)
    print("export_via_excel_com returned:", success)
    
    # Verify opening the exported file with Excel COM
    import win32com.client
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(out_path)
    print("SUCCESS! Exported file opened in Excel COM cleanly! Sheets:", len(wb.Sheets))
    wb.Close(False)
    excel.Quit()
except Exception as e:
    import traceback
    traceback.print_exc()
finally:
    excel_exporter.TEMPLATE_PATH = orig_template
