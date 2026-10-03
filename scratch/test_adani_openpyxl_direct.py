import sys
import traceback
sys.path.insert(0, '.')
import screener_client
import valuation_engine
import excel_exporter
import openpyxl
import shutil

print("Fetching ADANIENT...")
sd = screener_client.fetch_company_data("ADANIENT")
val = valuation_engine.calculate_valuation(sd)

dest_test = "exports/TEST_ADANI_OPENPYXL.xlsx"
shutil.copyfile('ITC Model.xlsx', dest_test)

wb = openpyxl.load_workbook(dest_test)
print("Testing populate_data_sheet_openpyxl...")
try:
    excel_exporter.populate_data_sheet_openpyxl(wb, sd)
    print("SUCCESS: populate_data_sheet_openpyxl completed without error.")
except Exception as e:
    print("ERROR in populate_data_sheet_openpyxl:")
    traceback.print_exc()

print("Testing populate_raw_fs_sheet_openpyxl...")
try:
    excel_exporter.populate_raw_fs_sheet_openpyxl(wb, sd, val)
    print("SUCCESS: populate_raw_fs_sheet_openpyxl completed.")
except Exception as e:
    print("ERROR in populate_raw_fs_sheet_openpyxl:")
    traceback.print_exc()

wb.save(dest_test)

wb_check = openpyxl.load_workbook(dest_test, data_only=True)
print("\nCheck values in TEST_ADANI_OPENPYXL.xlsx:")
print("Data Sheet B1:", wb_check['Data Sheet']['B1'].value)
print("Data Sheet K17:", wb_check['Data Sheet']['K17'].value)
print("Data Sheet K30:", wb_check['Data Sheet']['K30'].value)
