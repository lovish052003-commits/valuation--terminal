import sys, os, shutil, openpyxl
sys.path.insert(0, os.path.abspath('.'))
import screener_client
import excel_exporter

screener_data = screener_client.fetch_company_data('SUNPHARMA')
test_path = 'scratch/test_populate_ds_sunpharma.xlsx'
shutil.copyfile('ITC Model.xlsx', test_path)

wb = openpyxl.load_workbook(test_path)
excel_exporter.populate_data_sheet_openpyxl(wb, screener_data)
wb.save(test_path)

wb_check = openpyxl.load_workbook(test_path, data_only=True)
ws = wb_check['Data Sheet']
print("=== Populated Data Sheet Rows 16 to 25 ===")
for r in range(16, 26):
    lbl = ws.cell(r, 1).value
    vals = [ws.cell(r, c).value for c in range(2, 12)]
    print(f"Row {r:2d} | {str(lbl):25s} | {vals}")
