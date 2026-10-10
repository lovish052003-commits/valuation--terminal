import sys, os, time
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client
import valuation_engine
import excel_exporter
import shutil, pythoncom, win32com.client

data = screener_client.fetch_company_data("HDFCBANK")
val = valuation_engine.calculate_valuation(data)

dest_path = "exports/test_com_hdfc.xlsx"
shutil.copyfile(excel_exporter.TEMPLATE_PATH, dest_path)

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.ScreenUpdating = False

t0 = time.time()
print("Opening workbook...")
wb = excel.Workbooks.Open(os.path.abspath(dest_path), UpdateLinks=0, ReadOnly=False)
print(f"Opened in {time.time()-t0:.2f}s")

excel.Calculation = -4135  # Manual

t1 = time.time()
print("Populating Data Sheet...")
excel_exporter.populate_data_sheet(wb.Sheets('Data Sheet'), data)
print(f"Data Sheet populated in {time.time()-t1:.2f}s")

t2 = time.time()
print("Populating Raw FS...")
excel_exporter.populate_raw_fs_sheet(wb.Sheets('Raw FS'), data, val)
print(f"Raw FS populated in {time.time()-t2:.2f}s")

t3 = time.time()
print("Populating Cash Flow...")
excel_exporter.populate_cash_flow_statement_sheet(wb.Sheets('Cash Flow Statement'), data)
print(f"Cash Flow populated in {time.time()-t3:.2f}s")

t4 = time.time()
print("Populating Raw Data...")
excel_exporter.populate_raw_data_prices(wb.Sheets('Raw Data'), wb.Sheets('Beta-Regression'), data)
print(f"Raw Data populated in {time.time()-t4:.2f}s")

t5 = time.time()
print("Updating Comp Valuation...")
excel_exporter.update_comp_valuation_sheet(wb.Sheets('Comp_Valuation'), data, val)
print(f"Comp Valuation in {time.time()-t5:.2f}s")

t6 = time.time()
print("Updating WACC...")
excel_exporter.update_wacc_raw_data(wb.Sheets('Raw Data'), wb.Sheets('WACC'), data, val)
print(f"WACC in {time.time()-t6:.2f}s")

t7 = time.time()
print("Populating DuPont & Altman...")
excel_exporter.populate_dupont_altman_sheets(wb, data)
print(f"DuPont/Altman in {time.time()-t7:.2f}s")

t8 = time.time()
print("Calculating Full...")
excel.Calculation = -4105
excel.CalculateFull()
print(f"CalculateFull in {time.time()-t8:.2f}s")

wb.Save()
wb.Close(SaveChanges=True)
excel.Quit()
pythoncom.CoUninitialize()
print(f"Total time: {time.time()-t0:.2f}s")
