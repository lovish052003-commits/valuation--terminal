import sys, os
sys.path.insert(0, os.path.abspath('.'))
import screener_client
import valuation_engine
import excel_exporter
import openpyxl

print("1. Fetching SUNPHARMA data...")
sd = screener_client.fetch_company_data('SUNPHARMA')
print("Historical prices 1y count:", len(sd.get('historical_prices_1y', [])))

print("2. Calculating valuation...")
val = valuation_engine.calculate_valuation(sd)

print("3. Exporting valuation model...")
path = excel_exporter.export_valuation_model(sd, val)
print("Exported path:", path)

print("4. Inspecting exported file...")
wb = openpyxl.load_workbook(path, data_only=False)
ws_beta = wb['Beta-Regression']
ws_raw = wb['Raw Data']

print("Raw Data H6:", ws_raw['H6'].value)
print("Beta C10 formula:", ws_beta['C10'].value)
print("Beta B10 formula:", ws_beta['B10'].value)
print("Beta G10 formula:", ws_beta['G10'].value)

wb_val = openpyxl.load_workbook(path, data_only=True)
ws_beta_v = wb_val['Beta-Regression']
ws_raw_v = wb_val['Raw Data']
print("Raw Data H6 evaluated value:", ws_raw_v['H6'].value)
print("Beta C10 evaluated value:", ws_beta_v['C10'].value)
print("Beta L15 evaluated value:", ws_beta_v['L15'].value)
