import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client, valuation_engine, excel_exporter
import openpyxl

print("Fetching ADANIENT data...")
data = screener_client.fetch_company_data('ADANIENT')
print("Calculating valuation...")
val_res = valuation_engine.calculate_valuation(data)

dest_path = 'exports/ADANIENT_Valuation_Model.xlsx'
print(f"Exporting valuation model to {dest_path}...")
excel_exporter.export_via_excel_com(dest_path, data, val_res)

print("Verifying exported workbook with openpyxl...")
wb = openpyxl.load_workbook(dest_path, data_only=False)
wb_v = openpyxl.load_workbook(dest_path, data_only=True)

ws_sum_f = wb['AI Valuation Summary']
ws_sum_v = wb_v['AI Valuation Summary']
ws_dcf_v = wb_v['DCF']

print("\n--- AI Valuation Summary Sheet ---")
print("A5 (CMP): formula =", ws_sum_f['A5'].value, "| eval =", ws_sum_v['A5'].value)
print("B5 (Intrinsic): formula =", ws_sum_f['B5'].value, "| eval =", ws_sum_v['B5'].value)
print("C5 (MOS): formula =", ws_sum_f['C5'].value, "| eval =", ws_sum_v['C5'].value)
print("D5 (Verdict): formula =", ws_sum_f['D5'].value, "| eval =", ws_sum_v['D5'].value)
print("E5 (WACC): formula =", ws_sum_f['E5'].value, "| eval =", ws_sum_v['E5'].value)
print("F5 (Altman): formula =", ws_sum_f['F5'].value, "| eval =", ws_sum_v['F5'].value)
print("G5 (DuPont): formula =", ws_sum_f['G5'].value, "| eval =", ws_sum_v['G5'].value)

print("\n--- DCF Sheet ---")
print("D44 (CMP):", ws_dcf_v['D44'].value)
print("D42 (Intrinsic):", ws_dcf_v['D42'].value)
print("D35 (EV):", ws_dcf_v['D35'].value)
print("D39 (Equity):", ws_dcf_v['D39'].value)
print("D20 (WACC):", ws_dcf_v['D20'].value)
