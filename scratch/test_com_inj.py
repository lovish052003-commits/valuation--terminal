import shutil
import os
import sys
import json
sys.path.insert(0, os.path.abspath('.'))
import win32com.client
from inject_raw_fs import inject_target_financials_to_raw_fs

test_xlsx = "scratch/test_com_inj.xlsx"
shutil.copyfile("master_model_template.xlsx", test_xlsx)

with open("exports/.screener_cache/UNITEDTEA_data.json", "r", encoding="utf-8") as f:
    sd = json.load(f)

inject_target_financials_to_raw_fs(test_xlsx, sd)

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(os.path.abspath(test_xlsx))
    print(f"COM SUCCESS! Sheets: {len(wb.Sheets)}")
    for s in wb.Sheets:
        print(f" - {s.Name}")
    wb.Close(False)
finally:
    excel.Quit()
