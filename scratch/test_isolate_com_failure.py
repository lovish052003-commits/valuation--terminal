import shutil
import os
import sys
import openpyxl
import win32com.client
sys.path.insert(0, os.path.abspath('.'))

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    # Test 1: Direct copy of template
    t1 = "scratch/test1_copy.xlsx"
    shutil.copyfile("master_model_template.xlsx", t1)
    wb1 = excel.Workbooks.Open(os.path.abspath(t1))
    print("Test 1 (Direct Copy): SUCCESS, sheets =", len(wb1.Sheets))
    wb1.Close(False)

    # Test 2: Load and save with openpyxl WITHOUT any modifications
    t2 = "scratch/test2_openpyxl.xlsx"
    wb_ox = openpyxl.load_workbook("master_model_template.xlsx")
    wb_ox.save(t2)
    wb_ox.close()
    try:
        wb2 = excel.Workbooks.Open(os.path.abspath(t2))
        print("Test 2 (OpenPyXL Load/Save): SUCCESS, sheets =", len(wb2.Sheets))
        wb2.Close(False)
    except Exception as e2:
        print("Test 2 (OpenPyXL Load/Save): FAILED with", e2)

    # Test 3: OpenPyXL + strip_calc_chain
    t3 = "scratch/test3_strip.xlsx"
    shutil.copyfile(t2, t3)
    from excel_exporter import strip_calc_chain_from_xlsx
    strip_calc_chain_from_xlsx(t3)
    try:
        wb3 = excel.Workbooks.Open(os.path.abspath(t3))
        print("Test 3 (OpenPyXL + strip_calc_chain): SUCCESS, sheets =", len(wb3.Sheets))
        wb3.Close(False)
    except Exception as e3:
        print("Test 3 (OpenPyXL + strip_calc_chain): FAILED with", e3)

finally:
    excel.Quit()
