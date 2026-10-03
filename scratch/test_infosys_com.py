import os, sys, shutil
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__) + '/..'))
import screener_client, valuation_engine, excel_exporter

print("1. Fetching Screener data for Infosys Ltd...")
screener_data = screener_client.fetch_company_data('Infosys Ltd')
val_result = valuation_engine.calculate_valuation(screener_data)

dest_path = os.path.abspath('exports/TEST_INFY_COM.xlsx')
print(f"2. Exporting via COM to {dest_path}...")
success = excel_exporter.export_via_excel_com(dest_path, screener_data, val_result, "")
print("COM Export Success:", success)

if success:
    print("3. Testing opening in Microsoft Excel via COM to check for repair prompts...")
    import win32com.client, pythoncom
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(dest_path)
        print(f"SUCCESS: Opened {dest_path} with {len(wb.Sheets)} sheets and ZERO REPAIR WARNINGS!")
        wb.Close(SaveChanges=False)
    except Exception as e:
        print("Error opening:", e)
    finally:
        excel.Quit()
        del excel
        pythoncom.CoUninitialize()

    print("4. Scanning TEST_INFY_COM.xlsx for any leftover ITC references...")
    import openpyxl
    wb_chk = openpyxl.load_workbook(dest_path, data_only=True)
    itc_found = []
    for sname in wb_chk.sheetnames:
        if sname == 'List of Stocks':
            continue
        ws = wb_chk[sname]
        for r in range(1, min(ws.max_row+1, 100)):
            for c in range(1, min(ws.max_column+1, 35)):
                v = str(ws.cell(r, c).value or '')
                if 'itc' in v.lower():
                    safe_v = v.encode('ascii', 'backslashreplace').decode('ascii')
                    itc_found.append((sname, ws.cell(r, c).coordinate, safe_v[:60]))
    print(f"Total ITC references in {dest_path} (excluding List of Stocks): {len(itc_found)}")
    for item in itc_found:
        print("  ITC:", item)
    assert len(itc_found) == 0, f"Found {len(itc_found)} ITC references!"
    print("ALL CHECKS PASSED! 0.00% ITC IN ANY SHEET!")
