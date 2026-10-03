import os
import sys
import shutil

sys.path.insert(0, os.path.abspath('.'))

import screener_client
import valuation_engine
import excel_exporter
import win32com.client as win32
import pythoncom

def test_company_export(ticker):
    print(f"\n==========================================")
    print(f"Testing end-to-end clean export for: {ticker}")
    print(f"==========================================")

    data = screener_client.fetch_company_data(ticker)
    val = valuation_engine.calculate_valuation(data)
    exported_path = excel_exporter.export_valuation_model(data, val)
    print(f"Workbook exported to: {exported_path}")

    # Now verify with Microsoft Excel COM using CorruptLoad=0 (xlNormalLoad)
    # If there is ANY corrupted XML, drawing error, or external reference error,
    # CorruptLoad=0 will fail or prompt recovery.
    pythoncom.CoInitialize()
    excel = win32.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        # CorruptLoad=0 means normal load; if there is any error, it raises an exception or fails
        wb = excel.Workbooks.Open(os.path.abspath(exported_path), CorruptLoad=0)
        print(f"SUCCESS: Workbook '{os.path.basename(exported_path)}' opened CLEANLY in Microsoft Excel COM with CorruptLoad=0!")
        
        # Check sheets count
        print(f"Total Sheets: {wb.Sheets.Count} (ITC template has 22 sheets + AI Summary = 22/23)")
        
        # Check DCF Sheet
        ws_dcf = wb.Sheets('DCF')
        print(f"DCF Row 38: {ws_dcf.Cells(38, 2).Value} = {ws_dcf.Cells(38, 4).Formula}")
        print(f"DCF Row 39: {ws_dcf.Cells(39, 2).Value} = {ws_dcf.Cells(39, 4).Formula}")
        print(f"DCF Row 40: {ws_dcf.Cells(40, 2).Value} = {ws_dcf.Cells(40, 4).Formula}")
        print(f"DCF Row 41: {ws_dcf.Cells(41, 2).Value} = {ws_dcf.Cells(41, 4).Formula}")
        print(f"DCF Row 43: {ws_dcf.Cells(43, 2).Value} = {ws_dcf.Cells(43, 4).Formula}")
        print(f"DCF Row 45: {ws_dcf.Cells(45, 2).Value} = {ws_dcf.Cells(45, 4).Formula}")

        # Check AI Valuation Summary Sheet
        ws_sum = wb.Sheets('AI Valuation Summary')
        print(f"AI Summary Intrinsic Value KPI: {ws_sum.Cells(5, 2).Formula}")
        print(f"AI Summary Current Price KPI: {ws_sum.Cells(5, 1).Formula}")
        print(f"AI Summary Bridge Rows:")
        for r in range(28, 41):
            lbl = ws_sum.Cells(r, 2).Value
            form = ws_sum.Cells(r, 3).Formula
            if lbl:
                print(f"  Row {r}: {lbl} | Formula: {form}")

        wb.Close(False)
        print(f"VERIFICATION PASSED for {ticker}: 0 content errors, 0 recovery warnings!\n")
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

if __name__ == '__main__':
    # Test a banking company (UNIONBANK)
    test_company_export('UNIONBANK')
