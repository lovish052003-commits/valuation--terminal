import win32com.client, pythoncom, os, shutil

# 1. Restore pristine template
src = os.path.abspath('exports/test_com_beta.xlsx')
dst = os.path.abspath('ITC Model.xlsx')
shutil.copyfile(src, dst)
print("1. Restored ITC Model.xlsx from test_com_beta.xlsx")

# 2. Patch with COM (zero openpyxl corruption)
pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(dst)
    ws_beta = wb.Sheets('Beta-Regression')
    
    ws_beta.Range('B7').Formula = "='Data Sheet'!B1&\" Daily Returns\""
    ws_beta.Range('F7').Value = "Nifty 50 Daily Returns"
    ws_beta.Range('J7').Value = "Beta Drifting"

    bcd_formulas = []
    fgh_formulas = []
    for i in range(243):
        raw_r = 6 + i
        beta_r = 10 + i
        b_f = f"='Raw Data'!G{raw_r}"
        c_f = f"='Raw Data'!H{raw_r}"
        d_f = "-" if i == 0 else f"=C{beta_r-1}/C{beta_r}-1"
        bcd_formulas.append([b_f, c_f, d_f])

        f_f = f"='Raw Data'!G{raw_r}"
        g_f = f"='Raw Data'!J{raw_r}"
        h_f = "-" if i == 0 else f"=G{beta_r-1}/G{beta_r}-1"
        fgh_formulas.append([f_f, g_f, h_f])

    ws_beta.Range(ws_beta.Cells(10, 2), ws_beta.Cells(252, 4)).Formula = bcd_formulas
    ws_beta.Range(ws_beta.Cells(10, 6), ws_beta.Cells(252, 8)).Formula = fgh_formulas

    ws_beta.Range('O11').Formula = "=SLOPE(D10:D252, H10:H252)"
    ws_beta.Range('O12').Formula = "=_xlfn.COVARIANCE.S(D10:D252, H10:H252)/_xlfn.VAR.S(H10:H252)"
    ws_beta.Range('L9').Formula = "=O11"
    ws_beta.Range('L10').Value = 0.75
    ws_beta.Range('L12').Value = 1.0
    ws_beta.Range('L13').Value = 0.25
    ws_beta.Range('L15').Formula = "=(L9*L10)+(L12*L13)"

    wb.Save()
    wb.Close(SaveChanges=True)
    print("2. Successfully patched ITC Model.xlsx with COM!")
except Exception as e:
    print("FAILED during COM patch:", e)
finally:
    excel.Quit()
    pythoncom.CoUninitialize()

# 3. Verify clean reopening with COM
pythoncom.CoInitialize()
excel2 = win32com.client.DispatchEx('Excel.Application')
excel2.Visible = False
excel2.DisplayAlerts = False
try:
    wb2 = excel2.Workbooks.Open(dst)
    print("3. VERIFIED: ITC Model.xlsx opened cleanly with COM! Sheets:", wb2.Sheets.Count)
    wb2.Close(SaveChanges=False)
except Exception as e:
    print("3. FAILED verifying ITC Model.xlsx:", e)
finally:
    excel2.Quit()
    pythoncom.CoUninitialize()
