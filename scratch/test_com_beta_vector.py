import os, sys
import win32com.client, pythoncom

dest_path = os.path.abspath('exports/test_com_beta.xlsx')
import shutil
shutil.copyfile('ITC Model.xlsx', dest_path)

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(dest_path)
    ws_beta = wb.Sheets('Beta-Regression')
    max_pts = 247
    end_beta_r = 10 + max_pts - 1
    
    bcd_formulas = []
    fgh_formulas = []
    for i in range(max_pts):
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
        
    ws_beta.Range(ws_beta.Cells(10, 2), ws_beta.Cells(end_beta_r, 4)).Formula = bcd_formulas
    ws_beta.Range(ws_beta.Cells(10, 6), ws_beta.Cells(end_beta_r, 8)).Formula = fgh_formulas
    
    wb.Save()
    wb.Close(SaveChanges=True)
    print("SUCCESS: COM vector write for Beta-Regression worked perfectly!")
except Exception as e:
    print("ERROR in COM vector write:", e)
finally:
    excel.Quit()
    del excel
    pythoncom.CoUninitialize()
