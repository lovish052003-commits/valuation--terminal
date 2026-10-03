import os
import glob
import pythoncom
import win32com.client

files = glob.glob('*.xlsx')
pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

for f in files:
    if f.startswith('~'):
        continue
    try:
        path = os.path.abspath(f)
        print(f"Processing {f}...")
        wb = excel.Workbooks.Open(path, UpdateLinks=0)
        for ws in wb.Sheets:
            try:
                # xlPart = 2
                ws.Cells.Replace(What="think valuation school", Replacement="DCF Valuation Model", LookAt=2, MatchCase=False)
                ws.Cells.Replace(What="the valuation school", Replacement="DCF Valuation Model", LookAt=2, MatchCase=False)
                ws.Cells.Replace(What="valuation school", Replacement="DCF Valuation Model", LookAt=2, MatchCase=False)
            except Exception as e:
                pass
        
        wb.Save()
        wb.Close(SaveChanges=True)
        print(f"Saved {f}")
    except Exception as e:
        print(f"Error processing {f}: {e}")

excel.Quit()
pythoncom.CoUninitialize()
print("Done")
