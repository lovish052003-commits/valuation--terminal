import os
import pythoncom
import win32com.client

def main():
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    p = os.path.abspath('exports/ITC_Valuation_Model.xlsx')
    wb = excel.Workbooks.Open(p, UpdateLinks=0, ReadOnly=True)
    excel.CalculateFull()

    ws = wb.Sheets('DCF')
    cols = ['H', 'I', 'J', 'K', 'L', 'M']
    rows = [8, 9, 10, 11, 12, 13, 14, 16]
    for r in rows:
        lbl = ws.Range(f'B{r}').Value or ''
        vals = [f"{c}{r}={ws.Range(f'{c}{r}').Value}" for c in cols]
        print(f"R{r} ({lbl}): " + " | ".join(vals))

    wb.Close(False)
    excel.Quit()

if __name__ == '__main__':
    main()
