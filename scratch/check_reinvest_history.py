import os, openpyxl

def main():
    import pythoncom, win32com.client
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(os.path.abspath('exports/ITC_Valuation_Model.xlsx'), UpdateLinks=0, ReadOnly=True)
    ws = wb.Sheets('Intrinsic Valuation')
    cols = ['H', 'I', 'J', 'K', 'L']
    rows = [37, 38, 40, 44, 45, 47, 49, 52, 55, 64]
    for r in rows:
        lbl = ws.Range(f'B{r}').Value or ''
        vals = [f"{c}{r}={ws.Range(f'{c}{r}').Value}" for c in cols]
        print(f"R{r} ({lbl}): " + " | ".join(vals))
    wb.Close(False)
    excel.Quit()

if __name__ == '__main__':
    main()
