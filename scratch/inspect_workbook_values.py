import os
import sys

def main():
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print("win32com not available")
        return

    pythoncom.CoInitialize()
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
    except Exception as e:
        print(f"Excel Dispatch failed: {e}")
        return

    excel.Visible = False
    excel.DisplayAlerts = False

    p = os.path.abspath('exports/ITC_Valuation_Model.xlsx')
    if not os.path.exists(p):
        p = os.path.abspath('ITC Model.xlsx')
    
    print('Opening file:', p)
    try:
        wb = excel.Workbooks.Open(p, UpdateLinks=0, ReadOnly=True)
        excel.CalculateFull()

        ws_sum = wb.Sheets('AI Valuation Summary')
        ws_dcf = wb.Sheets('DCF')
        ws_wacc = wb.Sheets('WACC')

        print('--- AI Valuation Summary ---')
        for r in [4, 5, 8, 9, 10, 11, 12, 13, 14, 15]:
            print(f'Row {r}: A={ws_sum.Range(f"A{r}").Value} | B={ws_sum.Range(f"B{r}").Value} | C={ws_sum.Range(f"C{r}").Value} | D={ws_sum.Range(f"D{r}").Value} | E={ws_sum.Range(f"E{r}").Value} | F={ws_sum.Range(f"F{r}").Value}')

        print('--- DCF Sheet ---')
        for r in [18, 19, 20, 21, 22, 29, 33, 34, 35, 37, 38, 39, 40, 41, 43, 45, 46]:
            lbl = ws_dcf.Range(f'B{r}').Value
            val = ws_dcf.Range(f'D{r}').Value
            print(f'Row {r} ({lbl}): {val}')

        print('--- WACC Sheet ---')
        print('K26 (Rf):', ws_wacc.Range('K26').Value)
        print('K27 (ERP):', ws_wacc.Range('K27').Value)
        print('K28 (Beta):', ws_wacc.Range('K28').Value)
        print('K29 (Ke):', ws_wacc.Range('K29').Value)
        print('K40 (Ke):', ws_wacc.Range('K40').Value)
        print('K41 (We):', ws_wacc.Range('K41').Value)
        print('K43 (Kd post-tax):', ws_wacc.Range('K43').Value)
        print('K44 (Wd):', ws_wacc.Range('K44').Value)
        print('K46 (WACC):', ws_wacc.Range('K46').Value)

        wb.Close(False)
    except Exception as e:
        print(f"Error during Excel COM inspection: {e}")
    finally:
        excel.Quit()

if __name__ == '__main__':
    main()
