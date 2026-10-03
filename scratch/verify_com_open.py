import win32com.client
import os

excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb_path = os.path.abspath('exports/TATASTEEL_Valuation_Model.xlsx')
wb = excel.Workbooks.Open(wb_path)

for sname in ['Dupont Analysis', "Altman's Z Score", 'Comp_Valuation', 'DCF']:
    ws = wb.Sheets(sname)
    print(f'Sheet {sname} has {ws.Shapes.Count} shapes:')
    for i in range(1, ws.Shapes.Count + 1):
        sh = ws.Shapes(i)
        if 'Picture' in sh.Name:
            print(f'   Shape: {sh.Name} Left={sh.Left} Top={sh.Top} W={sh.Width} H={sh.Height}')

wb.Close(False)
excel.Quit()
print('Excel COM opened and verified successfully without errors!')
