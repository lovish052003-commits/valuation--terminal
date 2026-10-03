import win32com.client
import os

excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('ITC Model.xlsx'))

for ws in wb.Sheets:
    for i in range(1, ws.Shapes.Count + 1):
        sh = ws.Shapes(i)
        if sh.Type == 13 or 'Picture' in sh.Name:
            print(f'Sheet "{ws.Name}": Shape {i} Name={sh.Name} Left={sh.Left} Top={sh.Top} W={sh.Width} H={sh.Height}')

wb.Close(False)
excel.Quit()
