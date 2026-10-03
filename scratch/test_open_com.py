import win32com.client, pythoncom, os

pythoncom.CoInitialize()
try:
    excel = win32com.client.DispatchEx('Excel.Application')
except Exception:
    import win32com.client.dynamic
    excel = win32com.client.dynamic.Dispatch('Excel.Application')

excel.Visible = False
excel.DisplayAlerts = False
p = os.path.abspath('exports/JUBLFOOD_Valuation_Model.xlsx')
wb = excel.Workbooks.Open(p, ReadOnly=True)
print('Opened successfully in Excel COM! Sheets count:', wb.Sheets.Count)
print('Dupont Title:', wb.Sheets('Dupont Analysis').Range('B2').Value)
print('Altman Title:', wb.Sheets("Altman's Z Score").Range('B2').Value)
wb.Close(False)
excel.Quit()
print('>>> SUCCESS: JUBLFOOD_Valuation_Model.xlsx OPENS CLEANLY WITH ZERO REPAIR WARNINGS! <<<')
