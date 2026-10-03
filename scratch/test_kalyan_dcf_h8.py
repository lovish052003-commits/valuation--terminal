import win32com.client
import os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

wb_path = os.path.abspath('exports/KALYANKJIL_Valuation_Model.xlsx')
wb = excel.Workbooks.Open(wb_path)

ws_dcf = wb.Sheets('DCF')
print('Current H8 Formula:', ws_dcf.Range('H8').Formula)
print('Current H8 Value:', ws_dcf.Range('H8').Value)
print('Current D40 (Equity Val):', ws_dcf.Range('D40').Value)
print('Current D43 (Intrinsic Val):', ws_dcf.Range('D43').Value)

# Update H8 to use Data Sheet EBIT
ws_dcf.Range('H8').Formula = "='Data Sheet'!K28+'Data Sheet'!K27"
excel.CalculateFull()

print('\nAFTER FIX:')
print('New H8 Formula:', ws_dcf.Range('H8').Formula)
print('New H8 Value:', ws_dcf.Range('H8').Value)
print('New D35 (Operating Assets / EV):', ws_dcf.Range('D35').Value)
print('New D40 (Equity Val):', ws_dcf.Range('D40').Value)
print('New D43 (Intrinsic Val):', ws_dcf.Range('D43').Value)

ws_sum = wb.Sheets('AI Valuation Summary')
print('Summary B5 (Intrinsic Val):', ws_sum.Range('B5').Value)
print('Summary C5 (Margin of Safety):', ws_sum.Range('C5').Value)
print('Summary D5 (Verdict):', ws_sum.Range('D5').Value)

wb.Close(SaveChanges=False)
excel.Quit()
