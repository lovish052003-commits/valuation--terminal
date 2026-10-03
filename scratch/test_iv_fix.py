import win32com.client, os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

wb_path = os.path.abspath('exports/KALYANKJIL_Valuation_Model.xlsx')
wb = excel.Workbooks.Open(wb_path)

ws_iv = wb.Sheets('Intrinsic Valuation')
for col_l in ['I', 'J', 'K', 'L']:
    ws_iv.Range(f'{col_l}40').Formula = f'=IFERROR({col_l}38/{col_l}37, 0)'
    ws_iv.Range(f'{col_l}52').Formula = f'=IFERROR({col_l}51/{col_l}49, 0)'

ws_iv.Range('L54').Formula = '=IFERROR(AVERAGE(I52:L52), 0.45)'
ws_iv.Range('L55').Formula = '=IFERROR(MEDIAN(I52:L52), 0.45)'

ws_dcf = wb.Sheets('DCF')
ws_dcf.Range('I11').Formula = '=IFERROR(IF(\'Intrinsic Valuation\'!$L$55>0, \'Intrinsic Valuation\'!$L$55, 0.45), 0.45)'
ws_dcf.Range('H8').Formula = "='Data Sheet'!K28+'Data Sheet'!K27"

excel.CalculateFull()

print('Intrinsic Valuation L55 Value:', ws_iv.Range('L55').Value)
print('DCF I11 Value:', ws_dcf.Range('I11').Value)
print('DCF D35 (Operating Assets / EV):', ws_dcf.Range('D35').Value)
print('DCF D40 (Equity Value):', ws_dcf.Range('D40').Value)
print('DCF D43 (Intrinsic Value/Share):', ws_dcf.Range('D43').Value)

ws_sum = wb.Sheets('AI Valuation Summary')
print('Summary B5 (Intrinsic Value):', ws_sum.Range('B5').Value)
print('Summary C5 (Margin of Safety):', ws_sum.Range('C5').Value)
print('Summary D5 (Verdict):', ws_sum.Range('D5').Value)

wb.Close(SaveChanges=False)
excel.Quit()
