import win32com.client, pythoncom, os

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(os.path.abspath('ITC Model.xlsx'))
ws = wb.Sheets('Forecasting')

# Check array properties
print('C5 HasArray:', ws.Range('C5').HasArray)
print('C5 CurrentArray:', ws.Range('C5').CurrentArray.Address)

# To replace array formula in C5:C14:
ws.Range('C5:C14').ClearContents()
ws.Range('C5:C13').FormulaArray = "=TRANSPOSE('Data Sheet'!B16:J16)"
ws.Range('C14').Formula = '=C13+365'

ws.Range('H5:H14').ClearContents()
ws.Range('H5:H13').FormulaArray = "=TRANSPOSE('Data Sheet'!B16:J16)"
ws.Range('H14').Formula = '=H13+365'

ws.Range('M5:M14').ClearContents()
ws.Range('M5:M13').FormulaArray = "=TRANSPOSE('Data Sheet'!B16:J16)"
ws.Range('M14').Formula = '=M13+365'

# Ensure rows 15-19 are dynamic formulas
for r in range(15, 20):
    prev = r - 1
    ws.Range(f'C{r}').Formula = f'=C{prev}+365'
    ws.Range(f'H{r}').Formula = f'=H{prev}+365'
    ws.Range(f'M{r}').Formula = f'=M{prev}+365'

excel.CalculateFull()

print('\nResults in Forecasting:')
for r in range(13, 20):
    c_val = ws.Range(f'C{r}').Text
    c_f = ws.Range(f'C{r}').Formula
    h_val = ws.Range(f'H{r}').Text
    m_val = ws.Range(f'M{r}').Text
    print(f'Row {r:2d}: C={c_val:<15} (Formula={c_f}) | H={h_val:<15} | M={m_val:<15}')

wb.Close(SaveChanges=False)
excel.Quit()
pythoncom.CoUninitialize()
print('DONE!')
