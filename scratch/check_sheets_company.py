import openpyxl

wb = openpyxl.load_workbook('Nestle India Model.xlsx', data_only=True)
wb_f = openpyxl.load_workbook('Nestle India Model.xlsx', data_only=False)

print("=== CHECKING ALL SHEETS IN 'Nestle India Model.xlsx' FOR NESTLE VS ITC DATA ===")
for sheet in wb.sheetnames:
    ws = wb[sheet]
    ws_f = wb_f[sheet]
    # Check A1:B10
    sample_text = ""
    for r in range(1, 15):
        for c in range(1, 6):
            v = str(ws.cell(r, c).value or '')
            sample_text += " " + v
    is_nestle = 'nestle' in sample_text.lower()
    is_itc = 'itc' in sample_text.lower()
    print(f"Sheet: {sheet:<25} | Nestle mentioned: {str(is_nestle):<5} | ITC mentioned: {str(is_itc):<5}")
