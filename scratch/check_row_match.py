import os, sys
sys.path.insert(0, os.path.abspath('.'))
import openpyxl
from excel_exporter import is_company_match

wb = openpyxl.load_workbook('exports/TATASTEEL_Valuation_Model.xlsx', data_only=True)
ws = wb['Comp_Valuation']

c_name = 'Tata Steel Ltd'
ticker = 'TATASTEEL'

for r in range(12, 22):
    txt = str(ws.cell(r, 2).value or '')
    matched = is_company_match(txt, c_name, ticker)
    print(f'Row {r}: "{txt}" -> match={matched}')
