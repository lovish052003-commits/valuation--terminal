import os
import openpyxl

def main():
    p = 'exports/ITC_Valuation_Model.xlsx'
    wb = openpyxl.load_workbook(p, data_only=False)
    ws = wb['Control']
    for r in range(1, 20):
        for c in ['B', 'C', 'J', 'K', 'L']:
            v = ws[f'{c}{r}'].value
            if v is not None:
                print(f'{c}{r}: {v}')

if __name__ == '__main__':
    main()
