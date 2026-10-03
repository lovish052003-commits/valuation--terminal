import os
import openpyxl
import pythoncom
import win32com.client as win32
import re

def is_company_match(p_name, c_name, ticker=''):
    if not p_name or not c_name:
        return False
    def clean_tokens(s):
        s = re.sub(r'[^\w\s]', '', s.lower())
        stopwords = {'ltd', 'limited', 'inds', 'industries', 'ind', 'india', 'co', 'corp', 'corporation'}
        tokens = [t for t in s.split() if t not in stopwords]
        return tokens
    
    tp = clean_tokens(p_name)
    tc = clean_tokens(c_name)
    if not tp or not tc:
        return False
    
    if ticker and ticker.lower() in p_name.lower().replace(' ', ''):
        return True
    
    if ' '.join(tp) == ' '.join(tc):
        return True
    
    brand_p = tp[0]
    brand_c = tc[0]
    if brand_p == brand_c:
        return True
    if len(brand_p) >= 4 and len(brand_c) >= 4:
        if brand_p.startswith(brand_c) or brand_c.startswith(brand_p):
            return True
    return False

def build_peer_range(col, tr):
    if tr <= 12:
        return f"{col}13:{col}21"
    elif tr >= 21:
        return f"{col}12:{col}20"
    elif tr == 13:
        return f"({col}12,{col}14:{col}21)"
    elif tr == 20:
        return f"({col}12:{col}19,{col}21)"
    else:
        return f"({col}12:{col}{tr-1},{col}{tr+1}:{col}21)"

# Test on exports/TATASTEEL_Valuation_Model.xlsx
model_path = os.path.abspath('exports/TATASTEEL_Valuation_Model.xlsx')

pythoncom.CoInitialize()
excel = win32.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(model_path)
    ws_comp = wb.Sheets('Comp_Valuation')
    
    # 1. Detect target company row
    c_name = "Tata Steel Ltd"
    ticker = "TATASTEEL"
    
    target_row = 12
    for r in range(12, 22):
        cell_text = str(ws_comp.Range(f'B{r}').Text or ws_comp.Range(f'B{r}').Value or '').strip()
        print(f"Row {r}: '{cell_text}'")
        if is_company_match(cell_text, c_name, ticker):
            target_row = r
            print(f"--> MATCH FOUND! Target Company '{c_name}' is in Row {target_row}")
            break
            
    # 2. Update peer benchmark statistics
    for col in ['O', 'P', 'Q']:
        pr = build_peer_range(col, target_row)
        ws_comp.Range(f'{col}23').Formula = f'=MAX({pr})'
        ws_comp.Range(f'{col}24').Formula = f'=QUARTILE({pr},1)'
        ws_comp.Range(f'{col}25').Formula = f'=MEDIAN({pr})'
        ws_comp.Range(f'{col}26').Formula = f'=AVERAGE({pr})'
        ws_comp.Range(f'{col}27').Formula = f'=QUARTILE({pr},3)'
        ws_comp.Range(f'{col}28').Formula = f'=MIN({pr})'
        
    # 3. Update Valuation section
    ws_comp.Range('B30').Value = f"{c_name} Comparable Valuation"
    ws_comp.Range('O32').Formula = f'=K{target_row}*O26'
    ws_comp.Range('P32').Formula = f'=L{target_row}*P26'
    ws_comp.Range('Q32').Formula = f'=M{target_row}*Q26'
    
    ws_comp.Range('O33').Formula = f'=$G${target_row}'
    ws_comp.Range('P33').Formula = f'=$G${target_row}'
    ws_comp.Range('Q33').Formula = f'=$G${target_row}'
    
    ws_comp.Range('O34').Formula = '=O32-O33'
    ws_comp.Range('P34').Formula = '=P32-P33'
    ws_comp.Range('Q34').Formula = '=Q32'
    
    ws_comp.Range('O35').Formula = f'=$E${target_row}'
    ws_comp.Range('P35').Formula = f'=$E${target_row}'
    ws_comp.Range('Q35').Formula = f'=$E${target_row}'
    
    ws_comp.Range('O37').Formula = '=O34/O35'
    ws_comp.Range('P37').Formula = '=P34/P35'
    ws_comp.Range('Q37').Formula = '=Q34/Q35'
    
    ws_comp.Range('O39').Formula = f'=IF(O37>$D${target_row},"Overvalued","Undervalued")'
    ws_comp.Range('P39').Formula = f'=IF(P37>$D${target_row},"Overvalued","Undervalued")'
    ws_comp.Range('Q39').Formula = f'=IF(Q37>$D${target_row},"Overvalued","Undervalued")'
    
    excel.CalculateFullRebuild()
    
    print("\n--- Verified Evaluated Cells ---")
    print("P27 (75th Percentile formula):", ws_comp.Range('P27').Formula)
    print("P27 evaluated value:", ws_comp.Range('P27').Value)
    print("O26 (Average EV/Rev):", ws_comp.Range('O26').Value)
    print("O32 (Implied EV formula):", ws_comp.Range('O32').Formula)
    print("O32 evaluated value:", ws_comp.Range('O32').Value)
    print("G13 (Net debt of Tata Steel):", ws_comp.Range('G13').Value)
    print("O33 (Net debt in valuation):", ws_comp.Range('O33').Value)
    print("E13 (Shares of Tata Steel):", ws_comp.Range('E13').Value)
    print("O35 (Shares in valuation):", ws_comp.Range('O35').Value)
    print("O37 (Implied price per share):", ws_comp.Range('O37').Value)
    print("P37 (Implied price per share):", ws_comp.Range('P37').Value)
    print("Q37 (Implied price per share):", ws_comp.Range('Q37').Value)
    print("O39 (Verdict):", ws_comp.Range('O39').Value)
    print("P39 (Verdict):", ws_comp.Range('P39').Value)
    print("Q39 (Verdict):", ws_comp.Range('Q39').Value)

    wb.Save()
    wb.Close(SaveChanges=True)
finally:
    excel.Quit()
    pythoncom.CoUninitialize()

print("\nFinished test successfully!")
