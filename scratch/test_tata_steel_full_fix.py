import os
import sys
sys.path.insert(0, os.path.abspath('.'))
import openpyxl
import pythoncom
import win32com.client as win32
import re
import pandas as pd

from screener_client import fetch_company_data
from valuation_engine import calculate_valuation

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

# 1. Fetch data
screener_data = fetch_company_data('TATASTEEL')
val_res = calculate_valuation(screener_data)

# 2. Test updating Raw FS and Comp_Valuation on a fresh copy of ITC Model.xlsx
import shutil
test_out = os.path.abspath('exports/TATASTEEL_TEST_FULL.xlsx')
shutil.copy('ITC Model.xlsx', test_out)

pythoncom.CoInitialize()
excel = win32.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(test_out)
    ws_raw = wb.Sheets('Raw FS')
    ws_comp = wb.Sheets('Comp_Valuation')
    
    peers_df = screener_data.get('peers_df', pd.DataFrame())
    c_name = screener_data.get('company_name', '').strip()
    ticker = screener_data.get('ticker', '').strip()
    
    # Write Peers into Raw FS
    if peers_df is not None and not peers_df.empty:
        for p_idx, (_, p_row) in enumerate(peers_df.head(10).iterrows(), start=1):
            r_left = 56 + p_idx   # 57 to 66
            r_right = 55 + p_idx  # 56 to 65
            
            p_name = str(p_row.get('Company') or p_row.get('Name') or '').strip()
            p_cmp = float(p_row.get('CMP  Rs.') or p_row.get('CMP Rs.') or 0)
            p_mcap = float(p_row.get('Mar Cap  Rs.Cr.') or p_row.get('Mar Cap Rs.Cr.') or 0)
            sales_qtr = float(p_row.get('Sales Qtr  Rs.Cr.') or p_row.get('Sales Qtr Rs.Cr.') or 0)
            np_qtr = float(p_row.get('NP Qtr  Rs.Cr.') or p_row.get('NP Qtr Rs.Cr.') or 0)
            pe_val = float(p_row.get('P/E') or 0)
            
            # Left side
            ws_raw.Cells(r_left, 11).Value = p_idx
            ws_raw.Cells(r_left, 12).Value = p_name
            ws_raw.Cells(r_left, 13).Value = p_cmp
            if p_cmp > 0 and p_mcap > 0:
                ws_raw.Cells(r_left, 14).Value = round(p_mcap / p_cmp, 2)
                
            # Right side
            ws_raw.Cells(r_right, 44).Value = p_mcap
            ws_raw.Cells(r_right, 45).Formula = f'=AZ{r_right}-BA{r_right}'
            ws_raw.Cells(r_right, 46).Formula = f'=AR{r_right}+AS{r_right}'
            
            if sales_qtr > 0:
                ws_raw.Cells(r_right, 47).Value = sales_qtr * 4
            if np_qtr > 0:
                ws_raw.Cells(r_right, 49).Value = np_qtr * 4
                
            # If this peer is target company, fill exact debt and cash
            if is_company_match(p_name, c_name, ticker):
                ws_raw.Cells(r_right, 52).Value = float(val_res.get('total_debt', 0))
                ws_raw.Cells(r_right, 53).Value = float(val_res.get('cash_estimate', 0))
                
    # Detect target row in Comp_Valuation
    target_row = 12
    for r in range(12, 22):
        cell_text = str(ws_comp.Range(f'B{r}').Text or ws_comp.Range(f'B{r}').Value or '').strip()
        if is_company_match(cell_text, c_name, ticker):
            target_row = r
            break
            
    print(f"Detected Target Company '{c_name}' in Comp_Valuation Row {target_row}")
    
    # Update Peer Statistics (Rows 23-28)
    for col in ['O', 'P', 'Q']:
        pr = build_peer_range(col, target_row)
        ws_comp.Range(f'{col}23').Formula = f'=MAX({pr})'
        ws_comp.Range(f'{col}24').Formula = f'=QUARTILE({pr},1)'
        ws_comp.Range(f'{col}25').Formula = f'=MEDIAN({pr})'
        ws_comp.Range(f'{col}26').Formula = f'=AVERAGE({pr})'
        ws_comp.Range(f'{col}27').Formula = f'=QUARTILE({pr},3)'
        ws_comp.Range(f'{col}28').Formula = f'=MIN({pr})'
        
    # Update Valuation Section (Rows 30-39)
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
    
    print("\n--- Row 12 (Peer 1 = JSW Steel) ---")
    print("  Name (B12):", ws_comp.Range('B12').Value)
    print("  Price (D12):", ws_comp.Range('D12').Value)
    print("  Shares (E12):", ws_comp.Range('E12').Value)
    print("  Net Debt (G12):", ws_comp.Range('G12').Value)
    print("  EV (H12):", ws_comp.Range('H12').Value)
    print("  Revenue (K12):", ws_comp.Range('K12').Value)
    print("  EV/Rev (O12):", ws_comp.Range('O12').Value)
    print("  EV/EBITDA (P12):", ws_comp.Range('P12').Value)
    
    print("\n--- Row 13 (Peer 2 = Target Tata Steel) ---")
    print("  Name (B13):", ws_comp.Range('B13').Value)
    print("  Price (D13):", ws_comp.Range('D13').Value)
    print("  Shares (E13):", ws_comp.Range('E13').Value)
    print("  Net Debt (G13):", ws_comp.Range('G13').Value)
    print("  EV (H13):", ws_comp.Range('H13').Value)
    print("  Revenue (K13):", ws_comp.Range('K13').Value)
    print("  EV/Rev (O13):", ws_comp.Range('O13').Value)
    print("  EV/EBITDA (P13):", ws_comp.Range('P13').Value)
    
    print("\n--- Peer Benchmark Statistics (Rows 23-28) ---")
    print("  P23 (High EV/EBITDA):", ws_comp.Range('P23').Formula, "->", ws_comp.Range('P23').Value)
    print("  P26 (Average EV/EBITDA):", ws_comp.Range('P26').Formula, "->", ws_comp.Range('P26').Value)
    print("  P27 (75th Percentile):", ws_comp.Range('P27').Formula, "->", ws_comp.Range('P27').Value)
    
    print("\n--- Comparable Valuation (Rows 30-39) ---")
    print("  Title (B30):", ws_comp.Range('B30').Value)
    print("  Implied EV (O32, P32, Q32):", ws_comp.Range('O32').Formula, "->", ws_comp.Range('O32').Value)
    print("  Net Debt (O33):", ws_comp.Range('O33').Formula, "->", ws_comp.Range('O33').Value)
    print("  Shares (O35):", ws_comp.Range('O35').Formula, "->", ws_comp.Range('O35').Value)
    print("  Implied Price (O37, P37, Q37):", ws_comp.Range('O37').Value, ws_comp.Range('P37').Value, ws_comp.Range('Q37').Value)
    print("  Verdict (O39):", ws_comp.Range('O39').Formula, "->", ws_comp.Range('O39').Value)

    wb.Save()
    wb.Close(SaveChanges=True)
finally:
    excel.Quit()
    pythoncom.CoUninitialize()

print("\nAll tests completed!")
