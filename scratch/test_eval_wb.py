import os, sys, openpyxl
sys.path.insert(0, os.path.abspath('.'))
from screener_client import clean_num

def evaluate_openpyxl_workbook(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=False)
    ws_ds = wb['Data Sheet']
    ws_dcf = wb['DCF']
    ws_iv = wb['Intrinsic Valuation']
    ws_wacc = wb['WACC']
    
    # 1. Base inputs
    cmp_price = clean_num(ws_ds['B8'].value)
    shares = clean_num(ws_ds['K70'].value) / 10000000.0 if clean_num(ws_ds['K70'].value) > 100000 else clean_num(ws_ds['K70'].value)
    if shares <= 0:
        shares = clean_num(ws_ds['B6'].value) / 10000000.0 if clean_num(ws_ds['B6'].value) > 100000 else clean_num(ws_ds['B6'].value)

    # Cash, Investments, Debt, Minority
    cash = clean_num(ws_ds['K69'].value)
    investments = clean_num(ws_ds['K64'].value)
    cash_liquid = cash + investments
    debt = clean_num(ws_ds['K59'].value)
    minority = clean_num(ws_ds['K73'].value)

    # Base Operating EBIT (Intrinsic Valuation L38)
    # Check if raw value in Raw FS or Data Sheet
    ebit_base = 0.0
    if 'Raw FS' in wb.sheetnames:
        ws_rfs = wb['Raw FS']
        rev = clean_num(ws_rfs['AD6'].value)
        opex = clean_num(ws_rfs['AD10'].value)
        if rev > 0:
            ebit_base = rev - opex
    if ebit_base <= 0:
        ebit_base = clean_num(ws_ds['K30'].value)

    # Tax Rate (Raw Data T24)
    tax_rate = 0.2517
    if 'Raw Data' in wb.sheetnames:
        t_v = clean_num(wb['Raw Data']['T24'].value)
        if 0.1 <= t_v <= 0.4:
            tax_rate = t_v

    # Invested Capital & ROIC
    # Fixed Assets (Data Sheet K62) + Working Capital (Data Sheet K65 - K60)
    fa = clean_num(ws_ds['K62'].value)
    oa = clean_num(ws_ds['K65'].value)
    ol = clean_num(ws_ds['K60'].value)
    wc = max(0.0, oa - ol)
    inv_cap = max(1.0, fa + wc)
    roic = (ebit_base * (1.0 - tax_rate)) / inv_cap

    # WACC
    # From WACC sheet or calculated
    wacc = 0.1137
    if 'WACC' in wb.sheetnames:
        we = clean_num(ws_wacc['K40'].value)
        ke = clean_num(ws_wacc['K41'].value)
        wd = clean_num(ws_wacc['K43'].value)
        kd = clean_num(ws_wacc['K44'].value)
        if we > 0 and ke > 0:
            wacc = (we * ke) + (wd * kd)
    wacc = max(0.07, min(wacc, 0.16))

    terminal_g = 0.04
    terminal_roic = max(roic, wacc)
    terminal_reinvest = terminal_g / terminal_roic if terminal_roic > 0 else 0.20

    # Base reinvestment rate (Intrinsic Valuation L55)
    # If not evaluated, calculate from historical IC changes
    base_reinvest = 0.0065  # Median of historical reinvestment
    # If ROIC is high, fundamental Year 1 reinvest
    
    print(f"CMP: {cmp_price}, Shares: {shares}, Cash+Inv: {cash_liquid}, Debt: {debt}, EBIT: {ebit_base}, ROIC: {roic*100:.2f}%, WACC: {wacc*100:.2f}%, Terminal ROIC: {terminal_roic*100:.2f}%, Term Reinvest: {terminal_reinvest*100:.2f}%")

evaluate_openpyxl_workbook('exports/NESTLEIND_Valuation_Model.xlsx')
