import openpyxl
import sys

# Ensure UTF-8 output
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

p = r"C:\Users\LENOVO\Downloads\Advance Financial Project\exports\SBIN_Valuation_Model.xlsx"
wb_f = openpyxl.load_workbook(p, data_only=False)
wb_v = openpyxl.load_workbook(p, data_only=True)

print(f"\n=======================================================")
print(f"DETAILED INSPECTION OF EXPORTED SBIN WORKBOOK: {p}")
print(f"=======================================================")

# 1. Data Sheet rows 1 to 20 (Market data, price, shares, etc.)
ws_f = wb_f['Data Sheet']
ws_v = wb_v['Data Sheet']
print("\n--- Data Sheet Rows 1 to 20 ---")
for r in range(1, 21):
    row_v = [ws_v.cell(r, c).value for c in range(1, 12)]
    print(f"Row {r:2d}: {row_v}")

# 2. Intrinsic Valuation around L38 and L69
if 'Intrinsic Valuation' in wb_f.sheetnames:
    ws_iv_f = wb_f['Intrinsic Valuation']
    ws_iv_v = wb_v['Intrinsic Valuation']
    print("\n--- Intrinsic Valuation Rows 30 to 45 ---")
    for r in range(30, 46):
        lbl = ws_iv_v.cell(r, 2).value or ws_iv_v.cell(r, 1).value
        l_form = ws_iv_f.cell(r, 12).value
        l_val = ws_iv_v.cell(r, 12).value
        print(f"Row {r:2d} | Label: {str(lbl)[:30]:30s} | L Form: {str(l_form)[:35]:35s} | L Val: {l_val}")

# 3. AI Valuation Summary rows 8 to 35
ws_ai_v = wb_v['AI Valuation Summary']
ws_ai_f = wb_f['AI Valuation Summary']
print("\n--- AI Valuation Summary Rows 8 to 35 ---")
for r in range(8, 36):
    r_v = [ws_ai_v.cell(r, c).value for c in range(1, 8)]
    if any(r_v):
        print(f"Row {r:2d} | {r_v}")

# 4. Comp_Valuation
if 'Comp_Valuation' in wb_f.sheetnames:
    ws_cv_v = wb_v['Comp_Valuation']
    ws_cv_f = wb_f['Comp_Valuation']
    print("\n--- Comp_Valuation Rows 10 to 40 ---")
    for r in range(10, 41):
        d_name = ws_cv_v.cell(r, 4).value
        m_cap = ws_cv_v.cell(r, 6).value
        ev = ws_cv_v.cell(r, 8).value
        rev = ws_cv_v.cell(r, 11).value
        ebitda = ws_cv_v.cell(r, 12).value
        net_inc = ws_cv_v.cell(r, 13).value
        ev_rev = ws_cv_v.cell(r, 15).value
        ev_ebitda = ws_cv_v.cell(r, 16).value
        pe = ws_cv_v.cell(r, 17).value
        print(f"Row {r:2d} | {str(d_name)[:18]:18s} | MCap:{m_cap} | EV:{ev} | Rev:{rev} | EBITDA:{ebitda} | NI:{net_inc} | EV/R:{ev_rev} | EV/E:{ev_ebitda} | P/E:{pe}")
