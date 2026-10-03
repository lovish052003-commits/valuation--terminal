import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client
import valuation_engine
import win32com.client as win32

data = screener_client.fetch_company_data('ADANIENT')
val_res = valuation_engine.calculate_valuation(data)

print("=== TERMINAL (valuation_engine.py) RESULTS ===")
print("Terminal CMP:", val_res.get('current_price'))
print("Terminal Intrinsic Value:", val_res.get('intrinsic_value'))
print("Terminal DCF Value:", val_res.get('dcf_value'))
print("Terminal WACC:", val_res.get('wacc'))
print("Terminal EV:", val_res.get('enterprise_value'))
print("Terminal Equity Value:", val_res.get('equity_value'))
print("Terminal Margin of Safety:", val_res.get('margin_of_safety'))
print("Terminal Verdict:", val_res.get('verdict'))

excel_file = os.path.abspath('exports/ADANIENT_Valuation_Model.xlsx')
if os.path.exists(excel_file):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(excel_file)
    excel.CalculateFull()
    
    ws_dcf = wb.Sheets('DCF')
    print("\n=== EXCEL DCF SHEET RESULTS ===")
    print("Excel DCF D44 (CMP):", ws_dcf.Range('D44').Value)
    print("Excel DCF D42 (Intrinsic Value/Share):", ws_dcf.Range('D42').Value)
    print("Excel DCF D35 (Enterprise Value):", ws_dcf.Range('D35').Value)
    print("Excel DCF D39 (Equity Value):", ws_dcf.Range('D39').Value)
    print("Excel DCF D20 (WACC):", ws_dcf.Range('D20').Value)
    print("Excel DCF D40 (Shares Cr):", ws_dcf.Range('D40').Value)
    print("Excel DCF D37 (Cash):", ws_dcf.Range('D37').Value)
    print("Excel DCF D38 (Debt):", ws_dcf.Range('D38').Value)
    
    ws_sum = wb.Sheets('AI Valuation Summary')
    print("\n=== EXCEL AI VALUATION SUMMARY ===")
    print("Excel Summary A5 (CMP):", ws_sum.Range('A5').Value)
    print("Excel Summary B5 (Intrinsic Value):", ws_sum.Range('B5').Value)
    print("Excel Summary C5 (Margin of Safety):", ws_sum.Range('C5').Value)
    print("Excel Summary D5 (Verdict):", ws_sum.Range('D5').Value)
    print("Excel Summary E5 (WACC):", ws_sum.Range('E5').Value)
    
    wb.Close(SaveChanges=False)
    excel.Quit()
