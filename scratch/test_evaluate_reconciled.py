import os
import win32com.client as win32

dest = os.path.abspath('exports/ADANIENT_Valuation_Model_reconciled.xlsx')
excel = None
try:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
except Exception:
    excel = win32.Dispatch('Excel.Application')

excel.Visible = False
excel.DisplayAlerts = False
excel.ScreenUpdating = False

try:
    wb = excel.Workbooks.Open(dest)
    excel.CalculateFull()
    
    ws_dcf = wb.Sheets('DCF')
    iv_dcf = ws_dcf.Range('D42').Value
    cmp_dcf = ws_dcf.Range('D44').Value
    ev_dcf = ws_dcf.Range('D35').Value
    eq_dcf = ws_dcf.Range('D39').Value
    fcf_y1 = ws_dcf.Range('I12').Value
    fcf_y5 = ws_dcf.Range('M12').Value
    reinvest_y1 = ws_dcf.Range('I11').Value
    
    print("=== EVALUATED DCF SHEET ===")
    print(f"Year 1 Reinvestment Rate: {reinvest_y1*100:.1f}%")
    print(f"Year 1 FCFF: Rs {fcf_y1:,.1f} Cr")
    print(f"Year 5 FCFF: Rs {fcf_y5:,.1f} Cr")
    print(f"Enterprise Value: Rs {ev_dcf:,.1f} Cr")
    print(f"Equity Value: Rs {eq_dcf:,.1f} Cr")
    print(f"DCF Intrinsic Value/Share: Rs {iv_dcf:,.2f}")
    print(f"Current Market Price: Rs {cmp_dcf:,.2f}")
    
    ws_sum = wb.Sheets('AI Valuation Summary')
    iv_sum = ws_sum.Range('B5').Value
    cmp_sum = ws_sum.Range('A5').Value
    mos_sum = ws_sum.Range('C5').Value
    verdict_sum = ws_sum.Range('D5').Value
    wacc_sum = ws_sum.Range('E5').Value
    
    print("\n=== EVALUATED AI VALUATION SUMMARY ===")
    print(f"Summary Intrinsic Value: Rs {iv_sum:,.2f}")
    print(f"Summary CMP: Rs {cmp_sum:,.2f}")
    print(f"Margin of Safety: {mos_sum*100:.1f}%")
    print(f"Verdict: {verdict_sum}")
    print(f"WACC: {wacc_sum*100:.2f}%")
    
    ws_comp = wb.Sheets('Comp_Valuation')
    med_ev_rev = ws_comp.Range('O25').Value
    med_ev_ebitda = ws_comp.Range('P25').Value
    print("\n=== EVALUATED COMP_VALUATION ===")
    print(f"Sector Median EV/Rev: {med_ev_rev:.2f}x" if med_ev_rev else "EV/Rev: N/A")
    print(f"Sector Median EV/EBITDA: {med_ev_ebitda:.2f}x" if med_ev_ebitda else "EV/EBITDA: N/A")
    
    wb.Save()
    wb.Close(SaveChanges=True)
finally:
    excel.Quit()
