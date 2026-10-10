import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client

def test_workbook_parity():
    sd = screener_client.fetch_company_data('ITC')
    
    # Let's inspect the exact values from Data Sheet / Raw FS for ITC
    # Base EBIT
    # In ITC Model.xlsx:
    # 'Intrinsic Valuation'!L38: EBIT
    # Let's check what L38 is in ITC Model.xlsx
    import openpyxl
    wb = openpyxl.load_workbook('exports/ITC_Valuation_Model.xlsx', data_only=True)
    ws_dcf = wb['DCF']
    ws_iv = wb['Intrinsic Valuation']
    ws_wacc = wb['WACC']
    
    print("DCF H8 (Base EBIT):", ws_dcf['H8'].value)
    print("DCF I8 (Year 1 EBIT):", ws_dcf['I8'].value)
    print("DCF J8 (Year 2 EBIT):", ws_dcf['J8'].value)
    print("DCF K8 (Year 3 EBIT):", ws_dcf['K8'].value)
    print("DCF L8 (Year 4 EBIT):", ws_dcf['L8'].value)
    print("DCF M8 (Year 5 EBIT):", ws_dcf['M8'].value)
    print("---")
    print("DCF H9 (Tax Rate):", ws_dcf['H9'].value)
    print("DCF I11 (Year 1 Reinvest):", ws_dcf['I11'].value)
    print("DCF M11 (Year 5 Reinvest):", ws_dcf['M11'].value)
    print("DCF D20 (WACC):", ws_dcf['D20'].value)
    print("DCF D19 (Terminal g):", ws_dcf['D19'].value)
    print("DCF D22 (Terminal ROIC):", ws_dcf['D22'].value)
    print("DCF D21 (Terminal Reinvest):", ws_dcf['D21'].value)
    print("DCF D29 (Terminal Value):", ws_dcf['D29'].value)
    print("DCF D33 (PV FCFF):", ws_dcf['D33'].value)
    print("DCF D34 (PV TV):", ws_dcf['D34'].value)
    print("DCF D35 (Operating Assets / EV):", ws_dcf['D35'].value)
    print("DCF D37 (Cash):", ws_dcf['D37'].value)
    print("DCF D38 (Debt):", ws_dcf['D38'].value)
    print("DCF D39 (Minority Interest):", ws_dcf['D39'].value)
    print("DCF D40 (Equity Value):", ws_dcf['D40'].value)
    print("DCF D41 (Shares):", ws_dcf['D41'].value)
    print("DCF D43 (Intrinsic Value):", ws_dcf['D43'].value)

if __name__ == '__main__':
    test_workbook_parity()
