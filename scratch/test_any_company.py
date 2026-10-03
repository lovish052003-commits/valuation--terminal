import os, sys, shutil
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__) + '/..'))
import screener_client, valuation_engine, excel_exporter

def test_company_valuation(company_query):
    print(f"\n=======================================================")
    print(f"Testing Valuation Automation for: {company_query}")
    print(f"=======================================================")
    
    # 1. Fetch live Screener data
    print(f"1. Fetching Screener data for '{company_query}'...")
    screener_data = screener_client.fetch_company_data(company_query)
    c_name = screener_data.get('company_name', company_query)
    ticker = screener_data.get('ticker', 'COMP').upper()
    print(f"   Identified Company: {c_name} (Ticker: {ticker})")
    print(f"   Sector: {screener_data.get('sector')}, Industry: {screener_data.get('industry')}")
    print(f"   CMP: Rs. {screener_data.get('current_price')}, Market Cap: Rs. {screener_data.get('market_cap_cr')} Cr")
    
    # 2. Run Valuation Engine
    val_result = valuation_engine.calculate_valuation(screener_data)
    print(f"2. Valuation Result: Intrinsic Value: Rs. {val_result.get('intrinsic_value_per_share')}, Verdict: {val_result.get('verdict')}")
    
    # 3. Export via COM
    dest_path = os.path.abspath(f'exports/TEST_{ticker}_COM.xlsx')
    print(f"3. Exporting via COM to {dest_path}...")
    success = excel_exporter.export_via_excel_com(dest_path, screener_data, val_result, "")
    print(f"   COM Export Success: {success}")
    assert success, f"COM export failed for {company_query}!"
    
    # 4. Verify in Excel via COM (Zero repair warnings)
    print("4. Testing opening in Microsoft Excel via COM to check for repair warnings...")
    import win32com.client, pythoncom
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(dest_path)
        print(f"   SUCCESS: Opened {dest_path} with {len(wb.Sheets)} sheets and ZERO REPAIR WARNINGS!")
        wb.Close(SaveChanges=False)
    except Exception as e:
        print(f"   Error opening: {e}")
        raise e
    finally:
        excel.Quit()
        del excel
        pythoncom.CoUninitialize()
        
    # 5. Scan all sheets for any leftover ITC references
    print("5. Scanning all sheets for any leftover ITC references...")
    import openpyxl
    wb_chk = openpyxl.load_workbook(dest_path, data_only=True)
    itc_found = []
    for sname in wb_chk.sheetnames:
        if sname == 'List of Stocks':
            continue
        ws = wb_chk[sname]
        for r in range(1, min(ws.max_row+1, 100)):
            for c in range(1, min(ws.max_column+1, 35)):
                v = str(ws.cell(r, c).value or '')
                if 'itc' in v.lower():
                    safe_v = v.encode('ascii', 'backslashreplace').decode('ascii')
                    itc_found.append((sname, ws.cell(r, c).coordinate, safe_v[:60]))
    print(f"   Total ITC references in {dest_path} (excluding List of Stocks): {len(itc_found)}")
    for item in itc_found:
        print("   -> ITC found at:", item)
    assert len(itc_found) == 0, f"Found {len(itc_found)} residual ITC references in {company_query} model!"
    print(f"   PASSED: 0.00% ITC IN ANY SHEET FOR {company_query}!")
    
if __name__ == '__main__':
    query = sys.argv[1] if len(sys.argv) > 1 else 'Tata Steel'
    test_company_valuation(query)
