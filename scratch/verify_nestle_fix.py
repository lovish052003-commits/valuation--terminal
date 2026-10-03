import os, sys, gc
import win32com.client, pythoncom
import pandas as pd
import openpyxl

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

sys.path.insert(0, os.path.abspath('.'))
from screener_client import fetch_company_data
from valuation_engine import calculate_valuation
from excel_exporter import export_valuation_model

def verify_nestle():
    symbol = 'NESTLEIND'
    print(f"1. Fetching live company data for {symbol}...")
    screener_data = fetch_company_data(symbol)
    print(f"   Company: {screener_data.get('company_name')}")
    print(f"   CMP: Rs. {screener_data.get('current_price')}")
    print(f"   Market Cap: Rs. {screener_data.get('market_cap_cr')} Cr")
    print(f"   Shares: {screener_data.get('shares_in_cr'):.2f} Cr")
    print(f"   Historical Price Points: {len(screener_data.get('historical_prices', []))}")

    print("\n2. Calculating valuation...")
    val_result = calculate_valuation(screener_data)
    dcf = val_result.get('dcf', {})
    print(f"   DCF Target Value Per Share: Rs. {dcf.get('target_value_per_share')}")
    print(f"   Valuation Verdict: {val_result.get('recommendation')}")

    print("\n3. Exporting full model to Excel...")
    out_file = export_valuation_model(screener_data, val_result, "### Test AI Summary Report")
    print(f"   Exported successfully to: {out_file}")

    print("\n4. Inspecting exported Data Sheet cells...")
    wb = openpyxl.load_workbook(out_file, data_only=True)
    ws = wb['Data Sheet']

    check_rows = [
        (1, "Company Name"),
        (7, "Face Value"),
        (8, "Current Price"),
        (9, "Market Cap"),
        (17, "Sales"),
        (30, "Net Profit"),
        (31, "Dividend Amount"),
        (42, "Qtr Sales"),
        (57, "Equity Share Capital"),
        (67, "Receivables"),
        (68, "Inventory"),
        (69, "Cash & Bank"),
        (70, "No. of Equity Shares"),
        (71, "New Bonus Shares"),
        (72, "Face value"),
        (75, "Other Assets Plug"),
        (82, "CFO"),
        (90, "PRICE:"),
        (93, "Adjusted Equity Shares in Cr")
    ]

    for r, name in check_rows:
        vals = [ws.cell(r, c).value for c in range(2, 12)]
        print(f"Row {r:2d} ({name:30s}): {vals[:3]} ... {vals[-2:]}")

    wb.close()
    print("\n=== Verification Completed Successfully! ===")

if __name__ == '__main__':
    verify_nestle()
