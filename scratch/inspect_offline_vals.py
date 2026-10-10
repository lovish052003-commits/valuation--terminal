import os, sys, json
sys.path.insert(0, os.path.abspath('.'))
from screener_client import load_offline_company_data
import valuation_engine

for ticker in ['NESTLEIND', 'ITC']:
    d = load_offline_company_data(ticker)
    if not d:
        print(f"No cache for {ticker}")
        continue
    val = valuation_engine.calculate_valuation(d)
    print(f"\n--- {ticker} Valuation Engine Current Output ---")
    print("Intrinsic Value:", val['intrinsic_value_per_share'])
    print("Current Price:", val['current_price'])
    print("WACC:", val['wacc'])
    print("Growth Rate:", val['growth_rate'])
    print("Terminal Growth:", val['terminal_growth'])
    print("EV:", val['enterprise_value'])
    print("Equity Value:", val['equity_value'])
    print("Cash:", val['cash_estimate'])
    print("Debt:", val['total_debt'])
    print("Shares:", val['shares_cr'])
    print("DCF Table Y1 EBIT:", val['dcf_table'][0]['ebit'], "FCFF:", val['dcf_table'][0]['fcff'], "Reinvest:", val['dcf_table'][0]['reinvestment_rate'])
