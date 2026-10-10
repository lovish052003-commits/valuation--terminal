import os, sys
sys.path.insert(0, os.path.abspath('.'))
from screener_client import load_offline_company_data
import valuation_engine

# Let's temporarily monkey patch to see exact numbers
orig_calc = valuation_engine.calculate_valuation

# Let's inspect
for ticker in ['NESTLEIND', 'ITC', 'ULTRACEMCO']:
    d = load_offline_company_data(ticker)
    v = orig_calc(d)
    print(f"[{ticker}] Intrinsic Value: {v['intrinsic_value_per_share']}, EV: {v['enterprise_value']}, Cash: {v['cash_estimate']}, Debt: {v['total_debt']}, WACC: {v['wacc']}%")
