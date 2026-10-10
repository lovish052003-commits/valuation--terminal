import os, sys
sys.path.insert(0, os.path.abspath('.'))
from screener_client import load_offline_company_data
import valuation_engine

# Let's see what happens if we patch valuation_engine logic temporarily to test
original_calc = valuation_engine.calculate_valuation

def test_patched(ticker):
    d = load_offline_company_data(ticker)
    v = original_calc(d)
    print(f"=== {ticker} CURRENT ===")
    print(f"IV: {v['intrinsic_value_per_share']}, EV: {v['enterprise_value']}, Cash: {v['cash_estimate']}, Debt: {v['total_debt']}, WACC: {v['wacc']}%, g: {v['growth_rate']}%, tg: {v['terminal_growth']}%")

test_patched('NESTLEIND')
test_patched('ITC')
