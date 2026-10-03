import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from screener_client import fetch_company_data
from valuation_engine import calculate_valuation

seq = ["ONGC", "TCS", "ITC", "SBIN", "LICI", "ICICIAMC", "ONGC"]
vals = []

for s in seq:
    sd = fetch_company_data(s)
    v = calculate_valuation(sd)
    vals.append((s, v))
    print(f"{s}: Intrinsic={v.get('intrinsic_value_per_share')} | WACC={v.get('wacc')}")

v_first = vals[0][1]
v_last = vals[6][1]

print("\n--- DETAILED COMPARISON OF ONGC RUN 1 vs RUN 7 ---")
print(f"Intrinsic: Run 1 = {v_first.get('intrinsic_value_per_share')} vs Run 7 = {v_last.get('intrinsic_value_per_share')}")
print(f"WACC: Run 1 = {v_first.get('wacc')} vs Run 7 = {v_last.get('wacc')}")
print(f"Shares: Run 1 = {v_first.get('shares_cr')} vs Run 7 = {v_last.get('shares_cr')}")
print(f"Equity Value: Run 1 = {v_first.get('equity_value')} vs Run 7 = {v_last.get('equity_value')}")
print(f"Enterprise Value: Run 1 = {v_first.get('enterprise_value')} vs Run 7 = {v_last.get('enterprise_value')}")
print(f"PV FCFF Sum: Run 1 = {v_first.get('pv_fcff_sum')} vs Run 7 = {v_last.get('pv_fcff_sum')}")
print(f"PV Terminal: Run 1 = {v_first.get('pv_terminal_value')} vs Run 7 = {v_last.get('pv_terminal_value')}")
print(f"Total Debt: Run 1 = {v_first.get('total_debt')} vs Run 7 = {v_last.get('total_debt')}")
print(f"Cash Estimate: Run 1 = {v_first.get('cash_estimate')} vs Run 7 = {v_last.get('cash_estimate')}")
print(f"Base Reinvest Rate: Run 1 = {v_first.get('base_reinvest_rate')} vs Run 7 = {v_last.get('base_reinvest_rate')}")
print(f"Terminal Reinvest Rate: Run 1 = {v_first.get('terminal_reinvest_rate')} vs Run 7 = {v_last.get('terminal_reinvest_rate')}")

# Check key inputs in valuation_inputs / financial_data
for k in v_first.get('valuation_inputs', {}):
    if v_first['valuation_inputs'][k] != v_last['valuation_inputs'].get(k):
        print(f"valuation_inputs diff: {k} -> {v_first['valuation_inputs'][k]} vs {v_last['valuation_inputs'].get(k)}")

for k in v_first.get('financial_data', {}):
    if v_first['financial_data'][k] != v_last['financial_data'].get(k):
        print(f"financial_data diff: {k} -> {v_first['financial_data'][k]} vs {v_last['financial_data'].get(k)}")
