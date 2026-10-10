import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client, valuation_engine

sd = screener_client.fetch_company_data('ITC')

# Let's inspect what valuation_engine extracts for ITC
tables = sd.get('tables', {})
bs_df = tables.get('balance-sheet')
pl_df = tables.get('profit-loss')
cf_df = tables.get('cash-flow')

_, fixed_assets_hist = valuation_engine.extract_metric_series(bs_df, ['Fixed Assets', 'Net Block'])
_, other_assets_hist = valuation_engine.extract_metric_series(bs_df, ['Other Assets'])
_, other_liab_hist = valuation_engine.extract_metric_series(bs_df, ['Other Liabilities'])
_, depr_hist = valuation_engine.extract_metric_series(pl_df, ['Depreciation'])
_, op_profit_hist = valuation_engine.extract_metric_series(pl_df, ['Operating Profit', 'EBITDA'])

print("fixed_assets_hist:", fixed_assets_hist[-3:])
print("other_assets_hist:", other_assets_hist[-3:])
print("other_liab_hist:", other_liab_hist[-3:])
print("depr_hist:", depr_hist[-3:])
print("op_profit_hist:", op_profit_hist[-3:])

# Capex from cash flow statement
_, fa_purchased_hist = valuation_engine.extract_metric_series(cf_df, ['Fixed assets purchased'])
_, fa_sold_hist = valuation_engine.extract_metric_series(cf_df, ['Fixed assets sold'])
print("fa_purchased_hist:", fa_purchased_hist[-3:] if fa_purchased_hist else None)
print("fa_sold_hist:", fa_sold_hist[-3:] if fa_sold_hist else None)

# In Raw FS:
# AD27 is fa_purchased (-2199), AD28 is fa_sold (0), AD10 is depr (1777)
# Net capex = 2199 - 1777 = 422
# Change in WC = (other_assets[-1] - other_liab[-1]) - (other_assets[-2] - other_liab[-2])
wc_curr = other_assets_hist[-1] - other_liab_hist[-1]
wc_prev = other_assets_hist[-2] - other_liab_hist[-2]
print("wc_curr:", wc_curr, "wc_prev:", wc_prev, "delta_wc:", wc_curr - wc_prev)

latest_depr = depr_hist[-1]
if fa_purchased_hist:
    raw_capex = abs(fa_purchased_hist[-1])
else:
    raw_capex = max(0.0, (fixed_assets_hist[-1] - fixed_assets_hist[-2]) + latest_depr)
net_capex = max(0.0, raw_capex - latest_depr)
print("raw_capex:", raw_capex, "latest_depr:", latest_depr, "net_capex:", net_capex)

tot_reinvest = net_capex + (wc_curr - wc_prev)
print("tot_reinvest:", tot_reinvest)

latest_ebit = op_profit_hist[-1] - latest_depr
tax_rate = 0.2517
nopat = latest_ebit * (1 - tax_rate)
rr = tot_reinvest / nopat
print("latest_ebit:", latest_ebit, "nopat:", nopat, "rr:", rr)

invested_capital = fixed_assets_hist[-1] + wc_curr
roic = nopat / invested_capital
print("invested_capital:", invested_capital, "roic:", roic)
print("growth_rate (rr * roic):", rr * roic)
