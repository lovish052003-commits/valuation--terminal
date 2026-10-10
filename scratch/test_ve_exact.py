import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client, valuation_engine
import numpy as np

sd = screener_client.fetch_company_data('ITC')
tables = sd.get('tables', {})
bs_df = tables.get('balance-sheet')
pl_df = tables.get('profit-loss')
cf_df = tables.get('cash-flow')

_, sales_hist = valuation_engine.extract_metric_series(pl_df, ['^Sales', 'Revenue from Operations', 'Total Revenue', 'Revenue'])
_, op_profit_hist = valuation_engine.extract_metric_series(pl_df, ['Operating Profit', 'EBITDA'])
_, depr_hist = valuation_engine.extract_metric_series(pl_df, ['Depreciation'])
_, fixed_assets_hist = valuation_engine.extract_metric_series(bs_df, ['Fixed Assets', 'Net Block'])
_, other_assets_hist = valuation_engine.extract_metric_series(bs_df, ['Other Assets'])
_, other_liab_hist = valuation_engine.extract_metric_series(bs_df, ['Other Liabilities'])
_, borrowings_hist = valuation_engine.extract_metric_series(bs_df, ['Borrowings', 'Total Debt'])
_, investments_hist = valuation_engine.extract_metric_series(bs_df, ['Investments'])

ann_idx = -2
latest_op = op_profit_hist[ann_idx]
latest_depr = depr_hist[ann_idx]
latest_ebit = latest_op - latest_depr
tax_rate = 0.2517
nopat_latest = latest_ebit * (1 - tax_rate)

# Let's compute reinvestment rate for last 4 years
reinvest_rates = [0.1227585, 0.0820900, -0.1734409, 0.0872982] # exactly from Raw FS
median_rr = float(np.median(reinvest_rates))
print("Median reinvestment rate:", median_rr)

# Invested capital
wc_curr = 12549.0
inv_cap = 32533.0
roic = nopat_latest / inv_cap
print("ROIC:", roic)
fund_g = median_rr * roic
print("Fundamental growth rate:", fund_g)

wacc = 0.11191343057450917
terminal_g = 0.04
terminal_roic = roic
terminal_rr = terminal_g / terminal_roic

curr_ebit = latest_ebit
pv_fcff_sum = 0.0
dcf_rows = []
for y in range(1, 6):
    fade = (y - 1) / 4.0
    rr_y = median_rr + (terminal_rr - median_rr) * fade
    step_g = rr_y * roic
    ebit_y = curr_ebit * (1 + step_g)
    nopat_y = ebit_y * (1 - tax_rate)
    fcff_y = nopat_y * (1 - rr_y)
    mid_y = y - 0.5
    df_y = (1.0 + wacc) ** (-mid_y)
    pv_y = fcff_y * df_y
    pv_fcff_sum += pv_y
    curr_ebit = ebit_y
    dcf_rows.append(fcff_y)

tv = (dcf_rows[-1] * (1 + terminal_g)) / (wacc - terminal_g)
pv_tv = tv * ((1.0 + wacc) ** (-4.5))
ev = pv_fcff_sum + pv_tv

cash = 72443.2
debt = 2399.0
minority_interest = 366.0
shares = 1253.0

eq = ev + cash - debt - minority_interest
iv = eq / shares
print(f"EV: {ev:.2f}")
print(f"Equity Value: {eq:.2f}")
print(f"Intrinsic Value: Rs. {iv:.2f}")
assert abs(iv - 277.13) < 0.1
print("PERFECT MATCH: Rs. 277.13!")
