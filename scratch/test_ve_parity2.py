import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client, valuation_engine

sd = screener_client.fetch_company_data('ITC')
tables = sd.get('tables', {})
bs_df = tables.get('balance-sheet')
pl_df = tables.get('profit-loss')

ann_idx = -2
_, fixed_assets_hist = valuation_engine.extract_metric_series(bs_df, ['Fixed Assets', 'Net Block'])
_, other_assets_hist = valuation_engine.extract_metric_series(bs_df, ['Other Assets'])
_, other_liab_hist = valuation_engine.extract_metric_series(bs_df, ['Other Liabilities'])
_, depr_hist = valuation_engine.extract_metric_series(pl_df, ['Depreciation'])
_, op_profit_hist = valuation_engine.extract_metric_series(pl_df, ['Operating Profit', 'EBITDA'])

latest_depr = depr_hist[ann_idx]
latest_op = op_profit_hist[ann_idx]
latest_ebit = latest_op - latest_depr
nopat_latest = latest_ebit * (1 - 0.2517)

capex = max(0.0, (fixed_assets_hist[-1] - fixed_assets_hist[-2]) + latest_depr)
net_capex_latest = max(0.0, capex - latest_depr)

wc_curr = other_assets_hist[ann_idx] - other_liab_hist[ann_idx]
wc_prev = other_assets_hist[ann_idx - 1] - other_liab_hist[ann_idx - 1]
delta_wc = wc_curr - wc_prev

latest_reinvest = net_capex_latest + delta_wc
latest_rr = latest_reinvest / nopat_latest
print("latest_ebit:", latest_ebit)
print("nopat_latest:", nopat_latest)
print("capex:", capex, "net_capex:", net_capex_latest)
print("wc_curr:", wc_curr, "wc_prev:", wc_prev, "delta_wc:", delta_wc)
print("latest_reinvest:", latest_reinvest)
print("latest_rr:", latest_rr)

invested_capital = fixed_assets_hist[ann_idx] + wc_curr
curr_roic = nopat_latest / invested_capital
print("invested_capital:", invested_capital)
print("curr_roic:", curr_roic)
print("fund_growth:", latest_rr * curr_roic)
