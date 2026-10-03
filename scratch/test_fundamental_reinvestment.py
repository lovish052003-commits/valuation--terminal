import os
import sys
import json
import numpy as np
import pandas as pd
import math

sys.path.insert(0, os.path.abspath('.'))
from screener_client import clean_num, classify_company

def calculate_fundamental_reinvestment_engine(screener_data, custom_params=None):
    params = {
        'tax_rate': 0.30,
        'terminal_growth': 0.04,
        'growth_rate': None,
        'wacc': None,
        'risk_free_rate': 0.071,
        'erp': 0.055,
    }
    if custom_params:
        params.update({k: v for k, v in custom_params.items() if v is not None})

    company_type = screener_data.get('company_type')
    if not company_type:
        company_type = classify_company(
            screener_data.get('sector', ''),
            screener_data.get('industry', ''),
            screener_data.get('company_name', ''),
            screener_data.get('about', '')
        )

    FINANCIAL_TYPES = {'BANK', 'NBFC', 'INSURANCE', 'ASSET_MANAGEMENT', 'BROKING', 'OTHER_FINANCIAL'}
    is_financial = company_type in FINANCIAL_TYPES

    tables = screener_data.get('tables', {})
    pl_df = tables.get('profit-loss')
    bs_df = tables.get('balance-sheet')
    cf_df = tables.get('cash-flow')

    def extract_metric(df, patterns):
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:
            return [], []
        metric_col = 'Metric' if 'Metric' in df.columns else df.columns[0]
        years = [c for c in df.columns if c != metric_col]
        for pat in patterns:
            m = df[df[metric_col].astype(str).str.contains(pat, case=False, na=False, regex=True)]
            if not m.empty:
                r = m.iloc[0]
                vals = [clean_num(r.get(y, 0)) for y in years]
                return years, vals
        return years, [0.0] * len(years)

    pl_years, sales_hist = extract_metric(pl_df, ['^Sales', 'Revenue', 'Total Income'])
    _, op_profit_hist = extract_metric(pl_df, ['Operating Profit', 'EBITDA', 'Financing Profit'])
    _, depr_hist = extract_metric(pl_df, ['Depreciation'])
    _, interest_hist = extract_metric(pl_df, ['Interest'])
    _, pbt_hist = extract_metric(pl_df, ['Profit before tax', 'PBT'])
    _, tax_hist = extract_metric(pl_df, ['Tax'])
    _, pat_hist = extract_metric(pl_df, ['Net profit', 'PAT'])

    bs_years, eq_cap_hist = extract_metric(bs_df, ['Equity Capital', 'Share Capital'])
    _, reserves_hist = extract_metric(bs_df, ['Reserves'])
    _, borrowings_hist = extract_metric(bs_df, ['Borrowings', 'Total Debt'])
    _, other_liab_hist = extract_metric(bs_df, ['Other Liabilities'])
    _, fixed_assets_hist = extract_metric(bs_df, ['Fixed Assets', 'Net Block'])
    _, other_assets_hist = extract_metric(bs_df, ['Other Assets'])

    # Find common annual periods (exclude TTM)
    ann_pl_years = [y for y in pl_years if 'TTM' not in str(y).upper()]
    common_years = [y for y in ann_pl_years if y in bs_years]

    eff_tax_rate = params['tax_rate']
    reinvestment_warnings = []

    if is_financial:
        reinvestment_warnings.append(f"Financial-company classification detected ({company_type}); FCFF framework not applicable. Valued via Cost of Equity Excess Return Model.")

    valid_roics = []
    historical_roics = []
    valid_reinvest_rates = []
    historical_reinvest_series = []
    historical_nopats = []
    historical_inv_caps = []

    for idx, yr in enumerate(common_years):
        pl_idx = pl_years.index(yr)
        bs_idx = bs_years.index(yr)

        pbt_i = pbt_hist[pl_idx] if pl_idx < len(pbt_hist) else 0.0
        int_i = interest_hist[pl_idx] if pl_idx < len(interest_hist) else 0.0
        op_i = op_profit_hist[pl_idx] if pl_idx < len(op_profit_hist) else 0.0
        depr_i = depr_hist[pl_idx] if pl_idx < len(depr_hist) else 0.0
        tax_i = tax_hist[pl_idx] if pl_idx < len(tax_hist) else 0.0

        if pbt_i != 0 and int_i != 0:
            ebit_i = pbt_i + int_i
        else:
            ebit_i = op_i - depr_i

        if pbt_i > 0 and tax_i > 0:
            t_rate = tax_i / pbt_i
            eff_t = t_rate if 0.05 <= t_rate <= 0.45 else eff_tax_rate
        else:
            eff_t = eff_tax_rate

        nopat_i = ebit_i * (1.0 - eff_t)
        historical_nopats.append({'year': str(yr), 'nopat': round(nopat_i, 2)})

        fa_i = fixed_assets_hist[bs_idx] if bs_idx < len(fixed_assets_hist) else 0.0
        oa_i = other_assets_hist[bs_idx] if bs_idx < len(other_assets_hist) else 0.0
        ol_i = other_liab_hist[bs_idx] if bs_idx < len(other_liab_hist) else 0.0
        wc_i = max(0.0, oa_i - ol_i)
        inv_cap_i = fa_i + wc_i
        historical_inv_caps.append({'year': str(yr), 'invested_capital': round(inv_cap_i, 2)})

        # ROIC calculation & data quality rules
        if inv_cap_i > 0 and not math.isnan(nopat_i):
            roic_i = nopat_i / inv_cap_i
            if not math.isnan(roic_i) and not math.isinf(roic_i) and abs(roic_i) <= 3.0:
                valid_roics.append(roic_i)
                historical_roics.append({'year': str(yr), 'roic': round(roic_i * 100, 2)})

        # Reinvestment rate calculation
        if idx > 0:
            prev_bs_idx = bs_years.index(common_years[idx - 1])
            prev_fa = fixed_assets_hist[prev_bs_idx] if prev_bs_idx < len(fixed_assets_hist) else 0.0
            prev_oa = other_assets_hist[prev_bs_idx] if prev_bs_idx < len(other_assets_hist) else 0.0
            prev_ol = other_liab_hist[prev_bs_idx] if prev_bs_idx < len(other_liab_hist) else 0.0
            prev_wc = max(0.0, prev_oa - prev_ol)

            capex_i = max(0.0, fa_i - prev_fa + depr_i)
            delta_wc_i = wc_i - prev_wc
            reinvest_i = capex_i + delta_wc_i

            if nopat_i > 0:
                rr_i = reinvest_i / nopat_i
                if not math.isnan(rr_i) and not math.isinf(rr_i) and -2.0 <= rr_i <= 5.0:
                    valid_reinvest_rates.append(rr_i)
                    historical_reinvest_series.append({'year': str(yr), 'reinvestment_rate': round(rr_i * 100, 2)})

    # Normalized ROIC
    if valid_roics:
        normalized_roic = float(np.median(valid_roics))
    else:
        normalized_roic = 0.0

    # Historical Median Reinvestment Rate (diagnostic reference)
    if valid_reinvest_rates:
        historical_median_reinvestment_rate = float(np.median(valid_reinvest_rates))
    else:
        historical_median_reinvestment_rate = None

    # Quality and Volatility Diagnostics
    if len(valid_roics) < 3:
        reinvestment_warnings.append(f"Fewer than 3 valid historical ROIC observations available ({len(valid_roics)} found); confidence is LOW.")
        reinvestment_confidence = "LOW"
    else:
        roic_std = float(np.std(valid_roics))
        if roic_std > 0.15:
            reinvestment_warnings.append(f"Historical ROIC is highly volatile across periods (std dev: {roic_std*100:.1f}%).")
        reinvestment_confidence = "HIGH" if len(valid_roics) >= 5 and roic_std < 0.10 else "MEDIUM"

    # Expected Growth Engine
    if params['growth_rate'] is not None:
        expected_growth_rate = float(params['growth_rate'])
        growth_source = "Analyst Forecast"
    elif normalized_roic > 0 and historical_median_reinvestment_rate and 0.15 <= historical_median_reinvestment_rate <= 0.85:
        fund_g = normalized_roic * historical_median_reinvestment_rate
        if 0.02 <= fund_g <= 0.25:
            expected_growth_rate = round(fund_g, 4)
            growth_source = "Fundamental Estimate"
        else:
            # Fallback to Sales CAGR
            if len(sales_hist) >= 4 and sales_hist[-4] > 0 and sales_hist[-1] > 0:
                cagr_3 = (sales_hist[-1] / sales_hist[-4]) ** (1.0 / 3.0) - 1.0
                expected_growth_rate = round(float(np.clip(cagr_3, 0.03, 0.18)), 4)
                growth_source = "Historical Normalized"
            else:
                expected_growth_rate = 0.052
                growth_source = "Fallback Estimate"
    elif len(sales_hist) >= 4 and sales_hist[-4] > 0 and sales_hist[-1] > 0:
        cagr_3 = (sales_hist[-1] / sales_hist[-4]) ** (1.0 / 3.0) - 1.0
        expected_growth_rate = round(float(np.clip(cagr_3, 0.03, 0.18)), 4)
        growth_source = "Historical Normalized"
    else:
        expected_growth_rate = 0.052
        growth_source = "Fallback Estimate"

    # Fundamental Reinvestment Rate
    if normalized_roic <= 0:
        # Case C: Negative / Zero ROIC
        reinvestment_methodology = "Normalized Historical Reinvestment (ROIC Non-Positive Fallback)"
        if historical_median_reinvestment_rate and 0.15 <= historical_median_reinvestment_rate <= 0.85:
            fundamental_reinvestment_rate = historical_median_reinvestment_rate
        else:
            fundamental_reinvestment_rate = 0.35
        reinvestment_confidence = "LOW"
        reinvestment_warnings.append(f"Normalized ROIC is non-positive ({normalized_roic*100:.1f}%); fundamental reinvestment methodology unavailable.")
    elif expected_growth_rate < 0:
        # Case D: Negative Growth
        raw_rr = expected_growth_rate / normalized_roic
        fundamental_reinvestment_rate = raw_rr
        reinvestment_warnings.append(f"Expected growth is negative ({expected_growth_rate*100:.1f}%); implied reinvestment rate is negative ({raw_rr*100:.1f}%).")
        reinvestment_confidence = "LOW"
        reinvestment_methodology = "Growth / Sustainable ROIC (Negative Growth)"
    else:
        raw_rr = expected_growth_rate / normalized_roic
        if raw_rr > 1.0:
            # Case B: Implied Reinvestment > 100%
            fundamental_reinvestment_rate = raw_rr
            reinvestment_confidence = "LOW"
            reinvestment_methodology = "Growth / Sustainable ROIC"
            reinvestment_warnings.append(f"Implied reinvestment rate exceeds 100% ({raw_rr*100:.1f}%); expected growth ({expected_growth_rate*100:.1f}%) is higher than normalized ROIC ({normalized_roic*100:.1f}%).")
        else:
            # Case A: Normal (0% to 100%)
            fundamental_reinvestment_rate = raw_rr
            reinvestment_methodology = "Growth / Sustainable ROIC"

    # Terminal ROIC & Terminal Reinvestment Rate
    terminal_g = params['terminal_growth']
    wacc_est = 0.10  # default WACC baseline for terminal state convergence
    if normalized_roic > wacc_est:
        terminal_roic = min(0.18, max(wacc_est, 0.5 * normalized_roic + 0.5 * wacc_est))
    elif 0 < normalized_roic <= wacc_est:
        terminal_roic = max(wacc_est, 0.10)
    else:
        terminal_roic = max(wacc_est, 0.10)
        reinvestment_warnings.append("Terminal ROIC assumption requires review; normalized ROIC was non-positive.")

    terminal_reinvest_rate = terminal_g / terminal_roic
    terminal_reinvest_rate = float(np.clip(terminal_reinvest_rate, 0.05, 0.85))

    # Consistency Check
    implied_g = fundamental_reinvestment_rate * normalized_roic
    growth_roic_consistency = "PASS" if abs(implied_g - expected_growth_rate) <= 0.005 else "WARNING"

    # 5-Year DCF Forecast Schedule Transition
    forecast_reinvest_rates = []
    for yr in range(1, 6):
        proj_rr = fundamental_reinvestment_rate + (terminal_reinvest_rate - fundamental_reinvestment_rate) * ((yr - 1) / 4.0)
        forecast_reinvest_rates.append(round(proj_rr * 100, 2))

    return {
        'ticker': screener_data.get('ticker'),
        'company_name': screener_data.get('company_name'),
        'company_type': company_type,
        'normalized_roic': round(normalized_roic * 100, 2),
        'expected_growth_rate': round(expected_growth_rate * 100, 2),
        'growth_source': growth_source,
        'fundamental_reinvestment_rate': round(fundamental_reinvestment_rate * 100, 2),
        'historical_median_reinvestment_rate': round(historical_median_reinvestment_rate * 100, 2) if historical_median_reinvestment_rate else None,
        'terminal_roic': round(terminal_roic * 100, 2),
        'terminal_growth_rate': round(terminal_g * 100, 2),
        'terminal_reinvestment_rate': round(terminal_reinvest_rate * 100, 2),
        'reinvestment_methodology': reinvestment_methodology,
        'reinvestment_confidence': reinvestment_confidence,
        'growth_roic_consistency': growth_roic_consistency,
        'reinvestment_warnings': reinvestment_warnings,
        'forecast_reinvest_rates': forecast_reinvest_rates,
        'historical_roics': historical_roics,
    }

if __name__ == '__main__':
    test_tickers = ['TATASTEEL', 'NESTLEIND', 'INFY', 'UNITEDTEA', 'OISL', 'SBIN']
    for t in test_tickers:
        p = os.path.join('exports', '.screener_cache', f'{t}_data.json')
        if not os.path.exists(p):
            continue
        with open(p, 'r', encoding='utf-8') as f:
            data = json.load(f)
        # convert tables if dict
        if 'tables' in data:
            for tb_k, tb_v in data['tables'].items():
                if isinstance(tb_v, dict) and 'columns' in tb_v and 'data' in tb_v:
                    data['tables'][tb_k] = pd.DataFrame(tb_v['data'], columns=tb_v['columns'])
        res = calculate_fundamental_reinvestment_engine(data)
        print(f"\n==========================================")
        print(f"Ticker: {res['ticker']} ({res['company_name']}) | Type: {res['company_type']}")
        print(f"Normalized ROIC: {res['normalized_roic']}%")
        print(f"Expected Growth: {res['expected_growth_rate']}% (Source: {res['growth_source']})")
        print(f"Fundamental Reinvestment Rate: {res['fundamental_reinvestment_rate']}%")
        print(f"Historical Median Reinvestment Rate: {res['historical_median_reinvestment_rate']}%")
        print(f"Terminal ROIC: {res['terminal_roic']}% | Terminal Growth: {res['terminal_growth_rate']}% | Terminal Reinvest: {res['terminal_reinvestment_rate']}%")
        print(f"5-Year Forecast Reinvest Rates (Y1-Y5): {res['forecast_reinvest_rates']}")
        print(f"Methodology: {res['reinvestment_methodology']}")
        print(f"Confidence: {res['reinvestment_confidence']} | Consistency: {res['growth_roic_consistency']}")
        if res['reinvestment_warnings']:
            print(f"Warnings: {res['reinvestment_warnings']}")
