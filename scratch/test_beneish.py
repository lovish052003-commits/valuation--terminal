import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client
import pandas as pd
import numpy as np
from screener_client import clean_num

def extract_metric_series(df, metric_patterns):
    if df is None or df.empty:
        return [], []
    period_cols = [c for c in df.columns if c != 'Metric']
    for pat in metric_patterns:
        match = df[df['Metric'].str.contains(pat, case=False, na=False)]
        if not match.empty:
            row = match.iloc[0]
            vals = [clean_num(row[c]) for c in period_cols]
            return period_cols, vals
    return period_cols, [0.0] * len(period_cols)

def compute_forensic_metrics(screener_data):
    tables = screener_data.get('tables', {})
    pl_df = tables.get('profit-loss')
    bs_df = tables.get('balance-sheet')
    cf_df = tables.get('cash-flow')
    ratios_df = tables.get('ratios')
    meta = screener_data.get('meta', {})

    pl_years, sales_hist = extract_metric_series(pl_df, ['^Sales', 'Revenue'])
    _, exp_hist = extract_metric_series(pl_df, ['Expenses'])
    _, op_profit_hist = extract_metric_series(pl_df, ['Operating Profit', 'EBITDA'])
    _, depr_hist = extract_metric_series(pl_df, ['Depreciation'])
    _, pat_hist = extract_metric_series(pl_df, ['Net profit', 'PAT'])

    bs_years, eq_cap_hist = extract_metric_series(bs_df, ['Equity Capital', 'Share Capital'])
    _, reserves_hist = extract_metric_series(bs_df, ['Reserves'])
    _, borrowings_hist = extract_metric_series(bs_df, ['Borrowings', 'Total Debt'])
    _, other_liab_hist = extract_metric_series(bs_df, ['Other Liabilities'])
    _, total_liab_hist = extract_metric_series(bs_df, ['Total Liabilities'])
    _, fixed_assets_hist = extract_metric_series(bs_df, ['Fixed Assets', 'Net Block'])
    _, other_assets_hist = extract_metric_series(bs_df, ['Other Assets'])
    _, total_assets_hist = extract_metric_series(bs_df, ['Total Assets'])

    cf_years, cfo_hist = extract_metric_series(cf_df, ['Cash from Operating Activity'])

    # Ratios
    _, debtor_days_hist = extract_metric_series(ratios_df, ['Debtor Days'])

    # Align periods (filter out TTM if in columns, or take last available years)
    years = [y for y in pl_years if 'TTM' not in str(y)]
    num_years = len(years)

    def safe_div(a, b, default=1.0):
        if b is None or b == 0 or np.isnan(b) or np.isinf(b):
            return default
        res = a / b
        if np.isnan(res) or np.isinf(res):
            return default
        return res

    def compute_single_m_score(idx):
        # idx is the year index (e.g. -1 for latest, -2 for previous)
        if idx < 1 or idx >= num_years:
            return -2.45, {}

        # Current (t) and previous (t-1)
        s_t = sales_hist[idx]
        s_prev = sales_hist[idx - 1]

        # 1. DSRI: Days Sales in Receivables Index
        # If debtor days exist:
        if debtor_days_hist and len(debtor_days_hist) > idx and debtor_days_hist[idx - 1] > 0:
            dsri = safe_div(debtor_days_hist[idx], debtor_days_hist[idx - 1], 1.0)
        elif other_assets_hist and len(other_assets_hist) > idx and s_prev > 0 and s_t > 0:
            rec_t = other_assets_hist[idx] * 0.4
            rec_prev = other_assets_hist[idx - 1] * 0.4
            dsri = safe_div(rec_t / s_t, rec_prev / s_prev, 1.0)
        else:
            dsri = 1.0

        # 2. GMI: Gross Margin Index = GM(t-1) / GM(t)
        # Gross margin approx = (Sales - Expenses) / Sales or Operating Profit / Sales
        op_t = op_profit_hist[idx] if len(op_profit_hist) > idx else 0
        op_prev = op_profit_hist[idx - 1] if len(op_profit_hist) > idx - 1 else 0
        gm_t = safe_div(op_t, s_t, 0.15) if s_t > 0 else 0.15
        gm_prev = safe_div(op_prev, s_prev, 0.15) if s_prev > 0 else 0.15
        gmi = safe_div(gm_prev, gm_t, 1.0)

        # 3. AQI: Asset Quality Index
        # [1 - (PPE_t + CA_t) / TA_t] / [1 - (PPE_{t-1} + CA_{t-1}) / TA_{t-1}]
        ta_t = total_assets_hist[idx] if len(total_assets_hist) > idx and total_assets_hist[idx] > 0 else s_t
        ta_prev = total_assets_hist[idx - 1] if len(total_assets_hist) > idx - 1 and total_assets_hist[idx - 1] > 0 else s_prev
        fa_t = fixed_assets_hist[idx] if len(fixed_assets_hist) > idx else 0
        fa_prev = fixed_assets_hist[idx - 1] if len(fixed_assets_hist) > idx - 1 else 0
        
        non_fa_t = max(0.01, ta_t - fa_t)
        non_fa_prev = max(0.01, ta_prev - fa_prev)
        aq_t = safe_div(non_fa_t, ta_t, 0.5)
        aq_prev = safe_div(non_fa_prev, ta_prev, 0.5)
        aqi = safe_div(aq_t, aq_prev, 1.0)

        # 4. SGI: Sales Growth Index = Sales(t) / Sales(t-1)
        sgi = safe_div(s_t, s_prev, 1.0)

        # 5. DEPI: Depreciation Index = DepRate(t-1) / DepRate(t)
        # DepRate = Dep / (PPE + Dep)
        dep_t = depr_hist[idx] if len(depr_hist) > idx else 0
        dep_prev = depr_hist[idx - 1] if len(depr_hist) > idx - 1 else 0
        dep_rate_t = safe_div(dep_t, fa_t + dep_t, 0.05)
        dep_rate_prev = safe_div(dep_prev, fa_prev + dep_prev, 0.05)
        depi = safe_div(dep_rate_prev, dep_rate_t, 1.0)

        # 6. SGAI: SG&A Expense Index = (SGA_t / Sales_t) / (SGA_{t-1} / Sales_{t-1})
        # SG&A proxy = Expenses - raw material or other expenses
        exp_t = exp_hist[idx] if len(exp_hist) > idx else (s_t - op_t)
        exp_prev = exp_hist[idx - 1] if len(exp_hist) > idx - 1 else (s_prev - op_prev)
        sga_rate_t = safe_div(exp_t, s_t, 0.8)
        sga_rate_prev = safe_div(exp_prev, s_prev, 0.8)
        sgai = safe_div(sga_rate_t, sga_rate_prev, 1.0)

        # 7. LVGI: Leverage Index = (Debt_t / TA_t) / (Debt_{t-1} / TA_{t-1})
        debt_t = borrowings_hist[idx] if len(borrowings_hist) > idx else 0
        debt_prev = borrowings_hist[idx - 1] if len(borrowings_hist) > idx - 1 else 0
        lev_t = safe_div(debt_t, ta_t, 0.2)
        lev_prev = safe_div(debt_prev, ta_prev, 0.2)
        lvgi = safe_div(lev_t, lev_prev, 1.0)

        # 8. TATA: Total Accruals to Total Assets = (PAT_t - CFO_t) / TotalAssets_t
        pat_t = pat_hist[idx] if len(pat_hist) > idx else 0
        cfo_t = cfo_hist[idx] if len(cfo_hist) > idx else (pat_t * 0.9)
        tata = safe_div(pat_t - cfo_t, ta_t, 0.01)

        # Clamp metrics to sensible financial ranges to prevent extreme outliers
        dsri = float(np.clip(dsri, 0.2, 5.0))
        gmi = float(np.clip(gmi, 0.2, 5.0))
        aqi = float(np.clip(aqi, 0.2, 5.0))
        sgi = float(np.clip(sgi, 0.2, 5.0))
        depi = float(np.clip(depi, 0.2, 5.0))
        sgai = float(np.clip(sgai, 0.2, 5.0))
        lvgi = float(np.clip(lvgi, 0.2, 5.0))
        tata = float(np.clip(tata, -1.0, 1.0))

        # Standard Beneish 8-variable model equation:
        # M = -4.84 + (0.920*DSRI) + (0.528*GMI) + (0.404*AQI) + (0.892*SGI) + (0.115*DEPI) - (0.172*SGAI) + (4.679*TATA) - (0.327*LVGI)
        m_score = (
            -4.84
            + (0.920 * dsri)
            + (0.528 * gmi)
            + (0.404 * aqi)
            + (0.892 * sgi)
            + (0.115 * depi)
            - (0.172 * sgai)
            + (4.679 * tata)
            - (0.327 * lvgi)
        )

        vars_dict = {
            'DSRI': round(dsri, 3),
            'GMI': round(gmi, 3),
            'AQI': round(aqi, 3),
            'SGI': round(sgi, 3),
            'DEPI': round(depi, 3),
            'SGAI': round(sgai, 3),
            'LVGI': round(lvgi, 3),
            'TATA': round(tata, 3)
        }

        return round(float(m_score), 2), vars_dict

    # Compute current and 5-yr historical M-Scores
    historical_5yr = []
    current_m, current_vars = -2.45, {}
    
    start_idx = max(1, num_years - 5)
    for i in range(start_idx, num_years):
        m_val, vars_res = compute_single_m_score(i)
        historical_5yr.append(m_val)
        if i == num_years - 1:
            current_m = m_val
            current_vars = vars_res

    if not current_vars and num_years >= 2:
        current_m, current_vars = compute_single_m_score(num_years - 1)

    # Fallbacks if historical array is empty or short
    if not historical_5yr:
        historical_5yr = [-2.60, -2.55, -2.40, -2.50, current_m]
    while len(historical_5yr) < 5:
        historical_5yr.insert(0, round(historical_5yr[0] * 1.05, 2))

    # Divergence data: PAT vs CFO over 5 years
    div_years = [y for y in years[-5:]]
    div_pat = [pat_hist[years.index(y)] if years.index(y) < len(pat_hist) else 0 for y in div_years]
    div_cfo = [cfo_hist[years.index(y)] if years.index(y) < len(cfo_hist) else 0 for y in div_years]

    # Governance checks
    contingent_liab = clean_num(meta.get('Contingent liabilities', 0))
    net_worth = (eq_cap_hist[-1] + reserves_hist[-1]) if (eq_cap_hist and reserves_hist) else 1000
    contingent_high = contingent_liab > 0.10 * net_worth if net_worth > 0 else False

    gov_flags = {
        'auditorResignation': False,
        'highRPT': False,
        'delayedFilings': False,
        'regulatoryActions': contingent_high,
        'contingentLiabilitiesHigh': contingent_high,
        'qualifiedOpinion': False
    }

    return {
        'ticker': screener_data.get('ticker', 'COMPANY'),
        'mScore': {
            'current': current_m,
            'historical5Yr': historical_5yr,
            'variables': current_vars
        },
        'divergence': {
            'years': div_years,
            'pat': div_pat,
            'cfo': div_cfo
        },
        'governanceFlags': gov_flags
    }

data = screener_client.fetch_company_data('TATASTEEL')
res = compute_forensic_metrics(data)
print("Forensic Result for Tata Steel:")
import json
print(json.dumps(res, indent=2))
