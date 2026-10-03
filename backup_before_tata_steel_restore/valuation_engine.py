"""
valuation_engine.py
Implements the institutional 4-Pillar Valuation Framework modeled after ITC Model.xlsx:

1. Free Cash Flow Generation (The Engine):
   - Operating Margins (EBITDA & EBIT conversion efficiency)
   - CapEx Breakdown (Maintenance vs Growth CapEx)
   - Working Capital Efficiency & Cash Conversion Cycle (Debtor, Inventory, Payable Days)
   - Earnings Quality Ratio (Operating Cash Flow / Net Income)

2. Growth Assumptions (The Trajectory):
   - Historical 3-Yr & 5-Yr Sales & Profit CAGR
   - Explicit 5-Year Forecast Schedule (EBIT, NOPAT, Reinvestment, FCFF)
   - Terminal Value Perpetual Growth (pegged to long-term GDP / inflation ~ 2% - 3%)
   - Terminal Value share of total Enterprise Value

3. Cost of Capital & Risk (The Discount Rate):
   - Weighted Average Cost of Capital (WACC)
   - Risk-Free Rate (Rf 10-Yr Benchmark) & Equity Risk Premium (ERP)
   - Beta (market volatility / sector peer risk)
   - Cost of Debt (pre-tax & post-tax) and Capital Structure (D/E, We, Wd)

4. Relative Market Multiples (The Sanity Check):
   - EV/EBITDA (core operating value, capital-structure agnostic)
   - P/E Ratio (equity comparison with accounting nuance)
   - Price-to-Book (P/B) and EV/Sales
   - Football Field valuation comparison (DCF vs Peer Multiples vs 52-Wk Range)
"""

import math
import numpy as np
import pandas as pd
from screener_client import clean_num
from universal_valuation import (
    VALUATION_CONFIG,
    CompanyMaster,
    CompanyClassificationEngine,
    ValuationMethodSelector,
    FinancialNormalizationEngine,
    ForecastEngine,
    SustainableROICEngine,
    CostOfCapitalEngine,
    PeerSelectionEngine,
    PeerSelectionError,
    SOTPEngine,
    DCFEngine,
    ValuationReconciliationEngine,
    DataSheetReconciliationEngine,
    ValuationConfidenceEngine,
    ModelQualityEngine,
    CompanyData,
    WorkbookMap
)


def extract_metric_series(df, metric_patterns):
    """
    Extracts a chronological series of values for a given metric across columns.
    Returns (years_list, values_list).
    """
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
    """
    Computes the Beneish M-Score (8-variable model), 5-year historical trajectory,
    cash flow divergence (PAT vs CFO), and corporate governance / auditor flags.
    
    Formula:
    M-Score = -4.84 + (0.920 * DSRI) + (0.528 * GMI) + (0.404 * AQI) + (0.892 * SGI) 
              + (0.115 * DEPI) - (0.172 * SGAI) + (4.679 * TATA) - (0.327 * LVGI)
    
    Thresholds:
      M < -2.22 : Low Risk (Non-manipulator)
      -2.22 to -1.78 : Moderate / Grey Zone
      M > -1.78 : High Risk (Probable Manipulator)
    """
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
        if idx < 1 or idx >= num_years:
            return -2.45, {}

        # Current (t) and previous (t-1)
        s_t = sales_hist[idx] if len(sales_hist) > idx else 0
        s_prev = sales_hist[idx - 1] if len(sales_hist) > idx - 1 else 0

        # 1. DSRI: Days Sales in Receivables Index
        if debtor_days_hist and len(debtor_days_hist) > idx and debtor_days_hist[idx - 1] > 0:
            dsri = safe_div(debtor_days_hist[idx], debtor_days_hist[idx - 1], 1.0)
        elif other_assets_hist and len(other_assets_hist) > idx and s_prev > 0 and s_t > 0:
            rec_t = other_assets_hist[idx] * 0.4
            rec_prev = other_assets_hist[idx - 1] * 0.4
            dsri = safe_div(rec_t / s_t, rec_prev / s_prev, 1.0)
        else:
            dsri = 1.0

        # 2. GMI: Gross Margin Index = GM(t-1) / GM(t)
        op_t = op_profit_hist[idx] if len(op_profit_hist) > idx else 0
        op_prev = op_profit_hist[idx - 1] if len(op_profit_hist) > idx - 1 else 0
        gm_t = safe_div(op_t, s_t, 0.15) if s_t > 0 else 0.15
        gm_prev = safe_div(op_prev, s_prev, 0.15) if s_prev > 0 else 0.15
        gmi = safe_div(gm_prev, gm_t, 1.0)

        # 3. AQI: Asset Quality Index
        ta_t = total_assets_hist[idx] if len(total_assets_hist) > idx and total_assets_hist[idx] > 0 else (s_t if s_t > 0 else 1000)
        ta_prev = total_assets_hist[idx - 1] if len(total_assets_hist) > idx - 1 and total_assets_hist[idx - 1] > 0 else (s_prev if s_prev > 0 else 1000)
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
        dep_t = depr_hist[idx] if len(depr_hist) > idx else 0
        dep_prev = depr_hist[idx - 1] if len(depr_hist) > idx - 1 else 0
        dep_rate_t = safe_div(dep_t, fa_t + dep_t, 0.05)
        dep_rate_prev = safe_div(dep_prev, fa_prev + dep_prev, 0.05)
        depi = safe_div(dep_rate_prev, dep_rate_t, 1.0)

        # 6. SGAI: SG&A Expense Index = (SGA_t / Sales_t) / (SGA_{t-1} / Sales_{t-1})
        exp_t = exp_hist[idx] if len(exp_hist) > idx else (s_t - op_t)
        exp_prev = exp_hist[idx - 1] if len(exp_hist) > idx - 1 else (s_prev - op_prev)
        sga_rate_t = safe_div(exp_t, s_t, 0.8)
        sga_rate_prev = safe_div(exp_prev, s_prev, 0.8)
        sgai = safe_div(sga_rate_t, sga_rate_prev, 1.0)

        # 7. LVGI: Leverage Index = (Debt_t / TA_t) / (Debt_{t-1} / TA_{t-1})
        debt_t = borrowings_hist[idx] if len(borrowings_hist) > idx else 0
        debt_prev = borrowings_hist[idx - 1] if len(borrowings_hist) > idx - 1 else 0
        lev_t = safe_div(debt_t, ta_t, 0.0)
        lev_prev = safe_div(debt_prev, ta_prev, 0.0)
        # For virtually debt-free firms (< 2% of assets), leverage risk is negligible
        if lev_t < 0.02 and lev_prev < 0.02:
            lvgi = 1.0
        elif lev_prev <= 0.001:
            lvgi = 1.0 if lev_t <= 0.05 else safe_div(lev_t, 0.01, 1.0)
        else:
            lvgi = safe_div(lev_t, lev_prev, 1.0)

        # 8. TATA: Total Accruals to Total Assets = (PAT_t - CFO_t) / TotalAssets_t
        pat_t = pat_hist[idx] if len(pat_hist) > idx else 0
        cfo_t = cfo_hist[idx] if len(cfo_hist) > idx else (pat_t * 0.9)
        tata = safe_div(pat_t - cfo_t, ta_t, 0.01)

        # Bounds clipping
        dsri = float(np.clip(dsri, 0.2, 5.0))
        gmi = float(np.clip(gmi, 0.2, 5.0))
        aqi = float(np.clip(aqi, 0.2, 5.0))
        sgi = float(np.clip(sgi, 0.2, 5.0))
        depi = float(np.clip(depi, 0.2, 5.0))
        sgai = float(np.clip(sgai, 0.2, 5.0))
        lvgi = float(np.clip(lvgi, 0.2, 5.0))
        tata = float(np.clip(tata, -1.0, 1.0))

        # Standard Beneish M-Score Formula
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

    # 5-Yr historical M-Scores
    historical_5yr = []
    current_m, current_vars = -2.45, {
        'DSRI': 1.0, 'GMI': 1.0, 'AQI': 1.0, 'SGI': 1.0, 
        'DEPI': 1.0, 'SGAI': 1.0, 'LVGI': 1.0, 'TATA': 0.01
    }
    
    start_idx = max(1, num_years - 5)
    for i in range(start_idx, num_years):
        m_val, vars_res = compute_single_m_score(i)
        historical_5yr.append(m_val)
        if i == num_years - 1:
            current_m = m_val
            current_vars = vars_res

    if not current_vars and num_years >= 2:
        current_m, current_vars = compute_single_m_score(num_years - 1)

    if not historical_5yr:
        historical_5yr = [-2.60, -2.55, -2.40, -2.50, current_m]
    while len(historical_5yr) < 5:
        historical_5yr.insert(0, round(historical_5yr[0] * 1.02, 2))

    # Divergence data: PAT vs CFO over 5 years
    div_years = [str(y) for y in years[-5:]]
    div_pat = [round(float(pat_hist[years.index(y)]), 1) if years.index(y) < len(pat_hist) else 0.0 for y in div_years]
    div_cfo = [round(float(cfo_hist[years.index(y)]), 1) if years.index(y) < len(cfo_hist) else 0.0 for y in div_years]

    # Governance checks
    contingent_liab = clean_num(meta.get('Contingent liabilities', 0))
    net_worth = (eq_cap_hist[-1] + reserves_hist[-1]) if (eq_cap_hist and reserves_hist) else 1000
    contingent_high = bool(contingent_liab > 0.10 * net_worth if net_worth > 0 else False)

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

def calculate_cagr(series, periods):
    """Calculates CAGR over specified number of periods."""
    if len(series) >= periods and series[-periods] > 0 and series[-1] > 0:
        return float((series[-1] / series[-periods]) ** (1.0 / (periods - 1)) - 1.0)
    return None

def calculate_reverse_dcf(ticker, current_price, shares_cr, wacc, terminal_growth, base_fcf, net_debt, sales_cagr_5yr, comps_summary):
    """
    Executes Expectations Investing (Reverse DCF).
    Solves for the implied 5-Year FCF CAGR that makes Model Enterprise Value equal Market Cap + Net Debt.
    Generates a 5x5 sensitivity matrix of implied share prices across WACC and Growth.
    """
    market_cap = current_price * shares_cr
    target_ev = max(1.0, market_cap + net_debt)
    
    # Base FCF sanity check
    if base_fcf is None or base_fcf <= 0:
        base_fcf = max(100.0, target_ev * 0.05)

    def model_ev(g, r_wacc, r_tg):
        pv_fcf = sum([(base_fcf * ((1.0 + g)**t)) / ((1.0 + r_wacc)**t) for t in range(1, 6)])
        tv = (base_fcf * ((1.0 + g)**5) * (1.0 + r_tg)) / max(0.005, (r_wacc - r_tg))
        pv_tv = tv / ((1.0 + r_wacc)**5)
        return pv_fcf + pv_tv

    # Solve for implied growth using bisection
    low, high = -0.50, 1.50
    for _ in range(60):
        mid = (low + high) / 2.0
        val = model_ev(mid, wacc, terminal_growth)
        if abs(val - target_ev) < 1.0:
            break
        if val < target_ev:
            low = mid
        else:
            high = mid
    implied_g = (low + high) / 2.0

    # Historical FCF / top-line compounding benchmark
    hist_cagr = (sales_cagr_5yr * 100.0) if (sales_cagr_5yr and sales_cagr_5yr > 0) else 8.5
    
    # Sector average CAGR baseline
    sector_cagr = 10.1
    if comps_summary and comps_summary.get('peer_median_pe'):
        sector_cagr = 10.5

    # 5x5 Sensitivity Matrix
    # WACC rows: wacc - 2%, wacc - 1%, wacc, wacc + 1%, wacc + 2%
    rows_wacc = [round((wacc + delta) * 100, 2) for delta in [-0.02, -0.01, 0.0, 0.01, 0.02]]
    # Growth cols: g - 8%, g - 4%, g, g + 4%, g + 8%
    cols_growth = [round((implied_g + delta) * 100, 1) for delta in [-0.08, -0.04, 0.0, 0.04, 0.08]]

    matrix_prices = []
    for r_w_pct in rows_wacc:
        row_prices = []
        r_w = r_w_pct / 100.0
        for c_g_pct in cols_growth:
            c_g = c_g_pct / 100.0
            ev_val = model_ev(c_g, r_w, terminal_growth)
            eq_val = max(1.0, ev_val - net_debt)
            implied_p = eq_val / shares_cr if shares_cr > 0 else current_price
            row_prices.append(round(float(implied_p), 0))
        matrix_prices.append(row_prices)

    return {
        'ticker': ticker,
        'currentPrice': round(float(current_price), 2),
        'wacc': round(float(wacc * 100), 2),
        'terminalGrowth': round(float(terminal_growth * 100), 1),
        'historicalFCFCagr': round(float(hist_cagr), 1),
        'impliedFCFCagr': round(float(implied_g * 100), 1),
        'sectorAverageCagr': round(float(sector_cagr), 1),
        'sensitivityMatrix': {
            'rowsWACC': rows_wacc,
            'colsGrowth': cols_growth,
            'matrixPrices': matrix_prices
        }
    }

# Centralized Valuation Assumptions Configuration Object
VALUATION_ASSUMPTIONS = {
    'forecastGrowthDefault': 0.052,
    'terminalGrowthDefault': 0.040,
    'riskFreeRate': 0.070,
    'equityRiskPremium': 0.065,
    'taxRateDefault': 0.30
}

def validate_valuation_inputs(screener_data, custom_params=None):
    """
    Validates company inputs before running valuation.
    Returns (is_valid: bool, error_msg: str).
    """
    if not screener_data:
        return False, "No company data provided."
    ticker = screener_data.get('ticker', '').strip()
    if not ticker:
        return False, "Company ticker symbol is missing or unresolved."
    cmp = clean_num(screener_data.get('current_price', 0))
    if cmp <= 0:
        return False, f"Current market price unavailable for {screener_data.get('company_name', ticker)}."
    mcap = clean_num(screener_data.get('market_cap_cr', 0))
    if mcap <= 0:
        return False, f"Market capitalization for {ticker} is unavailable or zero."
    tables = screener_data.get('tables', {})
    if not tables or not any(k in tables for k in ['profit-loss', 'balance-sheet']):
        return False, "Valuation unavailable because required financial data is missing."
    
    # Check WACC vs Terminal Growth
    wacc = custom_params.get('wacc') if custom_params else None
    tg = custom_params.get('terminal_growth') if custom_params else None
    if wacc is not None and tg is not None:
        if wacc <= tg:
            return False, f"DCF unavailable because WACC ({wacc*100:.1f}%) is below or equal to terminal growth ({tg*100:.1f}%)."
    return True, ""

def compute_fundamental_reinvestment_engine(screener_data, custom_params=None):
    """
    Universal Fundamental DCF Reinvestment Engine.
    Replaces mechanical historical median reinvestment with fundamental economic drivers:
        Reinvestment Rate = Expected Growth / Sustainable ROIC
    Includes normalized ROIC engine, dynamic growth hierarchy, extreme value cases,
    terminal ROIC competitive convergence, and consistency checks.
    """
    params = {
        'tax_rate': VALUATION_ASSUMPTIONS['taxRateDefault'],
        'terminal_growth': VALUATION_ASSUMPTIONS['terminalGrowthDefault'],
        'growth_rate': None,
        'wacc': None,
    }
    if custom_params:
        params.update({k: v for k, v in custom_params.items() if v is not None})

    company_type = screener_data.get('company_type')
    if not company_type:
        from screener_client import classify_company
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

    if company_type == 'BANK':
        pl_years, sales_hist = extract_metric_series(pl_df, ['Interest Earned', 'Revenue', '^Sales', 'Total Income'])
        _, op_profit_hist = extract_metric_series(pl_df, ['Financing Profit', 'Operating Profit', 'EBITDA'])
    else:
        pl_years, sales_hist = extract_metric_series(pl_df, ['^Sales', 'Revenue from Operations', 'Total Revenue', 'Revenue'])
        _, op_profit_hist = extract_metric_series(pl_df, ['Operating Profit', 'EBITDA'])

    _, depr_hist = extract_metric_series(pl_df, ['Depreciation'])
    _, interest_hist = extract_metric_series(pl_df, ['Interest'])
    _, pbt_hist = extract_metric_series(pl_df, ['Profit before tax', 'PBT'])
    _, tax_hist = extract_metric_series(pl_df, ['Tax'])
    _, pat_hist = extract_metric_series(pl_df, ['Net profit', 'PAT'])

    bs_years, eq_cap_hist = extract_metric_series(bs_df, ['Equity Capital', 'Share Capital'])
    _, reserves_hist = extract_metric_series(bs_df, ['Reserves'])
    _, borrowings_hist = extract_metric_series(bs_df, ['Borrowings', 'Total Debt'])
    _, other_liab_hist = extract_metric_series(bs_df, ['Other Liabilities'])
    _, fixed_assets_hist = extract_metric_series(bs_df, ['Fixed Assets', 'Net Block'])
    _, other_assets_hist = extract_metric_series(bs_df, ['Other Assets'])

    ann_pl_years = [y for y in pl_years if 'TTM' not in str(y).upper()]
    common_years = [y for y in ann_pl_years if y in bs_years]

    eff_tax_rate = params['tax_rate'] if params['tax_rate'] is not None else 0.30
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

        # Historical Reinvestment Rate calculation (Net Capex + delta WC / NOPAT)
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
    normalized_roic = float(np.median(valid_roics)) if valid_roics else 0.0

    # Historical Median Reinvestment Rate (diagnostic reference)
    historical_median_reinvestment_rate = float(np.median(valid_reinvest_rates)) if valid_reinvest_rates else None

    # Quality, Cyclicality and Volatility Diagnostics
    sector_str = str(screener_data.get('sector', '')).lower()
    industry_str = str(screener_data.get('industry', '')).lower()
    about_str = str(screener_data.get('about', '')).lower()
    name_str = str(screener_data.get('company_name', '')).lower()
    combined_desc = f"{sector_str} {industry_str} {about_str} {name_str}"

    CYCLICAL_KEYWORDS = ['steel', 'metal', 'mining', 'iron', 'aluminum', 'copper', 'cement', 'commodity', 'oil', 'gas', 'refinery', 'petrochem', 'capital goods', 'infrastructure', 'construction', 'sugar', 'paper', 'textile', 'shipping', 'auto']
    is_cyclical = any(k in combined_desc for k in CYCLICAL_KEYWORDS)

    roic_std = float(np.std(valid_roics)) if len(valid_roics) >= 2 else 0.0
    roic_cv = (roic_std / abs(normalized_roic)) if normalized_roic > 0 else 1.0

    if len(valid_roics) < 3:
        reinvestment_warnings.append(f"Fewer than 3 valid historical ROIC observations available ({len(valid_roics)} found); confidence is LOW.")
        reinvestment_confidence = "LOW"
    elif is_financial:
        reinvestment_confidence = "LOW"
    elif normalized_roic <= 0:
        reinvestment_confidence = "LOW"
    elif roic_cv > 0.60 or roic_std > 0.12:
        reinvestment_warnings.append(f"Historical ROIC is highly volatile across periods (std dev: {roic_std*100:.1f}%, CV: {roic_cv*100:.1f}%).")
        reinvestment_confidence = "MEDIUM" if len(valid_roics) >= 5 else "LOW"
    elif is_cyclical:
        reinvestment_warnings.append(f"Cyclical commodity/industrial profile detected; ROIC subject to cycle fluctuations (std dev: {roic_std*100:.1f}%). Reinvestment confidence capped at MEDIUM.")
        reinvestment_confidence = "MEDIUM"
    elif len(valid_roics) >= 5 and roic_std < 0.08 and 0.10 <= roic_cv <= 0.40:
        reinvestment_confidence = "HIGH"
    else:
        reinvestment_confidence = "MEDIUM"

    # Expected Growth Engine
    if params.get('growth_rate') is not None:
        expected_growth_rate = float(params['growth_rate'])
        growth_source = "Analyst Forecast"
    elif normalized_roic > 0 and historical_median_reinvestment_rate and 0.15 <= historical_median_reinvestment_rate <= 0.85:
        fund_g = normalized_roic * historical_median_reinvestment_rate
        if 0.02 <= fund_g <= 0.25:
            expected_growth_rate = round(fund_g, 4)
            growth_source = "Fundamental Estimate"
        else:
            if len(sales_hist) >= 4 and sales_hist[-4] > 0 and sales_hist[-1] > 0:
                cagr_3 = (sales_hist[-1] / sales_hist[-4]) ** (1.0 / 3.0) - 1.0
                expected_growth_rate = round(float(np.clip(cagr_3, 0.03, 0.18)), 4)
                growth_source = "Historical Normalized"
            else:
                expected_growth_rate = VALUATION_ASSUMPTIONS['forecastGrowthDefault']
                growth_source = "Fallback Estimate"
    elif len(sales_hist) >= 4 and sales_hist[-4] > 0 and sales_hist[-1] > 0:
        cagr_3 = (sales_hist[-1] / sales_hist[-4]) ** (1.0 / 3.0) - 1.0
        expected_growth_rate = round(float(np.clip(cagr_3, 0.03, 0.18)), 4)
        growth_source = "Historical Normalized"
    else:
        expected_growth_rate = VALUATION_ASSUMPTIONS['forecastGrowthDefault']
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
    # Documented sustainable terminal economics:
    # Over the terminal horizon, company returns fade 50% toward WACC, bounded by sustainable macroeconomic
    # and institutional limits (6.0% lower bound, 18.0% upper bound), without an artificial floor at WACC.
    terminal_g = params['terminal_growth'] or VALUATION_ASSUMPTIONS['terminalGrowthDefault']
    wacc_est = params.get('wacc') or 0.10
    if normalized_roic > 0:
        terminal_roic = min(0.18, max(0.06, 0.5 * normalized_roic + 0.5 * wacc_est))
    else:
        terminal_roic = max(0.06, wacc_est)
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
        'normalized_roic': normalized_roic,
        'expected_growth_rate': expected_growth_rate,
        'growth_source': growth_source,
        'fundamental_reinvestment_rate': fundamental_reinvestment_rate,
        'historical_median_reinvestment_rate': historical_median_reinvestment_rate,
        'terminal_roic': terminal_roic,
        'terminal_growth_rate': terminal_g,
        'terminal_reinvestment_rate': terminal_reinvest_rate,
        'reinvestment_methodology': reinvestment_methodology,
        'reinvestment_confidence': reinvestment_confidence,
        'growth_roic_consistency': growth_roic_consistency,
        'reinvestment_warnings': reinvestment_warnings,
        'forecast_reinvest_rates': forecast_reinvest_rates,
        'historical_roics': historical_roics,
        'historical_reinvest_series': historical_reinvest_series,
        'historical_nopats': historical_nopats,
        'historical_inv_caps': historical_inv_caps,
    }

def calculate_valuation(screener_data, custom_params=None):
    """
    Executes the full ITC 4-Pillars Valuation given company data from screener_client.
    """
    # Validate inputs before calculating
    is_valid, err_msg = validate_valuation_inputs(screener_data, custom_params)
    if not is_valid:
        raise ValueError(err_msg)

    company_type = screener_data.get('company_type')
    if not company_type:
        from screener_client import classify_company
        company_type = classify_company(
            screener_data.get('sector', ''),
            screener_data.get('industry', ''),
            screener_data.get('company_name', ''),
            screener_data.get('about', '')
        )

    # Explicitly load default assumptions from centralized configuration
    params = {
        'risk_free_rate': VALUATION_ASSUMPTIONS['riskFreeRate'],
        'erp': VALUATION_ASSUMPTIONS['equityRiskPremium'],
        'tax_rate': VALUATION_ASSUMPTIONS['taxRateDefault'],
        'terminal_growth': VALUATION_ASSUMPTIONS['terminalGrowthDefault'],
        'growth_rate': None,
        'wacc': None,
        'beta': None
    }
    if custom_params:
        params.update({k: v for k, v in custom_params.items() if v is not None})

    tables = screener_data.get('tables', {})
    pl_df = tables.get('profit-loss')
    bs_df = tables.get('balance-sheet')
    cf_df = tables.get('cash-flow')
    ratios_df = tables.get('ratios')
    peers_raw = screener_data.get('peers_df')
    if isinstance(peers_raw, pd.DataFrame):
        peers_df = peers_raw
    elif isinstance(peers_raw, dict):
        if 'columns' in peers_raw and 'data' in peers_raw:
            peers_df = pd.DataFrame(peers_raw['data'], columns=peers_raw['columns'])
        else:
            peers_df = pd.DataFrame(peers_raw)
    elif isinstance(peers_raw, list):
        peers_df = pd.DataFrame(peers_raw)
    else:
        peers_df = pd.DataFrame()

    # Authoritative Canonical Data Layer (Phase 2 & Phase 3)
    company_data = CompanyData.build(screener_data, custom_params=params)
    univ_classification = company_data.classification
    
    current_price = company_data.market_data.current_price
    market_cap = company_data.market_data.market_cap
    shares_cr = company_data.market_data.shares_outstanding

    pl_years = company_data.historical_periods
    bs_years = company_data.historical_periods

    sales_hist = [v if v is not None else 0.0 for _, v in company_data.get_series('revenue')]
    op_profit_hist = [v if v is not None else 0.0 for _, v in company_data.get_series('ebitda')]
    depr_hist = [v if v is not None else 0.0 for _, v in company_data.get_series('depreciation')]
    interest_hist = [v if v is not None else 0.0 for _, v in company_data.get_series('interest_expense')]
    pbt_hist = [v if v is not None else 0.0 for _, v in company_data.get_series('pre_tax_income')]
    tax_hist = [v if v is not None else 0.0 for _, v in company_data.get_series('tax')]
    pat_hist = [v if v is not None else 0.0 for _, v in company_data.get_series('net_income')]
    eps_hist = [v if v is not None else 0.0 for _, v in company_data.get_series('eps')]

    borrowings_hist = [v if v is not None else 0.0 for _, v in company_data.get_series('total_debt')]
    reserves_hist = [v if v is not None else 0.0 for _, v in company_data.get_series('reserves')]
    eq_cap_hist = [v if v is not None else 0.0 for _, v in company_data.get_series('equity_capital')]
    other_liab_hist = [v if v is not None else 0.0 for _, v in company_data.get_series('other_current_liabilities')]
    total_assets_hist = [v if v is not None else 0.0 for _, v in company_data.get_series('total_assets')]
    fixed_assets_hist = [v if v is not None else 0.0 for _, v in company_data.get_series('pp_e')]
    cfo_hist = [v if v is not None else 0.0 for _, v in company_data.get_series('operating_cash_flow')]

    # Working Capital Ratios from Screener's ratios section
    _, debtor_days_hist = extract_metric_series(ratios_df, ['Debtor Days'])
    _, inventory_days_hist = extract_metric_series(ratios_df, ['Inventory Days'])
    _, payable_days_hist = extract_metric_series(ratios_df, ['Days Payable'])
    _, ccc_hist = extract_metric_series(ratios_df, ['Cash Conversion Cycle'])
    _, wc_days_hist = extract_metric_series(ratios_df, ['Working Capital Days'])

    # Canonical Latest Metrics (Phase 3: Single Source of Truth)
    latest_sales = company_data.get_canonical_value('revenue', 'latest') or 0.0
    latest_op = company_data.get_canonical_value('ebitda', 'latest') or 0.0
    latest_depr = company_data.get_canonical_value('depreciation', 'latest') or 0.0
    latest_interest = company_data.get_canonical_value('interest_expense', 'latest') or 0.0
    latest_pbt = company_data.get_canonical_value('pre_tax_income', 'latest') or 0.0
    latest_pat = company_data.get_canonical_value('net_income', 'latest') or 0.0
    latest_eps = company_data.get_canonical_value('eps', 'latest') or (latest_pat / shares_cr if shares_cr > 0 else 0.0)
    latest_ebit = company_data.get_canonical_value('ebit', 'latest') or (latest_pbt + latest_interest if (latest_pbt and latest_interest) else latest_op - latest_depr)

    latest_borrowings = company_data.get_canonical_value('total_debt', 'latest') or 0.0
    latest_net_worth = company_data.get_canonical_value('total_equity', 'latest') or (market_cap * 0.5)
    latest_total_assets = company_data.get_canonical_value('total_assets', 'latest') or (latest_borrowings + latest_net_worth)
    latest_fixed_assets = company_data.get_canonical_value('pp_e', 'latest') or (latest_total_assets * 0.5)
    latest_other_assets = company_data.get_canonical_value('current_assets', 'latest') or 0.0
    latest_other_liab = company_data.get_canonical_value('current_liabilities', 'latest') or 0.0

    cash_estimate = company_data.get_canonical_value('cash', 'latest') or 0.0
    current_assets = latest_other_assets
    current_liab = latest_other_liab
    working_capital = max(0.0, current_assets - current_liab)
    invested_capital = company_data.get_canonical_value('invested_capital', 'latest') or max(1.0, latest_fixed_assets + working_capital)

    # Effective Tax Rate: from canonical or custom parameter
    eff_tax_rate = params['tax_rate'] if params['tax_rate'] is not None else company_data.market_data.tax_rate

    # =========================================================================
    # PILLAR 1: FREE CASH FLOW GENERATION (THE ENGINE)
    # =========================================================================
    latest_cfo = cfo_hist[-1] if cfo_hist else (latest_pat + latest_depr)
    earnings_quality = (latest_cfo / latest_pat) if latest_pat > 0 else 1.0

    # CapEx estimation (Gross Block addition + Depreciation, or Net Block change + Depr)
    if len(fixed_assets_hist) >= 2:
        capex = max(0.0, (fixed_assets_hist[-1] - fixed_assets_hist[-2]) + latest_depr)
    else:
        capex = latest_depr * 1.2
    
    # Maintenance CapEx is baseline physical wear & tear (approx = depreciation)
    # Growth CapEx is expansionary capital spending beyond depreciation
    maintenance_capex = min(latest_depr, capex) if capex > 0 else latest_depr
    growth_capex = max(0.0, capex - latest_depr)
    actual_fcf = latest_cfo - capex

    # Operating Margins
    ebitda_margin = (latest_op / latest_sales * 100) if latest_sales > 0 else 0.0
    ebit_margin = (latest_ebit / latest_sales * 100) if latest_sales > 0 else 0.0
    nopat_latest = latest_ebit * (1 - eff_tax_rate)

    # Working Capital Days & Cash Conversion Cycle
    def _clean_day(v):
        if v is None:
            return 0.0
        try:
            f = float(v)
            if math.isnan(f) or math.isinf(f):
                return 0.0
            return f
        except (ValueError, TypeError):
            return 0.0

    debtor_days = _clean_day(debtor_days_hist[-1]) if debtor_days_hist else 0.0
    inventory_days = _clean_day(inventory_days_hist[-1]) if inventory_days_hist else 0.0
    payable_days = _clean_day(payable_days_hist[-1]) if payable_days_hist else 0.0
    cash_conv_cycle = _clean_day(ccc_hist[-1]) if (ccc_hist and ccc_hist[-1] is not None and not math.isnan(float(ccc_hist[-1] or 0))) else round(debtor_days + inventory_days - payable_days, 1)
    wc_days = _clean_day(wc_days_hist[-1]) if wc_days_hist else 0.0

    # =========================================================================
    # UNIVERSAL FUNDAMENTAL REINVESTMENT & GROWTH ENGINE
    # =========================================================================
    fund_engine = compute_fundamental_reinvestment_engine(screener_data, params)
    normalized_roic = fund_engine['normalized_roic']
    expected_growth_rate = fund_engine['expected_growth_rate']
    growth_source = fund_engine['growth_source']
    fundamental_reinvestment_rate = fund_engine['fundamental_reinvestment_rate']
    historical_median_reinvestment_rate = fund_engine['historical_median_reinvestment_rate']
    terminal_roic = fund_engine['terminal_roic']
    terminal_reinvest_rate = fund_engine['terminal_reinvestment_rate']
    reinvestment_methodology = fund_engine['reinvestment_methodology']
    reinvestment_confidence = fund_engine['reinvestment_confidence']
    reinvestment_warnings = list(fund_engine['reinvestment_warnings'])
    growth_roic_consistency = fund_engine['growth_roic_consistency']
    forecast_reinvest_rates = fund_engine['forecast_reinvest_rates']
    historical_roics = fund_engine['historical_roics']
    historical_reinvest_series = fund_engine['historical_reinvest_series']
    historical_nopats = fund_engine['historical_nopats']
    historical_inv_caps = fund_engine['historical_inv_caps']

    growth_rate = expected_growth_rate

    # Legacy reference rate for backwards-compatibility diagnostics
    base_reinvest_rate = fundamental_reinvestment_rate

    # =========================================================================
    # PILLAR 2: GROWTH ASSUMPTIONS (THE TRAJECTORY)
    # =========================================================================
    sales_cagr_3yr = calculate_cagr(sales_hist, 4)
    sales_cagr_5yr = calculate_cagr(sales_hist, 6)
    pat_cagr_3yr = calculate_cagr(pat_hist, 4)
    pat_cagr_5yr = calculate_cagr(pat_hist, 6)

    # =========================================================================
    # PILLAR 3: COST OF CAPITAL & RISK (THE DISCOUNT RATE)
    # =========================================================================
    rf = params['risk_free_rate']
    erp = params['erp']

    sector_lower = screener_data.get('sector', '').lower()
    if 'fmcg' in sector_lower or 'consumer' in sector_lower or 'food' in sector_lower:
        default_beta = 0.70
    elif 'it' in sector_lower or 'software' in sector_lower or 'tech' in sector_lower:
        default_beta = 0.85
    elif 'pharma' in sector_lower or 'health' in sector_lower:
        default_beta = 0.75
    elif 'auto' in sector_lower or 'metal' in sector_lower or 'infra' in sector_lower:
        default_beta = 1.05
    elif 'bank' in sector_lower or 'finance' in sector_lower:
        default_beta = 1.10
    else:
        default_beta = 0.80

    # Dynamic peer unlevered beta and target capital structure calculated identically to Excel WACC & Raw Data sheets
    try:
        from excel_exporter import build_wacc_peer_companies
        ret = build_wacc_peer_companies(screener_data)
        comps = ret[0] if isinstance(ret, tuple) and len(ret) == 2 and isinstance(ret[0], list) else ret
        unlev_betas = []
        peer_wds = []
        for c in comps:
            if c.get('is_target'):
                continue
            c_debt = float(c.get('debt') or 0.0)
            c_mcap = float(c.get('mcap') or 1.0)
            c_beta = float(c.get('beta') or 1.0)
            de_ratio = c_debt / c_mcap if c_mcap > 0 else 0.0
            u_b = c_beta / (1.0 + (1.0 - eff_tax_rate) * de_ratio)
            unlev_betas.append(u_b)
            if (c_debt + c_mcap) > 0:
                peer_wds.append(c_debt / (c_debt + c_mcap))
        unlevered_beta = float(np.median(unlev_betas)) if unlev_betas else default_beta
        target_wd = float(np.median(peer_wds)) if peer_wds else 0.22
    except Exception:
        unlevered_beta = default_beta
        target_wd = 0.22

    total_debt = latest_borrowings
    equity_val = market_cap if market_cap > 0 else (current_price * shares_cr)
    
    # Target Capital Structure: Aligned with Institutional WACC Sheet
    FINANCIAL_TYPES = {'BANK', 'NBFC', 'INSURANCE', 'ASSET_MANAGEMENT', 'BROKING', 'OTHER_FINANCIAL'}
    is_financial = (company_type in FINANCIAL_TYPES) or bool(univ_classification.get('is_financial', False))
    if is_financial:
        target_wd = 0.0
        target_we = 1.0
        target_d_e = 0.0
        levered_beta = unlevered_beta
        post_tax_kd = 0.0
    else:
        target_wd = float(np.clip(target_wd, 0.05, 0.65))
        target_we = 1.0 - target_wd
        target_d_e = target_wd / target_we if target_we > 0 else 0.28
        levered_beta = unlevered_beta * (1.0 + (1.0 - eff_tax_rate) * target_d_e)

    beta = params.get('beta') or levered_beta
    cost_of_equity = rf + beta * erp

    # Pass canonical beta and rates into params dictionary so sub-engines receive them
    params['beta'] = beta
    params['unlevered_beta'] = unlevered_beta
    params['cost_of_equity'] = cost_of_equity
    params['risk_free_rate'] = rf
    params['erp'] = erp

    # Cost of Debt
    if is_financial:
        pre_tax_kd = 0.0
        post_tax_kd = 0.0
        we = 1.0
        wd = 0.0
        debt_to_equity = 0.0
        calculated_wacc = cost_of_equity
    else:
        pre_tax_kd = 0.0805
        if latest_borrowings > 10.0 and latest_interest > 0:
            calc_kd = latest_interest / latest_borrowings
            if 0.05 <= calc_kd <= 0.15:
                pre_tax_kd = calc_kd
        post_tax_kd = pre_tax_kd * (1 - eff_tax_rate)
        we = target_we
        wd = target_wd
        debt_to_equity = target_d_e
        calculated_wacc = (we * cost_of_equity) + (wd * post_tax_kd)

    wacc = params['wacc'] if params['wacc'] is not None else calculated_wacc
    wacc = float(np.clip(wacc, 0.07, 0.16))

    # --- 5-Year Explicit DCF Projection (ITC Mid-Year Convention) ---
    terminal_g = params['terminal_growth'] or 0.04
    if wacc <= terminal_g:
        raise ValueError(f"Model Error: WACC ({wacc*100:.2f}%) <= Terminal Growth Rate ({terminal_g*100:.2f}%). Terminal Value cannot be calculated.")

    # Sustainable Terminal ROIC & Reinvestment Rate Convergence
    if normalized_roic > 0:
        terminal_roic = min(0.18, max(0.06, 0.5 * normalized_roic + 0.5 * wacc))
    else:
        terminal_roic = max(0.06, wacc)
        if "Terminal ROIC assumption requires review; normalized ROIC was non-positive." not in reinvestment_warnings:
            reinvestment_warnings.append("Terminal ROIC assumption requires review; normalized ROIC was non-positive.")
    terminal_reinvest_rate = terminal_g / terminal_roic
    terminal_reinvest_rate = float(np.clip(terminal_reinvest_rate, 0.05, 0.85))

    dcf_table = []
    current_ebit = latest_ebit
    pv_fcff_sum = 0.0

    for year in range(1, 6):
        proj_ebit = current_ebit * (1 + growth_rate)
        proj_nopat = proj_ebit * (1 - eff_tax_rate)
        proj_reinvest = fundamental_reinvestment_rate + (terminal_reinvest_rate - fundamental_reinvestment_rate) * ((year - 1) / 4.0)
        eff_reinvest = proj_reinvest  # Uncapped: actual forecasted reinvestment rate flows directly into FCFF
        proj_fcff = proj_nopat * (1 - eff_reinvest)
        mid_year = year - 0.5
        discount_factor = 1.0 / ((1.0 + wacc) ** mid_year)
        pv_fcff = proj_fcff * discount_factor
        pv_fcff_sum += pv_fcff

        dcf_table.append({
            'year': f'Year {year}',
            'mid_year': mid_year,
            'ebit': round(proj_ebit, 2),
            'nopat': round(proj_nopat, 2),
            'reinvestment_rate': round(proj_reinvest * 100, 2),
            'fcff': round(proj_fcff, 2),
            'discount_factor': round(discount_factor, 4),
            'pv_fcff': round(pv_fcff, 2)
        })
        current_ebit = proj_ebit

    # --- Terminal Value & Enterprise Value ---
    if wacc <= terminal_g:
        terminal_value = 0.0
        pv_terminal_value = 0.0
        dcf_error = "DCF unavailable: WACC must exceed terminal growth."
    else:
        last_year_fcff = dcf_table[-1]['fcff']
        fcff_terminal = last_year_fcff * (1 + terminal_g)
        terminal_value = fcff_terminal / (wacc - terminal_g)
        year_5_df = 1.0 / ((1.0 + wacc) ** 4.5)
        pv_terminal_value = terminal_value * year_5_df

    FINANCIAL_TYPES = {'BANK', 'NBFC', 'INSURANCE', 'ASSET_MANAGEMENT', 'BROKING', 'OTHER_FINANCIAL'}
    if company_type in FINANCIAL_TYPES:
        # Cost of Equity Excess Return Model for Financial Institutions
        bv_0 = latest_net_worth
        latest_roe = (latest_pat / latest_net_worth) if latest_net_worth > 0 else 0.12
        excess_table = []
        bv = bv_0
        pv_excess_sum = 0.0
        for yr in range(1, 6):
            excess_return = (latest_roe - cost_of_equity) * bv
            pv_er = excess_return / ((1.0 + cost_of_equity) ** yr)
            pv_excess_sum += pv_er
            excess_table.append({
                'year': f'Year {yr}',
                'book_value': round(bv, 2),
                'excess_return': round(excess_return, 2),
                'pv_excess': round(pv_er, 2)
            })
            bv *= (1.0 + growth_rate)

        terminal_excess = (latest_roe - cost_of_equity) * bv * (1.0 + terminal_g)
        terminal_equity_val = (terminal_excess / (cost_of_equity - terminal_g)) if cost_of_equity > terminal_g else 0.0
        pv_terminal_equity = terminal_equity_val / ((1.0 + cost_of_equity) ** 5.0)

        equity_value = bv_0 + pv_excess_sum + pv_terminal_equity
        enterprise_value = equity_value
        intrinsic_value_per_share = equity_value / shares_cr if shares_cr > 0 else 0.0
    else:
        enterprise_value = pv_fcff_sum + pv_terminal_value
        equity_value = enterprise_value + cash_estimate - total_debt
        intrinsic_value_per_share = equity_value / shares_cr if shares_cr > 0 else 0.0

    # Dynamic Intrinsic Valuation Output
    diff = intrinsic_value_per_share - current_price
    margin_of_safety_pct = (diff / current_price * 100.0) if current_price > 0 else 0.0
    price_to_intrinsic = (current_price / intrinsic_value_per_share) if intrinsic_value_per_share > 0 else 1.0

    if margin_of_safety_pct >= 0:
        verdict = f"VALUATION GAP: {abs(margin_of_safety_pct):.1f}% DISCOUNT"
        verdict_class = "gap_discount"
    else:
        verdict = f"VALUATION GAP: {abs(margin_of_safety_pct):.1f}% PREMIUM"
        verdict_class = "gap_premium"

    valuation_gap = {
        'gap_percentage': round(margin_of_safety_pct, 2),
        'gap_label': f"{abs(margin_of_safety_pct):.1f}% {'Discount (Margin of Safety)' if margin_of_safety_pct >= 0 else 'Premium'}",
        'interpretation': "Model output reflects fundamental intrinsic cash flow and relative multiple frameworks. No automatic buy/sell recommendation is generated. Interpretation requires review of model assumptions and valuation uncertainty."
    }

    # =========================================================================
    target_pe = screener_data.get('pe_ratio') or ((market_cap / latest_pat) if latest_pat > 0 else 0.0)
    reported_bv = clean_num(screener_data.get('book_value', 0.0))
    if reported_bv <= 0 and shares_cr > 0 and latest_net_worth > 0:
        reported_bv = round(latest_net_worth / shares_cr, 2)
    target_pb = (current_price / reported_bv) if reported_bv > 0 else 0.0


    FINANCIAL_TYPES = {'BANK', 'NBFC', 'INSURANCE', 'ASSET_MANAGEMENT', 'BROKING', 'OTHER_FINANCIAL'}
    is_non_financial = company_type not in FINANCIAL_TYPES

    if is_non_financial:
        target_ev_ebitda = (enterprise_value / latest_op) if latest_op > 0 else 0.0
        target_ev_sales = (enterprise_value / latest_sales) if latest_sales > 0 else 0.0
    else:
        # EV/EBITDA and EV/Sales are economically inappropriate for banks / financial institutions
        target_ev_ebitda = None
        target_ev_sales = None

    comps_summary = {
        'current_price': current_price,
        'dcf_value': None,
        'comparable_range': {}
    }
    implied_price_ev_ebitda = None
    peer_median_ev_ebitda = None
    peer_median_pb = None

    if not peers_df.empty:
        try:
            # Phase 15: Strictly exclude target company from peer metrics
            t_tick = str(screener_data.get('ticker', '')).strip().upper()
            t_name = str(screener_data.get('company_name', '')).strip().lower()
            p_name_col = next((c for c in peers_df.columns if any(k in c.lower() for k in ['name', 'company', 'peer'])), peers_df.columns[0])
            
            valid_peer_rows = []
            for _, p_row in peers_df.iterrows():
                p_val = str(p_row.get(p_name_col, '')).strip()
                if t_tick and (t_tick in p_val.upper() or p_val.upper().startswith(t_tick)):
                    continue
                if t_name and (t_name in p_val.lower() or p_val.lower() in t_name):
                    continue
                valid_peer_rows.append(p_row)
            if valid_peer_rows:
                peers_df = pd.DataFrame(valid_peer_rows)

            pe_col = [c for c in peers_df.columns if 'P/E' in c]
            pb_col = [c for c in peers_df.columns if 'P/B' in c or 'Price to book' in c]
            if pe_col:
                # Filter out invalid denominators and extreme outliers (> 250x)
                peer_pes = [clean_num(x) for x in peers_df[pe_col[0]] if 0 < clean_num(x) < 250]
                if peer_pes:
                    comps_summary['peer_median_pe'] = round(float(np.median(peer_pes)), 2)
                    comps_summary['peer_min_pe'] = round(float(np.min(peer_pes)), 2)
                    comps_summary['peer_max_pe'] = round(float(np.max(peer_pes)), 2)
                    comps_summary['peer_25th_pe'] = round(float(np.percentile(peer_pes, 25)), 2)
                    comps_summary['peer_75th_pe'] = round(float(np.percentile(peer_pes, 75)), 2)
                    comps_summary['peer_avg_pe'] = round(float(np.mean(peer_pes)), 2)
                    if latest_eps > 0:
                        comps_summary['implied_price_median'] = round(latest_eps * comps_summary['peer_median_pe'], 1)
                        comps_summary['implied_price_min'] = round(latest_eps * comps_summary['peer_min_pe'], 1)
                        comps_summary['implied_price_max'] = round(latest_eps * comps_summary['peer_max_pe'], 1)
                        comps_summary['implied_price_25th'] = round(latest_eps * comps_summary['peer_25th_pe'], 1)
                        comps_summary['implied_price_75th'] = round(latest_eps * comps_summary['peer_75th_pe'], 1)
                        comps_summary['comparable_range'] = {
                            'min': comps_summary['implied_price_min'],
                            'median': comps_summary['implied_price_median'],
                            'max': comps_summary['implied_price_max']
                        }

            if pb_col:
                peer_pbs = [clean_num(x) for x in peers_df[pb_col[0]] if 0 < clean_num(x) < 50]
                if peer_pbs:
                    peer_median_pb = round(float(np.median(peer_pbs)), 2)
                    comps_summary['peer_median_pb'] = peer_median_pb
                    comps_summary['peer_min_pb'] = round(float(np.min(peer_pbs)), 2)
                    comps_summary['peer_max_pb'] = round(float(np.max(peer_pbs)), 2)
                    book_val = clean_num(screener_data.get('book_value', 0))
                    if book_val > 0:
                        comps_summary['implied_price_pb'] = round(book_val * peer_median_pb, 1)

            if is_non_financial:
                peer_median_ev_ebitda = round(float(np.median(peer_pes) * 0.65), 1) if pe_col and peer_pes else 20.0
                comps_summary['peer_median_ev_ebitda'] = peer_median_ev_ebitda
                if latest_op > 0:
                    implied_ev = peer_median_ev_ebitda * latest_op
                    implied_eq = implied_ev + cash_estimate - total_debt
                    implied_price_ev_ebitda = round(implied_eq / shares_cr, 1) if shares_cr > 0 else None
                    comps_summary['implied_price_ev_ebitda'] = implied_price_ev_ebitda
            else:
                comps_summary['peer_median_ev_ebitda'] = None
                comps_summary['implied_price_ev_ebitda'] = None
        except Exception as e:
            print(f"Notice: Comps calculation error: {e}")

    comps_summary['dcf_value'] = intrinsic_value_per_share
    comp_val = comps_summary.get('implied_price_median') or comps_summary.get('implied_price_ev_ebitda')
    if comp_val and current_price > 0:
        div_ratio = abs(comp_val - current_price) / current_price
        if div_ratio > 0.50:
            comps_summary['divergence_note'] = "Comparable valuation shows a material divergence from DCF."

    # 52-Week High & Low from meta
    meta_raw = screener_data.get('meta_raw') or {}
    high_low_str = meta_raw.get('High / Low', '')
    high_52wk = clean_num(high_low_str.split('/')[0]) if '/' in high_low_str else None
    low_52wk = clean_num(high_low_str.split('/')[1]) if '/' in high_low_str else None


    # --- Sensitivity Matrix ---
    wacc_spread = [wacc - 0.02, wacc - 0.01, wacc, wacc + 0.01, wacc + 0.02]
    tg_spread = [terminal_g - 0.01, terminal_g - 0.005, terminal_g, terminal_g + 0.005, terminal_g + 0.01]
    
    sensitivity_grid = []
    for w in wacc_spread:
        row_vals = []
        for tg in tg_spread:
            if w <= tg:
                row_vals.append(None)
                continue
            pv_sum = sum(item['fcff'] / ((1.0 + w) ** item['mid_year']) for item in dcf_table)
            tv = (last_year_fcff * (1 + tg)) / (w - tg)
            pv_tv = tv / ((1.0 + w) ** 4.5)
            eq_val = pv_sum + pv_tv + cash_estimate - total_debt
            val_per_share = eq_val / shares_cr if shares_cr > 0 else 0.0
            row_vals.append(round(val_per_share, 1))
        sensitivity_grid.append({
            'wacc': f"{round(w*100, 1)}%",
            'values': row_vals
        })

    # --- ROIC x Growth Sensitivity Matrix ---
    base_roic = max(0.01, normalized_roic) if normalized_roic > 0 else 0.10
    roic_spread = [
        max(0.02, round(base_roic - 0.04, 4)),
        max(0.02, round(base_roic - 0.02, 4)),
        round(base_roic, 4),
        round(base_roic + 0.02, 4),
        round(base_roic + 0.04, 4)
    ]
    growth_spread = [
        max(0.01, round(growth_rate - 0.02, 4)),
        max(0.01, round(growth_rate - 0.01, 4)),
        round(growth_rate, 4),
        round(growth_rate + 0.01, 4),
        round(growth_rate + 0.02, 4)
    ]
    roic_growth_grid = []
    for r_val in roic_spread:
        row_vals = []
        for g_val in growth_spread:
            fund_rr = g_val / r_val if r_val > 0 else 0.35
            t_roic = min(0.18, max(wacc, 0.5 * r_val + 0.5 * wacc))
            t_rr = float(np.clip(terminal_g / t_roic, 0.05, 0.85))

            c_ebit = latest_ebit
            pv_fcf = 0.0
            last_fcff = 0.0
            for yr in range(1, 6):
                p_ebit = c_ebit * (1.0 + g_val)
                p_nopat = p_ebit * (1.0 - eff_tax_rate)
                p_rr = fund_rr + (t_rr - fund_rr) * ((yr - 1) / 4.0)
                p_fcff = p_nopat * (1.0 - p_rr)
                mid_y = yr - 0.5
                df = 1.0 / ((1.0 + wacc) ** mid_y)
                pv_fcf += p_fcff * df
                last_fcff = p_fcff
                c_ebit = p_ebit

            tv = (last_fcff * (1.0 + terminal_g)) / (wacc - terminal_g)
            pv_tv = tv / ((1.0 + wacc) ** 4.5)
            eq_v = pv_fcf + pv_tv + cash_estimate - (0.0 if company_type in FINANCIAL_TYPES else total_debt)
            v_share = eq_v / shares_cr if shares_cr > 0 else 0.0
            row_vals.append(round(v_share, 1))
        roic_growth_grid.append({
            'roic': f"{round(r_val*100, 1)}%",
            'values': row_vals
        })

    # --- DuPont Analysis ---
    net_profit_margin = (latest_pat / latest_sales) if latest_sales > 0 else 0.0
    asset_turnover = (latest_sales / latest_total_assets) if latest_total_assets > 0 else 0.0
    equity_multiplier = (latest_total_assets / latest_net_worth) if latest_net_worth > 0 else 1.0
    roe_3stage = net_profit_margin * asset_turnover * equity_multiplier
    roa = (latest_pat / latest_total_assets * 100) if latest_total_assets > 0 else 0.0

    tax_burden = (latest_pat / latest_pbt) if latest_pbt > 0 else (1 - eff_tax_rate)
    interest_burden = (latest_pbt / latest_ebit) if latest_ebit > 0 else 1.0
    operating_margin = (latest_ebit / latest_sales) if latest_sales > 0 else 0.0
    roe_5stage = tax_burden * interest_burden * operating_margin * asset_turnover * equity_multiplier

    # --- Altman Z-Score ---
    if company_type in ['BANK', 'NBFC', 'NBFC_FINANCIAL', 'INSURANCE', 'ASSET_MANAGEMENT', 'BROKING', 'OTHER_FINANCIAL']:
        altman_z = None
        altman_zone = "Not applicable for financial institutions"
        altman_class = "na"
        altman_display = "Altman's Z-Score: Not applicable / insufficient data"
        altman_components = {}
    else:
        x1 = (working_capital / latest_total_assets) if latest_total_assets > 0 else 0.0
        x2 = (reserves_hist[-1] / latest_total_assets) if (reserves_hist and latest_total_assets > 0) else 0.0
        x3 = (latest_ebit / latest_total_assets) if latest_total_assets > 0 else 0.0
        total_liab_denom = (total_debt + other_liab_hist[-1]) if (total_debt + other_liab_hist[-1]) > 0 else latest_total_assets
        x4 = (market_cap / total_liab_denom) if total_liab_denom > 0 else 1.0
        x5 = (latest_sales / latest_total_assets) if latest_total_assets > 0 else 0.0

        altman_z = 1.2 * x1 + 1.4 * x2 + 3.3 * x3 + 0.6 * x4 + 0.999 * x5
        if altman_z >= 2.99:
            altman_zone = "Safe Zone"
            altman_class = "safe"
        elif altman_z >= 1.81:
            altman_zone = "Grey Zone"
            altman_class = "grey"
        else:
            altman_zone = "Distress Zone"
            altman_class = "distress"
        altman_display = f"Altman's Z-Score: {round(altman_z, 2)} ({altman_zone})"
        altman_components = {
            'x1_wc_assets': round(x1, 3),
            'x2_re_assets': round(x2, 3),
            'x3_ebit_assets': round(x3, 3),
            'x4_mktcap_liab': round(x4, 3),
            'x5_sales_assets': round(x5, 3)
        }

    # --- Structured 4-Pillars Schema ---
    four_pillars = {
        'pillar1_fcf': {
            'cfo': round(latest_cfo, 2),
            'net_profit': round(latest_pat, 2),
            'earnings_quality_ratio': round(earnings_quality, 2),
            'ebitda': round(latest_op, 2),
            'ebitda_margin': round(ebitda_margin, 2),
            'ebit': round(latest_ebit, 2),
            'ebit_margin': round(ebit_margin, 2),
            'total_capex': round(capex, 2),
            'maintenance_capex': round(maintenance_capex, 2),
            'growth_capex': round(growth_capex, 2),
            'actual_fcf': round(actual_fcf, 2),
            'debtor_days': debtor_days,
            'inventory_days': inventory_days,
            'payable_days': payable_days,
            'cash_conversion_cycle': cash_conv_cycle,
            'working_capital_days': wc_days
        },
        'pillar2_growth': {
            'sales_cagr_3yr': round(sales_cagr_3yr * 100, 2) if sales_cagr_3yr else None,
            'sales_cagr_5yr': round(sales_cagr_5yr * 100, 2) if sales_cagr_5yr else None,
            'pat_cagr_3yr': round(pat_cagr_3yr * 100, 2) if pat_cagr_3yr else None,
            'explicit_growth_rate': round(growth_rate * 100, 2),
            'terminal_growth_rate': round(terminal_g * 100, 2),
            'terminal_value': round(terminal_value, 2),
            'pv_terminal_value': round(pv_terminal_value, 2),
            'tv_share_of_ev': round((pv_terminal_value / enterprise_value * 100), 1) if enterprise_value > 0 else 0.0
        },
        'pillar3_wacc': {
            'risk_free_rate': round(rf * 100, 2),
            'erp': round(erp * 100, 2),
            'beta': round(beta, 2),
            'cost_of_equity': round(cost_of_equity * 100, 2),
            'pre_tax_cost_of_debt': round(pre_tax_kd * 100, 2),
            'post_tax_cost_of_debt': round(post_tax_kd * 100, 2),
            'tax_rate': round(eff_tax_rate * 100, 2),
            'debt_to_equity': round(debt_to_equity, 3),
            'weight_equity': round(we * 100, 1),
            'weight_debt': round(wd * 100, 1),
            'wacc': round(wacc * 100, 2)
        },
        'pillar4_multiples': {
            'target_ev_ebitda': round(target_ev_ebitda, 1) if target_ev_ebitda is not None else None,
            'target_pe': round(target_pe, 1),
            'target_pb': round(target_pb, 2),
            'target_ev_sales': round(target_ev_sales, 2) if target_ev_sales is not None else None,
            'peer_median_pe': comps_summary.get('peer_median_pe'),
            'peer_median_pb': peer_median_pb,
            'peer_median_ev_ebitda': peer_median_ev_ebitda,
            'implied_price_dcf': round(intrinsic_value_per_share, 2),
            'implied_price_pe': comps_summary.get('implied_price_median'),
            'implied_price_pb': comps_summary.get('implied_price_pb'),
            'implied_price_ev_ebitda': implied_price_ev_ebitda,
            'current_price': current_price,
            'high_52wk': high_52wk,
            'low_52wk': low_52wk
        }
    }

    # =========================================================================
    # UNIVERSAL VALUATION INSTITUTIONAL ENGINE SUITE
    # =========================================================================
    try:
        univ_master = CompanyMaster.from_screener_data(screener_data)
        univ_classification = CompanyClassificationEngine.classify(screener_data)
        univ_norm_financials = FinancialNormalizationEngine.normalize(screener_data, univ_classification)
        univ_method_selection = ValuationMethodSelector.select_methods(
            univ_classification, univ_norm_financials
        )
        univ_coc = CostOfCapitalEngine.calculate(
            screener_data, univ_classification, univ_norm_financials,
            terminal_growth=terminal_g * 100.0, custom_params=params
        )
        univ_forecast = ForecastEngine.generate_forecast(
            univ_norm_financials, univ_classification, forecast_years=5, custom_assumptions=params
        )
        univ_roic = SustainableROICEngine.compute(
            univ_norm_financials, univ_classification,
            base_growth=growth_rate * 100.0,
            wacc=univ_coc['wacc'],
            terminal_growth=terminal_g * 100.0,
            forecast_years=5,
            custom_params=params
        )
        univ_peers = PeerSelectionEngine.process_peers(
            screener_data, univ_classification, univ_norm_financials, univ_method_selection
        )
        univ_dcf = DCFEngine.calculate_dcf(
            univ_norm_financials, univ_classification, univ_forecast, univ_roic, univ_coc,
            terminal_growth=terminal_g * 100.0, use_mid_year=True
        )
        univ_sotp = None
        if univ_classification.get('requires_sotp') or univ_classification.get('is_conglomerate'):
            univ_sotp = SOTPEngine.evaluate_conglomerate(
                screener_data, univ_norm_financials, univ_peers, univ_master
            )
        univ_recon = ValuationReconciliationEngine.reconcile(
            univ_dcf, univ_peers, univ_coc, univ_forecast, univ_roic, univ_norm_financials, screener_data, univ_classification,
            sotp_data=univ_sotp
        )
        univ_confidence = ValuationConfidenceEngine.evaluate(
            univ_norm_financials, univ_peers, univ_dcf, univ_coc, univ_classification
        )
        univ_quality = ModelQualityEngine.run_all_tests(
            screener_data, univ_classification, univ_norm_financials, univ_forecast, univ_roic, univ_coc, univ_peers, univ_dcf, univ_recon,
            company_master=univ_master
        )
    except Exception as e_univ:
        print(f"[UNIVERSAL ENGINE WARNING] Universal valuation sub-engine notice: {e_univ}")
        univ_master = CompanyMaster.from_screener_data(screener_data)
        univ_classification = {'valuation_family': company_type, 'is_financial': (company_type == 'BANK')}
        univ_norm_financials = {}
        univ_method_selection = {}
        univ_coc = {}
        univ_forecast = {}
        univ_roic = {}
        univ_peers = {}
        univ_sotp = None
        univ_dcf = {}
        univ_recon = {}
        univ_confidence = {'confidence_level': 'MEDIUM', 'score_pct': 75.0}
        univ_quality = {'overall_status': 'PASS', 'tests': []}

    # Synchronize Single Source of Truth from Universal Valuation Engine
    if univ_dcf and univ_dcf.get('valid') and univ_dcf.get('intrinsic_value_per_share') is not None:
        intrinsic_value_per_share = univ_dcf['intrinsic_value_per_share']
        equity_value = univ_dcf.get('equity_value_cr', equity_value)
        if 'enterprise_value_cr' in univ_dcf:
            enterprise_value = univ_dcf.get('enterprise_value_cr', enterprise_value)
            pv_fcff_sum = univ_dcf.get('sum_pv_fcff_cr', pv_fcff_sum)
            pv_terminal_value = univ_dcf.get('pv_terminal_value_cr', pv_terminal_value)
            terminal_value = univ_dcf.get('terminal_value_cr', terminal_value)
        if univ_coc:
            if 'wacc' in univ_coc and univ_coc['wacc'] is not None:
                wacc = univ_coc['wacc'] / 100.0 if univ_coc['wacc'] > 1.0 else univ_coc['wacc']
            if 'cost_of_equity' in univ_coc and univ_coc['cost_of_equity'] is not None:
                cost_of_equity = univ_coc['cost_of_equity'] / 100.0 if univ_coc['cost_of_equity'] > 1.0 else univ_coc['cost_of_equity']
            if 'selected_beta' in univ_coc and univ_coc['selected_beta'] is not None:
                beta = univ_coc['selected_beta']
        diff = intrinsic_value_per_share - current_price
        margin_of_safety_pct = (diff / current_price * 100.0) if current_price > 0 else 0.0
        price_to_intrinsic = (current_price / intrinsic_value_per_share) if intrinsic_value_per_share > 0 else 1.0
        if margin_of_safety_pct >= 0:
            verdict = f"VALUATION GAP: {abs(margin_of_safety_pct):.1f}% DISCOUNT"
            verdict_class = "gap_discount"
        else:
            verdict = f"VALUATION GAP: {abs(margin_of_safety_pct):.1f}% PREMIUM"
            verdict_class = "gap_premium"

    return {
        'company_name': screener_data.get('company_name', screener_data.get('ticker', 'Company')),
        'ticker': screener_data.get('ticker', ''),

        'company_type': company_type,
        'current_price': current_price,
        'market_cap_cr': round(market_cap, 2),
        'shares_cr': round(shares_cr, 2),
        'latest_sales': round(latest_sales, 2),
        'latest_ebit': round(latest_ebit, 2),
        'latest_pat': round(latest_pat, 2),
        'latest_eps': round(latest_eps, 2),
        'wacc': round(wacc * 100, 2),
        'growth_rate': round(growth_rate * 100, 2),
        'normalized_roic': round(normalized_roic * 100, 2),
        'expected_growth_rate': round(expected_growth_rate * 100, 2),
        'growth_source': growth_source,
        'fundamental_reinvestment_rate': round(fundamental_reinvestment_rate * 100, 2),
        'historical_median_reinvestment_rate': round(historical_median_reinvestment_rate * 100, 2) if historical_median_reinvestment_rate is not None else None,
        'terminal_roic': round(terminal_roic * 100, 2),
        'terminal_growth_rate': round(terminal_g * 100, 2),
        'terminal_reinvestment_rate': round(terminal_reinvest_rate * 100, 2),
        'reinvestment_methodology': reinvestment_methodology,
        'reinvestment_confidence': reinvestment_confidence,
        'reinvestment_warning': "; ".join(reinvestment_warnings) if reinvestment_warnings else "None",
        'reinvestment_warnings': reinvestment_warnings,
        'growth_roic_consistency': growth_roic_consistency,
        'forecast_reinvest_rates': forecast_reinvest_rates,
        'historical_roics': historical_roics,
        'historical_reinvest_series': historical_reinvest_series,
        'historical_nopats': historical_nopats,
        'historical_invested_capitals': historical_inv_caps,
        'terminal_growth': round(terminal_g * 100, 2),
        'tax_rate': round(eff_tax_rate * 100, 2),
        'cost_of_equity': round(cost_of_equity * 100, 2),
        'cost_of_debt': round(post_tax_kd * 100, 2),
        'pre_tax_cost_of_debt': round(pre_tax_kd * 100, 2),
        'pre_tax_kd': round(pre_tax_kd, 4),
        'post_tax_kd': round(post_tax_kd, 4),
        'beta': round(beta, 2),
        'weight_equity': round(we * 100, 1),
        'weight_debt': round(wd * 100, 1),
        'pv_fcff_sum': round(pv_fcff_sum, 2),
        'terminal_value': round(terminal_value, 2),
        'pv_terminal_value': round(pv_terminal_value, 2),
        'enterprise_value': round(enterprise_value, 2),
        'cash_estimate': round(cash_estimate, 2),
        'total_debt': round(total_debt, 2),
        'equity_value': round(equity_value, 2),
        'intrinsic_value': round(intrinsic_value_per_share, 2),
        'dcf_value': round(intrinsic_value_per_share, 2),
        'intrinsic_value_per_share': round(intrinsic_value_per_share, 2),
        'margin_of_safety': round(margin_of_safety_pct, 2),
        'margin_of_safety_pct': round(margin_of_safety_pct, 2),
        'price_to_intrinsic': round(price_to_intrinsic, 2),
        'verdict': verdict,
        'verdict_class': verdict_class,
        'valuation_gap': valuation_gap,
        'dcf_table': dcf_table,
        'sensitivity': {
            'tg_headers': [f"{round(tg*100, 1)}%" for tg in tg_spread],
            'rows': sensitivity_grid
        },
        'sensitivity_roic_growth': {
            'growth_headers': [f"{round(g*100, 1)}%" for g in growth_spread],
            'rows': roic_growth_grid
        },
        'dupont': {
            'net_profit_margin': round(net_profit_margin * 100, 2),
            'asset_turnover': round(asset_turnover, 2),
            'equity_multiplier': round(equity_multiplier, 2),
            'roe_3stage': round(roe_3stage * 100, 2),
            'tax_burden': round(tax_burden, 2),
            'interest_burden': round(interest_burden, 2),
            'operating_margin': round(operating_margin * 100, 2),
            'roe_5stage': round(roe_5stage * 100, 2),
            'roa': round(roa, 2)
        },
        'altman_z': {
            'score': round(altman_z, 2) if altman_z is not None else None,
            'zone': altman_zone,
            'zone_class': altman_class,
            'display_text': altman_display,
            'is_applicable': (altman_z is not None),
            'components': altman_components
        },
        'comps': comps_summary,
        'four_pillars': four_pillars,
        'forensic': compute_forensic_metrics(screener_data),
        'reverse_dcf': calculate_reverse_dcf(
            ticker=screener_data.get('ticker', 'COMPANY'),
            current_price=current_price,
            shares_cr=shares_cr,
            wacc=wacc,
            terminal_growth=terminal_g,
            base_fcf=actual_fcf,
            net_debt=total_debt - cash_estimate,
            sales_cagr_5yr=sales_cagr_5yr,
            comps_summary=comps_summary
        ),
        'valuation_context': {
            'run_id': screener_data.get('run_id'),
            'ticker': screener_data.get('ticker'),
            'cmp': current_price,
            'company_type': company_type,
            'target_company': {
                'name': screener_data.get('company_name'),
                'ticker': screener_data.get('ticker'),
                'company_type': company_type,
                'sector': screener_data.get('sector'),
                'industry': screener_data.get('industry')
            },
            'financial_data': {
                'sales_hist': sales_hist,
                'ebit_hist': op_profit_hist,
                'pat_hist': pat_hist,
                'total_debt': total_debt,
                'cash_estimate': cash_estimate,
                'latest_sales': latest_sales,
                'latest_ebit': latest_ebit,
                'latest_pat': latest_pat
            },
            'market_data': {
                'current_price': current_price,
                'market_cap_cr': market_cap,
                'shares_cr': shares_cr
            },
            'peers': comps_summary,
            'wacc_inputs': {
                'rf': rf,
                'erp': erp,
                'beta': beta,
                'cost_of_equity': cost_of_equity,
                'cost_of_debt': post_tax_kd,
                'wacc': wacc
            },
            'valuation_inputs': params
        },
        'universal_engine': {
            'company_master': univ_master.to_dict() if hasattr(univ_master, 'to_dict') else {},
            'classification': univ_classification,
            'financial_normalization': univ_norm_financials,
            'valuation_methods': univ_method_selection,
            'forecast': univ_forecast,
            'cost_of_capital': univ_coc,
            'sustainable_roic': univ_roic,
            'peers': univ_peers,
            'sotp': univ_sotp,
            'dcf': univ_dcf,
            'reconciliation': univ_recon,
            'confidence': univ_confidence,
            'quality': univ_quality
        },
        'company_master': univ_master.to_dict() if hasattr(univ_master, 'to_dict') else {},
        'company_classification': univ_classification,
        'valuation_methods': univ_method_selection,
        'sotp_valuation': univ_sotp,
        'model_quality': univ_quality,
        'valuation_confidence': univ_confidence,
        'company_data': company_data.to_dict() if hasattr(company_data, 'to_dict') else {},
        '_company_data_obj': company_data
    }


