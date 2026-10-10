import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client

def calc_workbook_style_valuation(screener_data, custom_params=None):
    from valuation_engine import clean_num
    import numpy as np

    params = {
        'wacc': None,
        'growth_rate': None,
        'terminal_growth': None,
        'tax_rate': None,
        'risk_free_rate': 0.068,
        'erp': 0.065,
        'beta': None
    }
    if custom_params:
        params.update(custom_params)

    # Check financial / banking sector
    sec = screener_data.get('sector', '').lower()
    ind = screener_data.get('industry', '').lower()
    c_type = screener_data.get('company_type', '').upper()
    is_fin = ('bank' in sec or 'finance' in sec or 'insurance' in sec or 'bank' in ind or 'financial' in c_type or 'nbfc' in ind)

    current_price = clean_num(screener_data.get('current_price', 0.0))
    shares_cr = clean_num(screener_data.get('shares_in_cr') or screener_data.get('shares_outstanding_cr') or 0.0)
    market_cap = clean_num(screener_data.get('market_cap_cr', 0.0))
    if shares_cr <= 0 and current_price > 0 and market_cap > 0:
        shares_cr = market_cap / current_price

    norm = screener_data.get('normalized_financials', {}).get('historical', {})
    sales_hist = norm.get('sales', [])
    op_hist = norm.get('operating_profit', [])
    depr_hist = norm.get('depreciation', [])
    interest_hist = norm.get('interest', [])
    pbt_hist = norm.get('pbt', [])
    pat_hist = norm.get('pat', [])
    cfo_hist = norm.get('cfo', [])
    fixed_assets_hist = norm.get('fixed_assets', [])
    borrowings_hist = norm.get('borrowings', [])
    other_assets_hist = norm.get('other_assets', [])
    other_liab_hist = norm.get('other_liabilities', [])
    investments_hist = norm.get('investments', [])
    total_assets_hist = norm.get('total_assets', [])

    latest_sales = clean_num(sales_hist[-1]) if sales_hist else clean_num(screener_data.get('sales_latest', 0.0))
    latest_op = clean_num(op_hist[-1]) if op_hist else latest_sales * 0.15
    latest_depr = clean_num(depr_hist[-1]) if depr_hist else latest_sales * 0.03
    latest_interest = clean_num(interest_hist[-1]) if interest_hist else 0.0
    latest_pat = clean_num(pat_hist[-1]) if pat_hist else latest_sales * 0.10
    latest_borrowings = clean_num(borrowings_hist[-1]) if borrowings_hist else clean_num(screener_data.get('borrowings_cr', 0.0))
    latest_fixed_assets = clean_num(fixed_assets_hist[-1]) if fixed_assets_hist else 1000.0
    latest_other_assets = clean_num(other_assets_hist[-1]) if other_assets_hist else 500.0
    latest_other_liab = clean_num(other_liab_hist[-1]) if other_liab_hist else 300.0
    latest_investments = clean_num(investments_hist[-1]) if investments_hist else clean_num(screener_data.get('investments_cr', 0.0))

    # EBIT: PBT + Interest
    latest_pbt = clean_num(pbt_hist[-1]) if pbt_hist else (latest_op - latest_depr)
    latest_ebit = latest_pbt + latest_interest if latest_interest > 0 else max(latest_op - latest_depr, 1.0)
    if latest_ebit <= 0:
        latest_ebit = max(latest_op - latest_depr, 1.0)

    # Tax Rate
    eff_tax_rate = params['tax_rate'] if params['tax_rate'] is not None else 0.2517
    nopat_latest = latest_ebit * (1 - eff_tax_rate)

    # Invested Capital & Working Capital
    curr_nwc = max(0.0, latest_other_assets - latest_other_liab)
    invested_capital = max(1.0, latest_fixed_assets + curr_nwc)
    curr_roic = (nopat_latest / invested_capital) if invested_capital > 0 else 0.177

    # CapEx & Reinvestment
    latest_cfo = clean_num(cfo_hist[-1]) if cfo_hist else (latest_pat + latest_depr)
    if len(fixed_assets_hist) >= 2:
        capex = max(0.0, (fixed_assets_hist[-1] - fixed_assets_hist[-2]) + latest_depr)
    else:
        capex = latest_depr * 1.2

    net_capex = max(0.0, capex - latest_depr)
    if len(other_assets_hist) >= 2 and len(other_liab_hist) >= 2:
        prev_nwc = max(0.0, other_assets_hist[-2] - other_liab_hist[-2])
        delta_nwc = curr_nwc - prev_nwc
    else:
        delta_nwc = 0.0

    latest_reinvest = net_capex + delta_nwc
    if nopat_latest > 0 and 0.0 <= (latest_reinvest / nopat_latest) <= 0.85:
        base_reinvest_rate = latest_reinvest / nopat_latest
    else:
        base_reinvest_rate = 0.25

    # WACC
    rf = params['risk_free_rate']
    erp = params['erp']
    from excel_exporter import build_wacc_peer_companies, get_realistic_peer_beta, get_sector_key
    sec_k = get_sector_key(screener_data)
    comps = build_wacc_peer_companies(screener_data)
    unlev_betas = []
    for c in comps:
        if c.get('is_target'):
            continue
        c_debt = float(c.get('debt') or 0.0)
        c_mcap = float(c.get('mcap') or 1.0)
        c_beta = float(get_realistic_peer_beta(c, sec_k) or c.get('beta') or 1.0)
        de = c_debt / c_mcap if c_mcap > 0 else 0.0
        u_b = c_beta / (1.0 + (1.0 - eff_tax_rate) * de)
        unlev_betas.append(u_b)
    unlev_b = float(np.median(unlev_betas)) if unlev_betas else 0.80

    total_debt = latest_borrowings
    eq_market = market_cap if market_cap > 0 else (current_price * shares_cr)
    t_de = (total_debt / eq_market) if eq_market > 0 else 0.0
    lev_b = unlev_b * (1.0 + (1.0 - eff_tax_rate) * t_de)
    beta = params['beta'] or lev_b
    ke = rf + beta * erp

    pre_kd = 0.078
    if total_debt > 10.0 and latest_interest > 0:
        calc_kd = latest_interest / total_debt
        if 0.03 <= calc_kd <= 0.15:
            pre_kd = calc_kd
    post_kd = pre_kd * (1 - eff_tax_rate)

    tot_cap = total_debt + eq_market
    we = eq_market / tot_cap if tot_cap > 0 else 0.99
    wd = total_debt / tot_cap if tot_cap > 0 else 0.01
    wacc = params['wacc'] if params['wacc'] is not None else (we * ke + wd * post_kd)
    wacc = float(np.clip(wacc, 0.07, 0.18))

    # Terminal growth & terminal ROIC
    terminal_g = params['terminal_growth'] if params['terminal_growth'] is not None else 0.04
    terminal_roic = max(curr_roic, wacc)
    terminal_reinvest = min(max(terminal_g / terminal_roic, 0.0), 0.85)

    # Growth rate
    if params['growth_rate'] is not None:
        explicit_g = params['growth_rate']
    else:
        explicit_g = base_reinvest_rate * curr_roic

    # 5 Year DCF
    curr_ebit = latest_ebit
    pv_fcff_sum = 0.0
    dcf_table = []
    for y in range(1, 6):
        fade = (y - 1) / 4.0
        rr_y = base_reinvest_rate + (terminal_reinvest - base_reinvest_rate) * fade
        rr_y = min(max(rr_y, 0.0), 0.85)
        step_g = explicit_g if (params['growth_rate'] is not None) else (rr_y * curr_roic)
        ebit_y = curr_ebit * (1 + step_g)
        nopat_y = ebit_y * (1 - eff_tax_rate)
        fcff_y = nopat_y * (1 - rr_y)
        mid_y = y - 0.5
        df_y = (1.0 + wacc) ** (-mid_y)
        pv_fcff_y = fcff_y * df_y
        pv_fcff_sum += pv_fcff_y
        curr_ebit = ebit_y
        dcf_table.append({
            'year': f'Year {y}',
            'mid_year': mid_y,
            'ebit': round(ebit_y, 2),
            'nopat': round(nopat_y, 2),
            'reinvestment_rate': round(rr_y * 100, 2),
            'fcff': round(fcff_y, 2),
            'discount_factor': round(df_y, 4),
            'pv_fcff': round(pv_fcff_y, 2)
        })

    last_fcff = dcf_table[-1]['fcff']
    tv = (last_fcff * (1 + terminal_g)) / (wacc - terminal_g)
    pv_tv = tv * ((1.0 + wacc) ** (-4.5))
    ev = pv_fcff_sum + pv_tv

    # Cash & Liquid Inv
    cash_est = clean_num(screener_data.get('cash_and_equivalents_cr', 0.0))
    if cash_est <= 0:
        cash_est = clean_num(screener_data.get('cash_cr', 0.0))
    cash_est += latest_investments * 0.90

    # Minority Interest
    minority_val = clean_num(screener_data.get('minority_interest_cr', 0.0))
    if minority_val <= 0:
        bs_tbl = screener_data.get('tables', {}).get('balance-sheet')
        if bs_tbl is not None and hasattr(bs_tbl, 'empty') and not bs_tbl.empty:
            m_mi = bs_tbl[bs_tbl['Metric'].str.contains('Minority|Non controlling', case=False, na=False)]
            if not m_mi.empty:
                minority_val = clean_num(m_mi.iloc[0].iloc[-1])

    eq_val = ev + cash_est - total_debt - minority_val
    iv = eq_val / shares_cr if shares_cr > 0 else 0.0

    diff = iv - current_price
    upside_pct = (diff / current_price * 100.0) if current_price > 0 else 0.0
    mos_pct = (diff / iv * 100.0) if iv > 0 else 0.0

    verdict = "UNDERVALUED / BUY" if upside_pct > 15 else ("OVERVALUED / SELL" if upside_pct < -15 else "FAIRLY VALUED / HOLD")
    v_class = 'buy' if 'BUY' in verdict else ('sell' if 'SELL' in verdict else 'hold')

    return {
        'current_price': round(current_price, 2),
        'intrinsic_value': round(iv, 2),
        'intrinsic_value_per_share': round(iv, 2),
        'enterprise_value': round(ev, 2),
        'equity_value': round(eq_val, 2),
        'shares_cr': round(shares_cr, 2),
        'wacc': round(wacc * 100, 2),
        'growth_rate': round(explicit_g * 100, 2),
        'terminal_growth': round(terminal_g * 100, 2),
        'tax_rate': round(eff_tax_rate * 100, 2),
        'upside_pct': round(upside_pct, 2),
        'margin_of_safety_pct': round(mos_pct, 2),
        'verdict': verdict,
        'verdict_class': v_class,
        'dcf_table': dcf_table
    }

if __name__ == '__main__':
    sd = screener_client.fetch_company_data('ITC')
    res = calc_workbook_style_valuation(sd)
    print("ITC Results:")
    for k, v in res.items():
        if k != 'dcf_table':
            print(f"  {k}: {v}")

