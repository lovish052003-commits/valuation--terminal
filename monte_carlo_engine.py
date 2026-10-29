"""
monte_carlo_engine.py
======================
Institutional Monte Carlo Valuation Simulation Engine.

Performs stochastic multi-parameter DCF and Excess Return simulations across:
1. 5-Year Revenue / EBIT / Net Profit Growth Rate
2. Cost of Capital (WACC / Cost of Equity)
3. Operating Margin & Reinvestment Dynamics
4. Terminal Growth Rate

Generates comprehensive probability distributions, percentiles (P5 to P95),
probability of undervaluation vs current price, Value-at-Risk (VaR 95%),
scenario distributions, and parameter correlation sensitivity (Tornado data).
"""

import math
import numpy as np
from typing import Dict, Any, Optional

def run_monte_carlo_simulation(
    valuation_result: Dict[str, Any],
    screener_data: Optional[Dict[str, Any]] = None,
    config: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """
    Executes a vectorized Monte Carlo simulation (1,000 to 10,000 iterations)
    for equity valuation.

    Parameters:
    - valuation_result: dict produced by valuation_engine.calculate_valuation
    - screener_data: optional dict from screener_client.fetch_company_data
    - config: optional simulation parameters (iterations, growth_vol, wacc_vol, etc.)

    Returns:
    - JSON-serializable dictionary with distribution statistics, percentiles,
      probabilities, histogram bins, tornado sensitivity, and scenarios.
    """
    config = config or {}
    iterations = int(config.get('iterations', 2500))
    iterations = max(500, min(10000, iterations))

    # Base market & capital parameters
    current_price = float(valuation_result.get('current_price') or 0.0)
    if current_price <= 0 and screener_data:
        current_price = float(screener_data.get('current_price') or 0.0)
    
    shares_cr = float(valuation_result.get('shares_cr') or 1.0)
    if shares_cr <= 0:
        shares_cr = 1.0

    company_type = str(valuation_result.get('company_type') or 'OTHER_NON_FINANCIAL').upper()
    is_financial = company_type in {'BANK', 'NBFC', 'INSURANCE', 'ASSET_MANAGEMENT', 'BROKING', 'OTHER_FINANCIAL'}

    # Configurable volatilities (std dev)
    growth_vol = float(config.get('growth_volatility', 0.035))      # 3.5% std dev default
    wacc_vol = float(config.get('wacc_volatility', 0.012))          # 1.2% std dev default
    margin_vol = float(config.get('margin_volatility', 0.05))       # 5.0% margin/reinvest vol
    terminal_g_min = float(config.get('terminal_growth_min', 0.02)) # 2.0% min terminal g
    terminal_g_max = float(config.get('terminal_growth_max', 0.055))# 5.5% max terminal g

    # Seed for deterministic reproducibility if requested, else random
    seed = config.get('seed')
    rng = np.random.RandomState(int(seed) if seed is not None else None)

    # ---------------------------------------------------------
    # BRANCH A: Standard Non-Financial DCF Valuation Simulation
    # ---------------------------------------------------------
    if not is_financial:
        base_wacc = float(valuation_result.get('wacc', 11.5)) / 100.0 if valuation_result.get('wacc', 11.5) > 1.0 else float(valuation_result.get('wacc', 0.115))
        base_growth = float(valuation_result.get('growth_rate', 8.0)) / 100.0 if valuation_result.get('growth_rate', 8.0) > 1.0 else float(valuation_result.get('growth_rate', 0.08))
        base_tg = float(valuation_result.get('terminal_growth', 3.5)) / 100.0 if valuation_result.get('terminal_growth', 3.5) > 1.0 else float(valuation_result.get('terminal_growth', 0.035))
        eff_tax_rate = float(valuation_result.get('tax_rate', 25.17)) / 100.0 if valuation_result.get('tax_rate', 25.17) > 1.0 else float(valuation_result.get('tax_rate', 0.2517))
        
        latest_ebit = float(valuation_result.get('latest_ebit') or 1000.0)
        cash_est = float(valuation_result.get('cash_estimate') or 0.0)
        total_debt = float(valuation_result.get('total_debt') or 0.0)
        minority_int = float(valuation_result.get('minority_interest') or 0.0)

        # Reinvestment rate and ROIC
        four_p = valuation_result.get('four_pillars', {})
        p1 = four_p.get('pillar1_fcf', {})
        p3 = four_p.get('pillar3_wacc', {})
        curr_roic = float(p3.get('roic', 15.0)) / 100.0 if p3.get('roic', 15.0) > 1.0 else float(p3.get('roic', 0.15))
        
        # Base reinvestment rate
        base_reinvest = 0.35
        dcf_tbl = valuation_result.get('dcf_table', [])
        if dcf_tbl and len(dcf_tbl) > 0 and 'reinvestment_rate' in dcf_tbl[0]:
            base_reinvest = float(dcf_tbl[0]['reinvestment_rate']) / 100.0

        # Sample stochastic vectors
        sim_growth = rng.normal(base_growth, growth_vol, iterations)
        sim_growth = np.clip(sim_growth, -0.20, 0.50)

        sim_wacc = rng.normal(base_wacc, wacc_vol, iterations)
        sim_wacc = np.clip(sim_wacc, 0.065, 0.22)

        # Terminal growth clamped to be at least 1.2% below simulated WACC
        raw_tg = rng.triangular(terminal_g_min, base_tg, terminal_g_max, iterations)
        sim_tg = np.minimum(raw_tg, sim_wacc - 0.012)
        sim_tg = np.maximum(sim_tg, 0.015)

        # Margin / reinvestment multiplier
        sim_reinvest_mult = rng.normal(1.0, margin_vol, iterations)
        sim_reinvest_mult = np.clip(sim_reinvest_mult, 0.55, 1.45)

        # Vectorized 5-Year Cash Flow Projection
        pv_fcff_sum = np.zeros(iterations)
        curr_ebit = np.full(iterations, latest_ebit)

        for year in range(1, 6):
            fade = (year - 1) / 4.0
            sim_term_roic = np.maximum(curr_roic, sim_wacc)
            sim_term_reinvest = np.clip(sim_tg / np.maximum(sim_term_roic, 0.01), 0.0, 0.85)
            
            proj_reinvest = np.clip(
                (base_reinvest * sim_reinvest_mult) + 
                (sim_term_reinvest - (base_reinvest * sim_reinvest_mult)) * fade,
                0.0, 0.85
            )

            proj_ebit = curr_ebit * (1.0 + sim_growth)
            proj_nopat = proj_ebit * (1.0 - eff_tax_rate)
            proj_fcff = proj_nopat * (1.0 - proj_reinvest)

            mid_year = year - 0.5
            df = 1.0 / ((1.0 + sim_wacc) ** mid_year)
            pv_fcff_sum += proj_fcff * df
            curr_ebit = proj_ebit

        # Terminal Value
        last_fcff = proj_fcff
        fcff_terminal = last_fcff * (1.0 + sim_tg)
        terminal_value = fcff_terminal / np.maximum(sim_wacc - sim_tg, 0.008)
        df_terminal = 1.0 / ((1.0 + sim_wacc) ** 4.5)
        pv_terminal = terminal_value * df_terminal

        sim_ev = pv_fcff_sum + pv_terminal
        sim_equity = sim_ev + cash_est - total_debt - minority_int
        sim_iv = np.maximum(0.5, sim_equity / shares_cr)

        tornado_drivers = {
            'Revenue Growth Rate': round(float(np.corrcoef(sim_growth, sim_iv)[0, 1]), 3),
            'Cost of Capital (WACC)': round(float(np.corrcoef(sim_wacc, sim_iv)[0, 1]), 3),
            'Terminal Growth Rate': round(float(np.corrcoef(sim_tg, sim_iv)[0, 1]), 3),
            'CapEx Reinvestment': round(float(np.corrcoef(sim_reinvest_mult, sim_iv)[0, 1]), 3)
        }

    # ---------------------------------------------------------
    # BRANCH B: Financial / Bank Excess Return Valuation Simulation
    # ---------------------------------------------------------
    else:
        book_val = float(screener_data.get('book_value', 100.0) if screener_data else 100.0)
        curr_roe = float(screener_data.get('roe', 14.0) if screener_data else 14.0) / 100.0
        base_ke = float(valuation_result.get('cost_of_equity', 12.5)) / 100.0 if valuation_result.get('cost_of_equity', 12.5) > 1.0 else float(valuation_result.get('cost_of_equity', 0.125))
        base_tg = 0.05

        sim_roe = rng.normal(curr_roe, 0.025, iterations)
        sim_roe = np.clip(sim_roe, 0.04, 0.30)

        sim_ke = rng.normal(base_ke, wacc_vol, iterations)
        sim_ke = np.clip(sim_ke, 0.08, 0.20)

        sim_tg = rng.triangular(0.03, base_tg, 0.065, iterations)
        sim_tg = np.minimum(sim_tg, sim_ke - 0.012)

        # Excess Return Valuation: P/B = (ROE - g) / (Ke - g)
        justified_pb = (sim_roe - sim_tg) / np.maximum(sim_ke - sim_tg, 0.01)
        justified_pb = np.clip(justified_pb, 0.4, 5.5)
        sim_iv = np.maximum(1.0, book_val * justified_pb)

        tornado_drivers = {
            'Return on Equity (ROE)': round(float(np.corrcoef(sim_roe, sim_iv)[0, 1]), 3),
            'Cost of Equity (Ke)': round(float(np.corrcoef(sim_ke, sim_iv)[0, 1]), 3),
            'Terminal Growth Rate': round(float(np.corrcoef(sim_tg, sim_iv)[0, 1]), 3)
        }

    # ---------------------------------------------------------
    # STATISTICAL ANALYSIS & PERCENTILES
    # ---------------------------------------------------------
    sim_iv_clean = sim_iv[np.isfinite(sim_iv)]
    if len(sim_iv_clean) == 0:
        sim_iv_clean = np.array([current_price or 100.0])

    mean_iv = float(np.mean(sim_iv_clean))
    median_iv = float(np.median(sim_iv_clean))
    std_iv = float(np.std(sim_iv_clean))
    min_iv = float(np.min(sim_iv_clean))
    max_iv = float(np.max(sim_iv_clean))

    # Skewness
    if std_iv > 0:
        skewness = float(np.mean(((sim_iv_clean - mean_iv) / std_iv) ** 3))
    else:
        skewness = 0.0

    # Percentiles
    p5 = float(np.percentile(sim_iv_clean, 5))
    p10 = float(np.percentile(sim_iv_clean, 10))
    p25 = float(np.percentile(sim_iv_clean, 25))
    p50 = median_iv
    p75 = float(np.percentile(sim_iv_clean, 75))
    p90 = float(np.percentile(sim_iv_clean, 90))
    p95 = float(np.percentile(sim_iv_clean, 95))

    # Probability Metrics
    if current_price > 0:
        p_undervalued = float(np.mean(sim_iv_clean > current_price) * 100.0)
        p_overvalued = float(np.mean(sim_iv_clean <= current_price) * 100.0)
        expected_upside = float(((mean_iv - current_price) / current_price) * 100.0)
        
        # 95% Value at Risk (downside from current price at 5th percentile)
        var_95_rs = max(0.0, current_price - p5)
        var_95_pct = (var_95_rs / current_price) * 100.0 if current_price > 0 else 0.0
        
        # Scenario breakdown
        bull_pct = float(np.mean(sim_iv_clean >= 1.20 * current_price) * 100.0)
        fair_pct = float(np.mean((sim_iv_clean >= 0.85 * current_price) & (sim_iv_clean < 1.20 * current_price)) * 100.0)
        bear_pct = float(np.mean(sim_iv_clean < 0.85 * current_price) * 100.0)
    else:
        p_undervalued = 50.0
        p_overvalued = 50.0
        expected_upside = 0.0
        var_95_rs = 0.0
        var_95_pct = 0.0
        bull_pct = 33.3
        fair_pct = 33.4
        bear_pct = 33.3

    # Risk-Reward Ratio (Upside to P90 vs Downside to P10)
    if current_price > 0:
        p90_upside = max(0.01, p90 - current_price)
        p10_downside = max(0.01, current_price - p10)
        risk_reward_ratio = round(p90_upside / p10_downside, 2)
    else:
        risk_reward_ratio = 1.0

    # ---------------------------------------------------------
    # HISTOGRAM & CDF BINS GENERATION (FOR PLOTLY.JS & CLI)
    # ---------------------------------------------------------
    num_bins = 36
    # Filter extreme tail outliers (outside 1st - 99th percentile) for visually balanced histogram
    tail_low = np.percentile(sim_iv_clean, 0.5)
    tail_high = np.percentile(sim_iv_clean, 99.5)
    hist_sample = sim_iv_clean[(sim_iv_clean >= tail_low) & (sim_iv_clean <= tail_high)]
    if len(hist_sample) < 50:
        hist_sample = sim_iv_clean

    counts, bin_edges = np.histogram(hist_sample, bins=num_bins)
    total_samples = len(hist_sample)
    
    bin_midpoints = []
    frequencies = []
    bar_colors = []
    cdf_points = []
    cum_count = 0

    for i in range(len(counts)):
        mid = float((bin_edges[i] + bin_edges[i + 1]) / 2.0)
        freq = float(counts[i] / total_samples * 100.0)
        cum_count += counts[i]
        cdf = float(cum_count / total_samples * 100.0)

        bin_midpoints.append(round(mid, 1))
        frequencies.append(round(freq, 2))
        cdf_points.append(round(cdf, 1))
        
        # Color: Emerald if IV > CMP (undervalued zone), Rose if IV < CMP
        if current_price > 0 and mid > current_price:
            bar_colors.append('#10b981') # Emerald
        else:
            bar_colors.append('#f43f5e') # Rose / Amber

    # Recommendation badge
    if p_undervalued >= 70.0:
        sim_verdict = "STRONG BUY / HIGH PROBABILITY UNDERVALUED"
        sim_class = "buy"
    elif p_undervalued >= 50.0:
        sim_verdict = "FAVORABLE RISK/REWARD / BUY"
        sim_class = "buy"
    elif p_undervalued >= 30.0:
        sim_verdict = "FAIRLY VALUED / HOLD"
        sim_class = "hold"
    else:
        sim_verdict = "HIGH DOWNSIDE RISK / SELL"
        sim_class = "sell"

    return {
        'iterations': iterations,
        'current_price': round(current_price, 2),
        'base_intrinsic_value': round(float(valuation_result.get('intrinsic_value_per_share') or mean_iv), 2),
        'statistics': {
            'mean': round(mean_iv, 2),
            'median': round(median_iv, 2),
            'std_dev': round(std_iv, 2),
            'min': round(min_iv, 2),
            'max': round(max_iv, 2),
            'skewness': round(skewness, 2)
        },
        'percentiles': {
            'p5': round(p5, 2),
            'p10': round(p10, 2),
            'p25': round(p25, 2),
            'p50': round(p50, 2),
            'p75': round(p75, 2),
            'p90': round(p90, 2),
            'p95': round(p95, 2)
        },
        'probabilities': {
            'prob_undervalued': round(p_undervalued, 1),
            'prob_overvalued': round(p_overvalued, 1),
            'expected_upside_pct': round(expected_upside, 1),
            'var_95_rs': round(var_95_rs, 2),
            'var_95_pct': round(var_95_pct, 1),
            'risk_reward_ratio': risk_reward_ratio,
            'verdict': sim_verdict,
            'verdict_class': sim_class
        },
        'scenarios': {
            'bull_pct': round(bull_pct, 1),
            'fair_pct': round(fair_pct, 1),
            'bear_pct': round(bear_pct, 1)
        },
        'tornado_drivers': tornado_drivers,
        'histogram': {
            'bin_edges': [round(float(b), 1) for b in bin_edges],
            'bin_midpoints': bin_midpoints,
            'counts': [int(c) for c in counts],
            'frequencies': frequencies,
            'bar_colors': bar_colors,
            'cdf': cdf_points
        }
    }

def format_cli_monte_carlo(mc_data: Dict[str, Any]) -> str:
    """
    Formats the Monte Carlo simulation result into a rich terminal report with
    summary KPIs, percentiles table, and ASCII distribution histogram.
    """
    stats = mc_data['statistics']
    pct = mc_data['percentiles']
    probs = mc_data['probabilities']
    cmp = mc_data['current_price']
    hist = mc_data['histogram']
    drivers = mc_data['tornado_drivers']

    lines = []
    lines.append("=" * 75)
    lines.append("  MONTE CARLO STOCHASTIC VALUATION TERMINAL")
    lines.append(f"  {mc_data['iterations']:,} Iteration Vectorized DCF / Excess Return Simulation")
    lines.append("=" * 75)

    lines.append(f"\n  * Current Market Price (CMP) : Rs. {cmp:,.2f}")
    lines.append(f"  * Expected Intrinsic Value (Mean): Rs. {stats['mean']:,.2f} (Std Dev: Rs. {stats['std_dev']:,.2f})")
    lines.append(f"  * Median Intrinsic Value (P50) : Rs. {stats['median']:,.2f}")
    lines.append(f"  * Probability of Undervaluation : {probs['prob_undervalued']}%  [P(Intrinsic Value > CMP)]")
    lines.append(f"  * 95% Value at Risk (VaR)       : Rs. {probs['var_95_rs']:,.2f} ({probs['var_95_pct']}% downside from CMP)")
    lines.append(f"  * Risk / Reward Ratio (P90/P10) : {probs['risk_reward_ratio']}x")
    lines.append(f"  * Institutional Verdict         : {probs['verdict']}")

    lines.append("\n--- PERCENTILE DISTRIBUTION (Rs. / Share) ---")
    p_table = [
        f"  P5  (Extreme Bear) : Rs. {pct['p5']:>8.2f}  |  Upside: {((pct['p5']-cmp)/cmp*100):>+6.1f}%",
        f"  P10 (Bear Case)    : Rs. {pct['p10']:>8.2f}  |  Upside: {((pct['p10']-cmp)/cmp*100):>+6.1f}%",
        f"  P25 (Conservative) : Rs. {pct['p25']:>8.2f}  |  Upside: {((pct['p25']-cmp)/cmp*100):>+6.1f}%",
        f"  P50 (Median Base)  : Rs. {pct['p50']:>8.2f}  |  Upside: {((pct['p50']-cmp)/cmp*100):>+6.1f}%",
        f"  P75 (Optimistic)   : Rs. {pct['p75']:>8.2f}  |  Upside: {((pct['p75']-cmp)/cmp*100):>+6.1f}%",
        f"  P90 (Bull Case)    : Rs. {pct['p90']:>8.2f}  |  Upside: {((pct['p90']-cmp)/cmp*100):>+6.1f}%",
        f"  P95 (Super Bull)   : Rs. {pct['p95']:>8.2f}  |  Upside: {((pct['p95']-cmp)/cmp*100):>+6.1f}%",
    ]
    lines.extend(p_table)

    lines.append("\n--- KEY VALUATION SENSITIVITY DRIVERS (TORNADO CORRELATION) ---")
    for factor, corr in sorted(drivers.items(), key=lambda x: abs(x[1]), reverse=True):
        direction = "(+ Impact)" if corr > 0 else "(- Inverse)"
        bar = "#" * int(abs(corr) * 20)
        lines.append(f"  * {factor:<24} : {corr:>+5.2f} {direction} {bar}")

    # ASCII Distribution Curve
    lines.append("\n--- SIMULATION DISTRIBUTION HISTOGRAM ---")
    max_freq = max(hist['frequencies']) if hist['frequencies'] else 1.0
    for i in range(0, len(hist['bin_midpoints']), 2):
        mid = hist['bin_midpoints'][i]
        freq = hist['frequencies'][i]
        bar_len = int((freq / max_freq) * 35)
        marker = " <-- CMP" if abs(mid - cmp) < (hist['bin_midpoints'][1] - hist['bin_midpoints'][0]) * 1.5 else ""
        lines.append(f"  Rs. {mid:>7.1f} | {'#' * bar_len:<35} {freq:>4.1f}%{marker}")

    lines.append("=" * 75)
    return "\n".join(lines)
