"""
universal_valuation/dcf_engine.py
=================================
Universal Intrinsic DCF and Excess Return Valuation Engine.
Supports two institutional valuation architectures:
1. Standard Operating Companies: FCFF DCF using dynamic WACC, mid-year discounting convention,
   fundamental reinvestment rate trajectories, and explicit terminal value validation.
2. Financial Institutions (Banks/NBFCs): Excess Return Model (Cost of Equity vs Sustainable ROE)
   and FCFE equity discounting, strictly prohibiting Enterprise Value debt adjustments.
"""

from typing import Dict, Any, List, Optional

class DCFEngine:
    """
    Executes intrinsic enterprise and equity discounted cash flow valuations.
    """

    @classmethod
    def calculate_dcf(
        cls, 
        normalized_data: Dict[str, Any], 
        classification: Dict[str, Any],
        forecast_data: Dict[str, Any],
        roic_data: Dict[str, Any],
        cost_of_capital: Dict[str, Any],
        terminal_growth: float = 5.0,
        use_mid_year: bool = True
    ) -> Dict[str, Any]:
        """
        Calculates intrinsic equity value per share.
        
        Args:
            normalized_data: Output from FinancialNormalizationEngine
            classification: Output from CompanyClassificationEngine
            forecast_data: Output from ForecastEngine
            roic_data: Output from SustainableROICEngine
            cost_of_capital: Output from CostOfCapitalEngine
            terminal_growth: Terminal growth rate in %
            use_mid_year: Whether to use mid-year discounting (default: True)
            
        Returns:
            Dict containing complete DCF schedule, enterprise value, equity value,
            per-share intrinsic value, terminal value diagnostics, and sensitivity.
        """
        is_financial = classification.get('is_financial', False)
        
        if is_financial:
            return cls._calculate_excess_return_model(
                normalized_data, classification, roic_data, cost_of_capital, terminal_growth
            )
        else:
            return cls._calculate_fcff_dcf(
                normalized_data, forecast_data, roic_data, cost_of_capital, terminal_growth, use_mid_year
            )

    @classmethod
    def _calculate_fcff_dcf(
        cls, 
        normalized_data: Dict[str, Any], 
        forecast_data: Dict[str, Any],
        roic_data: Dict[str, Any],
        cost_of_capital: Dict[str, Any],
        terminal_growth: float,
        use_mid_year: bool
    ) -> Dict[str, Any]:
        if 0.0 < terminal_growth <= 0.25:
            terminal_growth = terminal_growth * 100.0
        wacc = cost_of_capital['wacc']
        if 0.0 < wacc <= 0.35:
            wacc = wacc * 100.0
        wacc_dec = wacc / 100.0
        g_dec = terminal_growth / 100.0
        
        # Validation Stop
        if wacc <= terminal_growth:
            return {
                "valid": False,
                "error": f"Valuation aborted: WACC ({wacc:.2f}%) <= Terminal Growth ({terminal_growth:.2f}%).",
                "intrinsic_value_per_share": 0.0
            }
            
        base_scenario = forecast_data['scenarios']['Base']
        revenues = base_scenario['revenue']
        ebits = base_scenario['ebit']
        nopats = base_scenario['nopat']
        reinvest_schedule = roic_data['reinvestment_fade_schedule']
        forecast_years = len(revenues)
        
        # Explicit FCFF Projections
        fcffs = []
        discount_factors = []
        pv_fcffs = []
        
        for t in range(1, forecast_years + 1):
            nopat_t = nopats[t - 1]
            reinvest_pct = min(0.85, reinvest_schedule[t - 1] / 100.0)
            reinvest_amt = nopat_t * reinvest_pct
            fcff_t = nopat_t - reinvest_amt
            fcffs.append(round(fcff_t, 2))
            
            # Discounting Convention
            t_period = (t - 0.5) if use_mid_year else float(t)
            df = 1.0 / ((1.0 + wacc_dec) ** t_period)
            discount_factors.append(round(df, 4))
            pv_fcffs.append(round(fcff_t * df, 2))
            
        sum_pv_fcff = sum(pv_fcffs)
        
        # Terminal Value Calculation
        # Terminal NOPAT = NOPAT_N * (1 + g)
        nopat_terminal = nopats[-1] * (1.0 + g_dec)
        terminal_reinvest_pct = roic_data['terminal_reinvestment_rate'] / 100.0
        terminal_fcff = nopat_terminal * (1.0 - terminal_reinvest_pct)
        
        terminal_value = terminal_fcff / (wacc_dec - g_dec)
        
        # Discount Terminal Value to Present
        t_term = forecast_years if not use_mid_year else (forecast_years)  # Terminal value at end of forecast
        df_terminal = 1.0 / ((1.0 + wacc_dec) ** forecast_years)
        pv_terminal_value = terminal_value * df_terminal
        
        # Enterprise Value
        enterprise_value = sum_pv_fcff + pv_terminal_value
        
        # Net Debt & Equity Bridge
        bs = normalized_data.get('bs', {})
        borrowings = bs.get('borrowings', [0.0])[-1] if bs.get('borrowings') else 0.0
        cash = bs.get('estimated_cash', [0.0])[-1] if bs.get('estimated_cash') else 0.0
        # Net Debt is un-clamped: if cash > borrowings, net_debt is negative (Net Cash surplus)
        net_debt = borrowings - cash
        
        raw_equity_value = enterprise_value - net_debt
        shares = max(0.01, float(normalized_data.get('shares_outstanding_cr', 100.0) or 100.0))
        
        distress_warning = None
        if raw_equity_value <= 0:
            distress_warning = f"CAPITAL DISTRESS: Net Debt (Rs. {net_debt:.1f} Cr) exceeds Enterprise Value (Rs. {enterprise_value:.1f} Cr); equity is impaired."
            equity_value = 1.0  # Floor for calculation
        else:
            equity_value = raw_equity_value
            
        intrinsic_per_share = round(equity_value / shares, 2)
        
        # Terminal Value Diagnostics
        tv_pct_of_ev = round((pv_terminal_value / enterprise_value) * 100.0, 2) if enterprise_value > 0 else 0.0
        tv_warning = distress_warning
        if tv_pct_of_ev > 85.0 and not tv_warning:
            tv_warning = f"High terminal-value dependence: Terminal Value represents {tv_pct_of_ev:.1f}% of Enterprise Value."

        schedule = []
        for i in range(forecast_years):
            schedule.append({
                "year": f"Year {i+1}",
                "revenue": revenues[i],
                "ebit": ebits[i],
                "nopat": nopats[i],
                "reinvestment_rate_pct": reinvest_schedule[i],
                "fcff": fcffs[i],
                "discount_factor": discount_factors[i],
                "pv_fcff": pv_fcffs[i]
            })

        # Compute Bear and Bull Scenarios (Phase 20)
        def _calc_scenario_val(scen_name, wacc_adj, g_adj):
            scen_data = forecast_data.get('scenarios', {}).get(scen_name)
            if not scen_data:
                return intrinsic_per_share
            scen_nopats = scen_data['nopat']
            s_wacc = max(0.06, wacc_dec + wacc_adj)
            s_g = min(s_wacc - 0.005, max(0.01, g_dec + g_adj))
            s_pv_fcffs = []
            for t in range(1, forecast_years + 1):
                s_nopat = scen_nopats[t - 1]
                s_reinvest = reinvest_schedule[t - 1] / 100.0
                s_fcff = s_nopat * (1.0 - s_reinvest)
                t_p = (t - 0.5) if use_mid_year else float(t)
                s_pv_fcffs.append(s_fcff / ((1.0 + s_wacc) ** t_p))
            s_term_nopat = scen_nopats[-1] * (1.0 + s_g)
            s_term_fcff = s_term_nopat * (1.0 - terminal_reinvest_pct)
            s_tv = s_term_fcff / (s_wacc - s_g)
            s_pv_tv = s_tv / ((1.0 + s_wacc) ** forecast_years)
            s_ev = sum(s_pv_fcffs) + s_pv_tv
            s_eq = max(1.0, s_ev - net_debt)
            return round(s_eq / shares, 2)

        bear_val = _calc_scenario_val('Bear', 0.005, -0.005)
        bull_val = _calc_scenario_val('Bull', -0.005, 0.005)
        bear_val = min(bear_val, intrinsic_per_share)
        bull_val = max(bull_val, intrinsic_per_share)

        return {
            "valid": True,
            "valuation_model": "FCFF_DCF",
            "discount_convention": "Mid-Year Convention" if use_mid_year else "Year-End Convention",
            "forecast_schedule": schedule,
            "sum_pv_fcff_cr": round(sum_pv_fcff, 2),
            "terminal_nopat_cr": round(nopat_terminal, 2),
            "terminal_reinvestment_rate_pct": roic_data['terminal_reinvestment_rate'],
            "terminal_fcff_cr": round(terminal_fcff, 2),
            "terminal_value_cr": round(terminal_value, 2),
            "pv_terminal_value_cr": round(pv_terminal_value, 2),
            "enterprise_value_cr": round(enterprise_value, 2),
            "borrowings_cr": round(borrowings, 2),
            "cash_cr": round(cash, 2),
            "net_debt_cr": round(net_debt, 2),
            "equity_value_cr": round(equity_value, 2),
            "shares_outstanding_cr": shares,
            "intrinsic_value_per_share": intrinsic_per_share,
            "tv_pct_of_ev": tv_pct_of_ev,
            "tv_warning": tv_warning,
            "wacc": wacc,
            "terminal_growth": terminal_growth,
            "scenarios": {
                "Bear": bear_val,
                "Base": intrinsic_per_share,
                "Bull": bull_val
            }
        }

    @classmethod
    def _calculate_excess_return_model(
        cls, 
        normalized_data: Dict[str, Any], 
        classification: Dict[str, Any],
        roic_data: Dict[str, Any],
        cost_of_capital: Dict[str, Any],
        terminal_growth: float
    ) -> Dict[str, Any]:
        """
        Executes Excess Return Model for Banks and Financial Institutions.
        Equity Value = Current Book Value of Equity + PV(Excess Returns on Equity).
        Excess Return = (Sustainable ROE - Cost of Equity) × Book Value of Equity.
        """
        bs = normalized_data.get('bs', {})
        book_value = float(bs.get('total_equity', [1000.0])[-1] if bs.get('total_equity') else 1000.0)
        shares = max(0.01, float(normalized_data.get('shares_outstanding_cr', 100.0) or 100.0))
        
        # Guard against negative book value / capital insolvency
        if book_value <= 0:
            return {
                "valid": False,
                "valuation_model": "EXCESS_RETURN_MODEL",
                "error": "Valuation aborted: Book Value of Equity is non-positive (Regulatory Capital Deficit / Severe Insolvency).",
                "current_book_value_cr": round(book_value, 2),
                "equity_value_cr": 0.0,
                "shares_outstanding_cr": shares,
                "intrinsic_value_per_share": 0.0,
                "institutional_note": "Financial Institution valuation aborted: Negative regulatory equity capital."
            }
        
        if 0.0 < terminal_growth <= 0.25:
            terminal_growth = terminal_growth * 100.0
        ke = cost_of_capital['cost_of_equity']
        if 0.0 < ke <= 0.35:
            ke = ke * 100.0
        ke_dec = ke / 100.0
        g_dec = terminal_growth / 100.0
        
        # Strict Gordon Growth validation for Cost of Equity
        if ke <= terminal_growth:
            return {
                "valid": False,
                "valuation_model": "EXCESS_RETURN_MODEL",
                "error": f"Valuation aborted: Cost of Equity ({ke:.2f}%) <= Terminal Growth ({terminal_growth:.2f}%).",
                "intrinsic_value_per_share": 0.0
            }
        
        # Sustainable ROE from ROIC engine (mapped for financials)
        sustainable_roe = roic_data.get('sustainable_roic', 14.5)
        roe_dec = sustainable_roe / 100.0
        
        # 5-Year Explicit Excess Return Projection
        forecast_years = 5
        bv_t = book_value
        pvs_excess_returns = []
        schedule = []
        
        # Retention Rate = Expected Growth / Sustainable ROE
        expected_loan_growth = 12.0
        retention_rate = min(0.85, max(0.40, expected_loan_growth / sustainable_roe))
        payout_ratio = 1.0 - retention_rate
        
        for t in range(1, forecast_years + 1):
            excess_return_pct = roe_dec - ke_dec
            excess_return_amt = bv_t * excess_return_pct
            df = 1.0 / ((1.0 + ke_dec) ** t)
            pv_er = excess_return_amt * df
            pvs_excess_returns.append(pv_er)
            
            schedule.append({
                "year": f"Year {t}",
                "book_value_beginning": round(bv_t, 2),
                "expected_roe_pct": round(sustainable_roe, 2),
                "cost_of_equity_pct": round(ke, 2),
                "excess_return_spread_pct": round((sustainable_roe - ke), 2),
                "excess_return_cr": round(excess_return_amt, 2),
                "discount_factor": round(df, 4),
                "pv_excess_return_cr": round(pv_er, 2)
            })
            
            # BV grows by retained earnings: BV_t = BV_(t-1) + NetIncome * Retention
            net_income_t = bv_t * roe_dec
            bv_t = bv_t + net_income_t * retention_rate
            
        sum_pv_er = sum(pvs_excess_returns)
        
        # Terminal Excess Return
        # Competitive fade: terminal ROE converges towards Ke + 1.5%
        terminal_roe = max(ke_dec, roe_dec * 0.90)
        terminal_excess_return = bv_t * (terminal_roe - ke_dec)
        
        if ke_dec > g_dec:
            terminal_value = terminal_excess_return / (ke_dec - g_dec)
        else:
            terminal_value = terminal_excess_return / 0.05
            
        df_terminal = 1.0 / ((1.0 + ke_dec) ** forecast_years)
        pv_terminal_value = terminal_value * df_terminal
        
        total_equity_value = book_value + sum_pv_er + pv_terminal_value
        intrinsic_per_share = round(total_equity_value / shares, 2)

        # Scenarios for Excess Return (Phase 20)
        def _calc_excess_scenario(roe_delta, ke_delta, g_delta):
            s_roe = max(0.04, roe_dec + roe_delta)
            s_ke = max(0.06, ke_dec + ke_delta)
            s_g = min(s_ke - 0.01, max(0.01, g_dec + g_delta))
            s_bv = book_value
            s_pvs = []
            for t in range(1, forecast_years + 1):
                s_er = s_bv * (s_roe - s_ke)
                s_pvs.append(s_er / ((1.0 + s_ke) ** t))
                s_bv = s_bv + (s_bv * s_roe * retention_rate)
            s_term_roe = max(s_ke, s_roe * 0.90)
            s_term_er = s_bv * (s_term_roe - s_ke)
            s_tv = s_term_er / (s_ke - s_g)
            s_pv_tv = s_tv / ((1.0 + s_ke) ** forecast_years)
            s_eq = max(1.0, book_value + sum(s_pvs) + s_pv_tv)
            return round(s_eq / shares, 2)

        bear_val = _calc_excess_scenario(-0.025, 0.005, -0.005)
        bull_val = _calc_excess_scenario(0.020, -0.005, 0.005)
        bear_val = min(bear_val, intrinsic_per_share)
        bull_val = max(bull_val, intrinsic_per_share)

        return {
            "valid": True,
            "valuation_model": "EXCESS_RETURN_MODEL",
            "current_book_value_cr": round(book_value, 2),
            "forecast_schedule": schedule,
            "sum_pv_excess_returns_cr": round(sum_pv_er, 2),
            "terminal_excess_return_cr": round(terminal_excess_return, 2),
            "pv_terminal_excess_return_cr": round(pv_terminal_value, 2),
            "equity_value_cr": round(total_equity_value, 2),
            "shares_outstanding_cr": shares,
            "intrinsic_value_per_share": intrinsic_per_share,
            "sustainable_roe_pct": round(sustainable_roe, 2),
            "cost_of_equity_pct": round(ke, 2),
            "spread_pct": round(sustainable_roe - ke, 2),
            "terminal_growth": terminal_growth,
            "scenarios": {
                "Bear": bear_val,
                "Base": intrinsic_per_share,
                "Bull": bull_val
            },
            "institutional_note": (
                "Financial Institution valuation: Enterprise Value / FCFF excluded. "
                "Equity Value derived via Excess Return on Regulatory Equity Capital."
            )
        }
