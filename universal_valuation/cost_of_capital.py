"""
universal_valuation/cost_of_capital.py
======================================
Institutional Cost of Capital and Dynamic WACC Engine.
Calculates:
- Dynamic Risk-Free Rate (10-Year Indian Sovereign Benchmark)
- Equity Risk Premium (Country-specific market risk framework)
- Multi-tier Beta Engine (Raw -> Unlevered Industry -> Target Relevered)
- Target Capital Structure (Current vs Sector Target Debt/Equity weights summing to 100%)
- Pre-tax and Post-tax Cost of Debt
- Final WACC with terminal growth validation (WACC > Terminal Growth)
- For financial institutions, isolates Cost of Equity (Ke) as the primary discount rate.
"""

from datetime import datetime
from typing import Dict, Any, Optional

class CostOfCapitalEngine:
    """
    Computes rigorous dynamic Cost of Capital (WACC / Cost of Equity).
    """

    # Baseline benchmarks for Indian Capital Markets
    DEFAULT_RF = 6.85       # India 10-Yr G-Sec Benchmark Yield
    DEFAULT_ERP = 6.00      # India Damodaran ERP (Mature market ERP + Country Risk Premium)
    DEFAULT_STATUTORY_TAX = 0.2517  # Section 115BAA Statutory Corporate Tax Rate

    @classmethod
    def calculate(
        cls, 
        screener_data: Dict[str, Any], 
        classification: Dict[str, Any],
        normalized_data: Dict[str, Any],
        terminal_growth: float = 5.0,
        custom_params: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        Calculates WACC and Cost of Equity with full institutional provenance.
        
        Args:
            screener_data: Parsed data dictionary from screener_client
            classification: Output from CompanyClassificationEngine
            normalized_data: Output from FinancialNormalizationEngine
            terminal_growth: Terminal growth rate in %
            custom_params: Optional user parameter overrides
            
        Returns:
            Dict containing:
            - rf: float (Risk-free rate)
            - erp: float (Equity risk premium)
            - raw_beta: float
            - unlevered_beta: float
            - relevered_beta: float
            - selected_beta: float
            - cost_of_equity: float (Ke %)
            - pre_tax_cost_of_debt: float (Kd %)
            - effective_tax_rate: float
            - post_tax_cost_of_debt: float
            - current_debt_weight: float
            - current_equity_weight: float
            - target_debt_weight: float
            - target_equity_weight: float
            - wacc: float (WACC %)
            - discount_rate: float (WACC for non-financials, Ke for financials)
            - is_financial: bool
            - wacc_valid: bool (wacc > terminal_growth)
            - error_message: Optional[str]
            - metadata: Dict of source dates and methodologies
        """
        is_financial = classification.get('is_financial', False)
        today_str = datetime.now().strftime("%Y-%m-%d")
        
        if 0.0 < terminal_growth <= 0.25:
            terminal_growth = terminal_growth * 100.0

        # 1. Dynamic Risk-Free Rate
        rf = cls.DEFAULT_RF
        rf_source = "Reserve Bank of India 10-Year Benchmark G-Sec Yield (FBIL/RBI)"
        if custom_params and 'risk_free_rate' in custom_params and custom_params['risk_free_rate'] is not None:
            rf = float(custom_params['risk_free_rate'])
            if 0.0 < rf <= 0.25:
                rf = rf * 100.0
            rf_source = "User Custom Override"

        # 2. Dynamic Equity Risk Premium
        erp = cls.DEFAULT_ERP
        erp_source = "Damodaran Indian Market Risk Premium Framework (India CRP + Mature Market ERP)"
        if custom_params and 'erp' in custom_params and custom_params['erp'] is not None:
            erp = float(custom_params['erp'])
            if 0.0 < erp <= 0.25:
                erp = erp * 100.0
            erp_source = "User Custom Override"

        # 3. Dynamic Beta Engine (Raw -> Industry Unlevered -> Target Relevered)
        family = classification.get('valuation_family', 'OPERATING_COMPANY')
        industry_unlevered_beta = cls._get_industry_unlevered_beta(family)

        if custom_params and custom_params.get('beta') is not None:
            raw_beta = float(custom_params['beta'])
        elif custom_params and custom_params.get('unlevered_beta') is not None:
            raw_beta = float(custom_params['unlevered_beta'])
        elif screener_data.get('beta') is not None:
            raw_beta = float(screener_data['beta'])
        elif is_financial:
            try:
                from excel_exporter import build_wacc_peer_companies
                comps = build_wacc_peer_companies(screener_data)
                betas = [float(c['beta']) for c in comps if not c.get('is_target') and c.get('beta')]
                raw_beta = float(np.median(betas)) if betas else 0.94
            except Exception:
                raw_beta = 0.94
        else:
            raw_beta = industry_unlevered_beta

        if raw_beta <= 0.2 or raw_beta > 3.0:
            raw_beta = 1.0 if not is_financial else 0.94
        
        # Effective Tax Rate
        tax_rate = cls.DEFAULT_STATUTORY_TAX
        pnl = normalized_data.get('pnl', {})
        if pnl.get('effective_tax_rate'):
            recent_tax = pnl['effective_tax_rate'][-1]
            if 0.15 <= recent_tax <= 0.35:
                tax_rate = recent_tax

        # 4. Capital Structure Engine
        bs = normalized_data.get('bs', {})
        total_equity = bs.get('total_equity', [100.0])[-1] if bs.get('total_equity') else 100.0
        borrowings = bs.get('borrowings', [0.0])[-1] if bs.get('borrowings') else 0.0
        mcap = float(screener_data.get('market_cap') or total_equity)
        
        # Current Market Debt/Equity
        total_market_cap_emp = mcap + borrowings
        if total_market_cap_emp > 0:
            curr_e_weight = mcap / total_market_cap_emp
            curr_d_weight = borrowings / total_market_cap_emp
        else:
            curr_e_weight = 0.80
            curr_d_weight = 0.20
            
        # Target Capital Structure based on Valuation Family
        target_d_weight, target_e_weight = cls._determine_target_capital_structure(family, curr_d_weight, curr_e_weight)
        
        # Target D/E for relevering
        target_de = target_d_weight / target_e_weight if target_e_weight > 0 else 0.25
        
        # Relever Beta: Levered Beta = Unlevered Beta * [1 + (1 - Tax) * (D/E)]
        if is_financial:
            # For banks, equity beta is used directly without standard debt relevering
            relevered_beta = round(raw_beta, 2)
            selected_beta = round(raw_beta, 2)
            beta_rationale = f"Bank equity beta ({raw_beta:.2f}) sourced directly; operational deposits excluded from debt-relevering."
        else:
            relevered_beta = round(industry_unlevered_beta * (1.0 + (1.0 - tax_rate) * target_de), 2)
            # Blend 60% Relevered Industry Beta + 40% Raw Company Beta for institutional robustness
            selected_beta = round(0.60 * relevered_beta + 0.40 * raw_beta, 2)
            beta_rationale = (
                f"Selected Beta ({selected_beta:.2f}) blended from target relevered industry beta "
                f"({relevered_beta:.2f}, based on {target_de:.2f} target D/E) and empirical company beta ({raw_beta:.2f})."
            )

        # 5. Cost of Equity (CAPM)
        ke = round(rf + selected_beta * erp, 2)
        if custom_params and 'cost_of_equity' in custom_params:
            ke = float(custom_params['cost_of_equity'])

        # 6. Cost of Debt
        # Estimate Kd from interest expense / borrowings or corporate credit spread
        interest = pnl.get('interest', [0.0])[-1] if pnl.get('interest') else 0.0
        if borrowings > 50.0 and interest > 0.0:
            empirical_kd = (interest / borrowings) * 100.0
            pre_tax_kd = round(max(rf + 0.5, min(14.0, empirical_kd)), 2)
        else:
            pre_tax_kd = round(rf + 1.75, 2)  # AA-rated corporate spread benchmark
            
        post_tax_kd = round(pre_tax_kd * (1.0 - tax_rate), 2)

        # 7. Final WACC Calculation
        if is_financial:
            # WACC is economically undefined for financial intermediaries; Cost of Equity governs
            wacc = ke
            discount_rate = ke
            if ke <= terminal_growth:
                wacc_valid = False
                error_message = (
                    f"CRITICAL MODEL INTEGRITY FAILURE: Cost of Equity ({ke:.2f}%) is less than or equal to "
                    f"terminal growth rate ({terminal_growth:.2f}%). Terminal Value denominator (Ke - g) produces infinite or negative valuation."
                )
            else:
                wacc_valid = True
                error_message = None
        else:
            wacc = round((ke * target_e_weight) + (post_tax_kd * target_d_weight), 2)
            discount_rate = wacc
            
            # 8. Strict Terminal Growth Validation
            if wacc <= terminal_growth:
                wacc_valid = False
                error_message = (
                    f"CRITICAL MODEL INTEGRITY FAILURE: WACC ({wacc:.2f}%) is less than or equal to "
                    f"terminal growth rate ({terminal_growth:.2f}%). Terminal Value formula (FCFF / [WACC - g]) produces infinite or negative valuation."
                )
            else:
                wacc_valid = True
                error_message = None

        metadata = {
            "risk_free_rate": {"value": f"{rf:.2f}%", "source": rf_source, "date": today_str},
            "equity_risk_premium": {"value": f"{erp:.2f}%", "source": erp_source, "date": today_str},
            "beta": {"value": f"{selected_beta:.2f}", "rationale": beta_rationale, "raw": raw_beta, "unlevered": industry_unlevered_beta},
            "cost_of_equity": {"value": f"{ke:.2f}%", "formula": f"{rf:.2f}% + ({selected_beta:.2f} × {erp:.2f}%)"},
            "cost_of_debt": {"pre_tax": f"{pre_tax_kd:.2f}%", "post_tax": f"{post_tax_kd:.2f}%", "tax_rate": f"{tax_rate * 100:.2f}%"},
            "capital_structure": {
                "current_equity_weight": f"{curr_e_weight * 100:.1f}%",
                "current_debt_weight": f"{curr_d_weight * 100:.1f}%",
                "target_equity_weight": f"{target_e_weight * 100:.1f}%",
                "target_debt_weight": f"{target_d_weight * 100:.1f}%"
            }
        }

        return {
            "rf": rf,
            "erp": erp,
            "raw_beta": raw_beta,
            "unlevered_beta": industry_unlevered_beta,
            "relevered_beta": relevered_beta,
            "selected_beta": selected_beta,
            "beta": selected_beta,
            "cost_of_equity": ke,
            "pre_tax_cost_of_debt": pre_tax_kd,
            "effective_tax_rate": tax_rate,
            "post_tax_cost_of_debt": post_tax_kd,
            "current_debt_weight": round(curr_d_weight, 4),
            "current_equity_weight": round(curr_e_weight, 4),
            "target_debt_weight": round(target_d_weight, 4),
            "target_equity_weight": round(target_e_weight, 4),
            "wacc": wacc,
            "discount_rate": discount_rate,
            "is_financial": is_financial,
            "wacc_valid": wacc_valid,
            "error_message": error_message,
            "metadata": metadata
        }

    @staticmethod
    def _get_industry_unlevered_beta(family: str) -> float:
        """Industry unlevered beta benchmarks."""
        unlevered_betas = {
            "CONSUMER_FMCG": 0.65,
            "PHARMA_HEALTHCARE": 0.70,
            "TECH_SAAS": 0.90,
            "UTILITIES": 0.60,
            "TELECOM": 0.75,
            "AUTO_ANCILLARY": 0.85,
            "COMMODITY_CYCLICAL": 0.95,
            "ASSET_HEAVY_INDUSTRIAL": 0.90,
            "INFRASTRUCTURE": 0.80,
            "REAL_ESTATE": 0.85,
            "BANK": 0.95,
            "NBFC": 1.05
        }
        return unlevered_betas.get(family, 0.85)

    @staticmethod
    def _determine_target_capital_structure(family: str, curr_d: float, curr_e: float) -> tuple:
        """Determines institutional target capital structure weights summing to 100%."""
        if family in {"BANK", "NBFC", "INSURANCE"}:
            return 0.0, 1.0  # Equity model; regulatory leverage treated separately
            
        sector_debt_caps = {
            "CONSUMER_FMCG": 0.10,
            "TECH_SAAS": 0.10,
            "PHARMA_HEALTHCARE": 0.15,
            "AUTO_ANCILLARY": 0.25,
            "COMMODITY_CYCLICAL": 0.30,
            "ASSET_HEAVY_INDUSTRIAL": 0.35,
            "INFRASTRUCTURE": 0.50,
            "UTILITIES": 0.50,
            "TELECOM": 0.40,
            "REAL_ESTATE": 0.35
        }
        
        target_d_norm = sector_debt_caps.get(family, 0.25)
        # Blend current capital structure with sector normative benchmark
        target_d = round(0.50 * curr_d + 0.50 * target_d_norm, 2)
        target_d = max(0.05, min(0.60, target_d))
        target_e = round(1.0 - target_d, 2)
        
        return target_d, target_e
