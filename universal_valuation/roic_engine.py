"""
universal_valuation/roic_engine.py
==================================
Fundamental ROIC and Reinvestment Rate Engine.
Implements the core institutional relationship:
    Reinvestment Rate = Expected Growth / Sustainable ROIC
Replaces historical median reinvestment rate with dynamic fundamental economic reinvestment.
Calculates historical, normalized, sustainable, and terminal ROIC with institutional bounds [6%, 18%]
and generates smooth fade trajectories for explicit DCF forecast periods.
"""

import numpy as np
from typing import Dict, Any, List, Optional

class SustainableROICEngine:
    """
    Computes company-specific sustainable ROIC and fundamental reinvestment trajectories.
    """

    # Institutional sanity bounds for Indian listed entities
    ROIC_FLOOR = 6.0    # Cost of debt / economic capital floor
    ROIC_CEILING = 22.0 # Competitive entry / economic rent dissipation ceiling

    @classmethod
    def compute(
        cls, 
        normalized_data: Dict[str, Any], 
        classification: Dict[str, Any],
        base_growth: float,
        wacc: float,
        terminal_growth: float = 5.0,
        forecast_years: int = 5,
        custom_params: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        Calculates sustainable ROIC, fundamental reinvestment rates, and explicit fade schedules.
        
        Args:
            normalized_data: Output from FinancialNormalizationEngine
            classification: Output from CompanyClassificationEngine
            base_growth: Base revenue/NOPAT growth rate in %
            wacc: Weighted Average Cost of Capital in %
            terminal_growth: Long-term terminal growth rate in % (default: 5.0%)
            forecast_years: Explicit forecast horizon (default: 5)
            custom_params: User override dictionary
            
        Returns:
            Dict containing:
            - historical_roics: List[float]
            - historical_median_roic: float
            - recent_roic: float
            - normalized_roic: float
            - sustainable_roic: float
            - terminal_roic: float
            - year1_reinvestment_rate: float (%)
            - terminal_reinvestment_rate: float (%)
            - historical_median_reinvestment: float (%) [diagnostic only]
            - reinvestment_fade_schedule: List[float] (%)
            - roic_fade_schedule: List[float] (%)
            - sanity_flags: List[str]
        """
        if 0.0 < base_growth <= 0.35:
            base_growth = base_growth * 100.0
        if 0.0 < terminal_growth <= 0.25:
            terminal_growth = terminal_growth * 100.0
        if 0.0 < wacc <= 0.35:
            wacc = wacc * 100.0

        ratios = normalized_data.get('ratios', {})
        roic_series = [r for r in ratios.get('roic', []) if r is not None and not np.isnan(r)]
        pnl = normalized_data.get('pnl', {})
        cf = normalized_data.get('cf', {})
        
        # 1. Historical ROIC Analysis
        recent_roic = roic_series[-1] if roic_series else 12.0
        hist_median_roic = float(np.median(roic_series)) if roic_series else recent_roic
        hist_mean_roic = float(np.mean(roic_series)) if roic_series else recent_roic
        
        # 2. Historical Reinvestment Rate (Diagnostic Metric Only)
        hist_reinvestment_diagnostic = cls._calculate_historical_reinvestment_diagnostic(pnl, cf)
        
        # 3. Industry ROIC Anchor based on Valuation Family
        industry_roic_anchor = cls._get_industry_roic_anchor(classification)
        
        # 4. Sustainable ROIC Determination
        # Blend: 40% Recent ROIC + 40% Historical Median + 20% Industry Anchor
        weighted_roic = (0.40 * recent_roic) + (0.40 * hist_median_roic) + (0.20 * industry_roic_anchor)
        
        # Apply sanity boundaries [6%, 22%]
        sanity_flags = []
        if weighted_roic < cls.ROIC_FLOOR:
            sanity_flags.append(f"Sustainable ROIC bounded to institutional floor of {cls.ROIC_FLOOR}% (unbounded was {weighted_roic:.1f}%).")
            sustainable_roic = cls.ROIC_FLOOR
        elif weighted_roic > cls.ROIC_CEILING:
            sanity_flags.append(f"Sustainable ROIC bounded to competitive ceiling of {cls.ROIC_CEILING}% (unbounded was {weighted_roic:.1f}%).")
            sustainable_roic = cls.ROIC_CEILING
        else:
            sustainable_roic = round(weighted_roic, 2)
            
        # User override if provided
        if custom_params and 'sustainable_roic' in custom_params:
            sustainable_roic = float(custom_params['sustainable_roic'])
            sanity_flags.append(f"Sustainable ROIC overridden by user parameter: {sustainable_roic}%.")

        # 5. Terminal ROIC Determination
        # In the terminal state, competitive forces dissipate excess returns towards WACC.
        # Terminal ROIC = min(sustainable_roic, max(wacc, 10.0))
        terminal_roic = round(min(sustainable_roic, max(wacc, 10.0)), 2)
        if custom_params and 'terminal_roic' in custom_params:
            terminal_roic = float(custom_params['terminal_roic'])

        # 6. Fundamental Reinvestment Rate Calculation
        # Year 1 Reinvestment = Expected Growth / Sustainable ROIC
        if sustainable_roic > 0:
            year1_reinvest = (base_growth / sustainable_roic) * 100.0
        else:
            year1_reinvest = 40.0
        # Bound Year 1 reinvestment to [10%, 85%] universal clamp
        year1_reinvest = max(10.0, min(85.0, year1_reinvest))
        
        # Terminal Reinvestment Rate = Terminal Growth / Terminal ROIC
        if terminal_roic > 0:
            terminal_reinvest = (terminal_growth / terminal_roic) * 100.0
        else:
            terminal_reinvest = (terminal_growth / wacc) * 100.0
        terminal_reinvest = max(15.0, min(80.0, terminal_reinvest))
        
        # 7. Explicit Fade Trajectory (Year 1 to Year N)
        reinvestment_fade = []
        roic_fade = []
        for t in range(1, forecast_years + 1):
            weight_t = (t - 1) / float(forecast_years - 1) if forecast_years > 1 else 0.0
            
            # Smooth linear convergence with universal 85% ceiling
            curr_reinvest = min(85.0, (1.0 - weight_t) * year1_reinvest + weight_t * terminal_reinvest)
            curr_roic = (1.0 - weight_t) * sustainable_roic + weight_t * terminal_roic
            
            reinvestment_fade.append(round(curr_reinvest, 2))
            roic_fade.append(round(curr_roic, 2))
            
        return {
            "historical_roics": [round(r, 2) for r in roic_series],
            "historical_median_roic": round(hist_median_roic, 2),
            "recent_roic": round(recent_roic, 2),
            "normalized_roic": round(hist_mean_roic, 2),
            "industry_roic_anchor": industry_roic_anchor,
            "sustainable_roic": sustainable_roic,
            "terminal_roic": terminal_roic,
            "year1_reinvestment_rate": round(year1_reinvest, 2),
            "terminal_reinvestment_rate": round(terminal_reinvest, 2),
            "historical_median_reinvestment_diagnostic": round(hist_reinvestment_diagnostic, 2),
            "reinvestment_fade_schedule": reinvestment_fade,
            "roic_fade_schedule": roic_fade,
            "sanity_flags": sanity_flags,
            "formula_documentation": {
                "fundamental_formula": "Reinvestment Rate = Expected Growth / Sustainable ROIC",
                "terminal_formula": "Terminal Reinvestment Rate = Terminal Growth / Sustainable Terminal ROIC",
                "fade_schedule": f"Linear convergence from Year 1 ({year1_reinvest:.1f}%) to Year {forecast_years} ({reinvestment_fade[-1]:.1f}%)"
            }
        }

    @staticmethod
    def _calculate_historical_reinvestment_diagnostic(pnl: Dict[str, List[float]], cf: Dict[str, List[float]]) -> float:
        """Calculates historical median reinvestment rate strictly for diagnostic comparison."""
        nopats = pnl.get('nopat', [])
        capexs = cf.get('capex', [])
        
        rates = []
        for n, c in zip(nopats, capexs):
            if n > 1.0 and c > 0:
                rate = (c / n) * 100.0
                if 0.0 <= rate <= 150.0:
                    rates.append(rate)
                    
        return float(np.median(rates)) if rates else 45.0

    @staticmethod
    def _get_industry_roic_anchor(classification: Dict[str, Any]) -> float:
        """Provides empirical institutional industry ROIC baselines."""
        family = classification.get('valuation_family', 'OPERATING_COMPANY')
        anchors = {
            "CONSUMER_FMCG": 18.0,
            "PHARMA_HEALTHCARE": 16.0,
            "TECH_SAAS": 18.0,
            "AUTO_ANCILLARY": 13.5,
            "COMMODITY_CYCLICAL": 11.5,
            "ASSET_HEAVY_INDUSTRIAL": 11.0,
            "INFRASTRUCTURE": 10.5,
            "UTILITIES": 9.5,
            "TELECOM": 10.0,
            "REAL_ESTATE": 10.0,
            "BANK": 14.0,  # ROE equivalent
            "NBFC": 14.5   # ROE equivalent
        }
        return anchors.get(family, 12.5)
