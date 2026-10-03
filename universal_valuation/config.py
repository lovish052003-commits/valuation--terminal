"""
universal_valuation/config.py
=============================
Universal Institutional Valuation Configuration.
Defines company-independent operational parameters, tolerances, baseline macroeconomic
rates, and feature flags.

NO company-specific assumptions, ticker rules, or hardcoded overrides are permitted here.
"""

from typing import Dict, Any

VALUATION_CONFIG: Dict[str, Any] = {
    # Forecast Horizon
    "forecast_years": 5,
    "terminal_method": "perpetuity_growth", # 'perpetuity_growth' or 'exit_multiple'
    "mid_year_convention": True,
    
    # Peer Comparables Configuration
    "peer_minimum": 3,
    "peer_target_count": 8,
    "peer_maximum": 15,
    "outlier_method": "IQR_MEDIAN", # Robust non-parametric outlier rejection
    "pe_outlier_upper": 150.0,
    "ev_ebitda_outlier_upper": 80.0,
    
    # Financial Architecture Flags
    "confidence_enabled": True,
    "scenario_analysis": True,
    "sensitivity_analysis": True,
    "sotp_enabled": True,
    "sector_method_selection": True,
    
    # Consistency Tolerances
    "market_cap_share_tolerance_pct": 5.0, # Discrepancy between Price * Shares vs MCap
    "debt_cash_tolerance_pct": 2.0,
    
    # Standard Macroeconomic Baselines (India Institutional)
    "default_rf": 7.00,             # 10Y Indian G-Sec Benchmark Yield (%)
    "default_erp": 6.00,            # Equity Risk Premium (%)
    "default_corporate_tax": 25.17, # Base Corporate Tax Rate under Sec 115BAA (%)
    "terminal_growth_ceiling": 5.50,# Absolute upper bound for terminal growth (must be < WACC)
    "default_terminal_growth": 4.50,# Standard mid-cycle long-term GDP inflation/growth (%)
    
    # Conglomerate / SOTP Configuration
    "holding_company_discount_pct": 20.0,
    
    # Reporting Basis Discipline
    "enforce_strict_reporting_basis": True,
}
