"""
universal_valuation/method_selector.py
======================================
Institutional Valuation Method Selector (Phase 8).
Selects primary and secondary valuation architectures based on company classification,
profitability profile, segment diversity, and financial structure.
Provides transparent analytical rationale for selected and excluded methodologies.
"""

import pandas as pd
from typing import Dict, Any, List

class ValuationMethodSelector:
    """
    Determines the appropriate valuation methodologies for any listed company.
    Enforces institutional valuation principles:
    - Banning EV/FCFF for financial institutions (banks, NBFCs, insurance).
    - Enforcing SOTP (Sum of the Parts) for multi-segment conglomerates and holding companies.
    - Enforcing mid-cycle normalization for cyclical/commodity companies.
    - Suppressing P/E and EV/EBITDA when underlying earnings/cash flows are negative.
    """

    @classmethod
    def select_methods(cls, classification: Dict[str, Any], financial_summary: Dict[str, Any]) -> Dict[str, Any]:
        """
        Selects valuation methods based on company classification and financial health.
        
        Args:
            classification: Output from CompanyClassificationEngine
            financial_summary: Dict containing latest/normalized metrics or CompanyMaster dict.
        
        Returns:
            Dict containing:
            - primary_method: str ('FCFF_DCF', 'EXCESS_RETURN', 'SOTP_SUM_OF_THE_PARTS', 'NAV_DDM')
            - secondary_methods: List[str]
            - applicable_multiples: List[str]
            - excluded_methods: Dict[str, str] (method -> reason for exclusion)
            - method_rationale: str
            - requires_cyclical_normalization: bool
            - requires_sotp: bool
        """
        family = classification.get("valuation_family", "OPERATING_COMPANY")
        is_financial = classification.get("is_financial", False)
        is_conglomerate = classification.get("is_conglomerate", False)
        is_holding_co = classification.get("is_holding_co", False)
        
        pnl = financial_summary.get('pnl', {})
        bs = financial_summary.get('bs', {})
        norm_metrics = financial_summary.get('normalized_metrics', {})
        tables = financial_summary.get('tables', {})

        latest_ebitda = (
            financial_summary.get("latest_ebitda") 
            or (pnl.get("ebitda", [0.0])[-1] if pnl.get("ebitda") else None)
            or (1.0 if norm_metrics.get("latest_ebitda_margin", 0) > 0 else 0.0)
            or 0.0
        )
        latest_net_income = (
            financial_summary.get("latest_net_income")
            or (pnl.get("net_income", [0.0])[-1] if pnl.get("net_income") else None)
            or 0.0
        )
        latest_book_value = (
            financial_summary.get("latest_book_value")
            or (bs.get("total_equity", [0.0])[-1] if bs.get("total_equity") else None)
            or 0.0
        )
        latest_revenue = (
            financial_summary.get("latest_revenue")
            or (pnl.get("revenue", [0.0])[-1] if pnl.get("revenue") else None)
            or 0.0
        )

        # Fallback: Extract from raw tables if pnl/bs not pre-normalized
        pl_df = tables.get('profit-loss')
        bs_df = tables.get('balance-sheet')
        if pl_df is not None and hasattr(pl_df, 'iterrows'):
            for _, row in pl_df.iterrows():
                metric_name = str(row.get('Metric', '')).strip().lower()
                # EBITDA / Operating Profit
                if latest_ebitda == 0.0 and any(k in metric_name for k in ['operating profit', 'ebitda']):
                    for col in reversed(list(pl_df.columns)):
                        if col != 'Metric' and pd.notna(row[col]):
                            try:
                                latest_ebitda = float(str(row[col]).replace(',', ''))
                                break
                            except (ValueError, TypeError):
                                pass
                # Net Income / PAT
                if latest_net_income == 0.0 and any(k in metric_name for k in ['net profit', 'pat', 'profit after tax']):
                    for col in reversed(list(pl_df.columns)):
                        if col != 'Metric' and pd.notna(row[col]):
                            try:
                                latest_net_income = float(str(row[col]).replace(',', ''))
                                break
                            except (ValueError, TypeError):
                                pass
                # Revenue / Sales
                if latest_revenue == 0.0 and any(k in metric_name for k in ['sales', 'revenue', 'turnover']):
                    for col in reversed(list(pl_df.columns)):
                        if col != 'Metric' and pd.notna(row[col]):
                            try:
                                latest_revenue = float(str(row[col]).replace(',', ''))
                                break
                            except (ValueError, TypeError):
                                pass

        if latest_book_value == 0.0 and bs_df is not None and hasattr(bs_df, 'iterrows'):
            eq_cap = 0.0
            reserves = 0.0
            for _, row in bs_df.iterrows():
                metric_name = str(row.get('Metric', '')).strip().lower()
                if 'equity capital' in metric_name or 'share capital' in metric_name:
                    for col in reversed(list(bs_df.columns)):
                        if col != 'Metric' and pd.notna(row[col]):
                            try:
                                eq_cap = float(str(row[col]).replace(',', ''))
                                break
                            except (ValueError, TypeError):
                                pass
                if 'reserves' in metric_name:
                    for col in reversed(list(bs_df.columns)):
                        if col != 'Metric' and pd.notna(row[col]):
                            try:
                                reserves = float(str(row[col]).replace(',', ''))
                                break
                            except (ValueError, TypeError):
                                pass
            if eq_cap != 0.0 or reserves != 0.0:
                latest_book_value = eq_cap + reserves
        
        primary_method = ""
        secondary_methods = []
        applicable_multiples = []
        excluded_methods = {}
        rationale_parts = []
        requires_cyclical_normalization = False
        requires_sotp = is_conglomerate or is_holding_co
        
        # 1. Financial Institutions (Banks, NBFCs, Insurance)
        if is_financial:
            excluded_methods["FCFF_DCF"] = (
                "Banned for financial institutions: debt and deposits represent core operational raw material "
                "and operating liabilities, not enterprise financing debt. Enterprise Value and WACC are economically undefined."
            )
            excluded_methods["EV_EBITDA"] = (
                "Banned for financial institutions: EBITDA is meaningless where interest expense is the primary operating cost of funds."
            )
            excluded_methods["EV_REVENUE"] = (
                "Banned for financial institutions: top-line revenue does not reflect financial intermediation spreads or risk-weighted assets."
            )
            
            if family == "BANK":
                primary_method = "EXCESS_RETURN"
                secondary_methods = ["FCFE_DCF", "PB_MULTIPLE", "PE_MULTIPLE"]
                rationale_parts.append(
                    "Banking institutions require equity-side valuation: Excess Return Model (Equity Cost of Capital vs ROE on regulatory equity) "
                    "supported by FCFE, P/B (Price-to-Book based on sustainable ROE vs Ke), and P/E."
                )
            elif family == "NBFC":
                primary_method = "EXCESS_RETURN"
                secondary_methods = ["FCFE_DCF", "PB_MULTIPLE", "PE_MULTIPLE"]
                rationale_parts.append(
                    "NBFCs and credit institutions rely on spread-based lending equity models: Excess Return on Book Equity "
                    "paired with sustainable Price-to-Book and Price-to-Earnings multiples."
                )
            elif family == "INSURANCE":
                primary_method = "EXCESS_RETURN"
                secondary_methods = ["PB_MULTIPLE", "PE_MULTIPLE"]
                rationale_parts.append(
                    "Insurance companies rely on underwriting float and regulatory capital surplus: Excess Return / Embedded-Value style "
                    "equity framework and Price-to-Book based on underwriting solvency."
                )
            elif family == "REIT_INVIT":
                primary_method = "NAV_DDM"
                secondary_methods = ["PB_MULTIPLE", "PE_MULTIPLE"]
                rationale_parts.append(
                    "REITs/InvITs distribute statutory mandatory cash flows: Net Asset Value (NAV) and Dividend Discount / Distributable Cash Flow models."
                )
            elif family == "HOLDING_COMPANY":
                primary_method = "SOTP_NAV_HOLDCO"
                secondary_methods = ["PB_MULTIPLE", "PE_MULTIPLE"]
                rationale_parts.append(
                    "Holding companies require Net Asset Value (NAV) of underlying holdings adjusted for standard holding-company discount (20%-30%)."
                )
            
            # Filter multiples based on financial values
            if latest_book_value > 0:
                applicable_multiples.append("P/B")
            else:
                excluded_methods["P/B"] = "Excluded: Book value of equity is non-positive or unavailable."
                
            if latest_net_income > 0:
                applicable_multiples.append("P/E")
            else:
                excluded_methods["P/E"] = "Excluded: Net earnings are negative or depressed."

        # 2. Conglomerates & Multi-Business Incubators (Phase 16)
        elif family == "CONGLOMERATE" or is_conglomerate:
            primary_method = "SOTP_SUM_OF_THE_PARTS"
            secondary_methods = ["FCFF_DCF", "EV_EBITDA", "EV_REVENUE"]
            requires_sotp = True
            requires_cyclical_normalization = True
            rationale_parts.append(
                "Conglomerates and multi-business incubators operate diverse divisions with divergent margins and capital requirements. "
                "Primary valuation is Sum-of-the-Parts (SOTP) segment appraisal paired with consolidated DCF to capture portfolio value."
            )
            if latest_ebitda > 0:
                applicable_multiples.append("EV/EBITDA")
            else:
                excluded_methods["EV/EBITDA"] = "Excluded: Reported consolidated EBITDA is non-positive."
                
            if latest_revenue > 0:
                applicable_multiples.append("EV/Revenue")
                
            if latest_net_income > 0:
                applicable_multiples.append("P/E")
            else:
                excluded_methods["P/E"] = "Excluded: Net earnings are non-positive."
                
            if latest_book_value > 0:
                applicable_multiples.append("P/B")

        # 3. Cyclical & Commodity Companies
        elif family in {"COMMODITY_CYCLICAL", "ASSET_HEAVY_INDUSTRIAL"}:
            primary_method = "FCFF_DCF"
            requires_cyclical_normalization = True
            secondary_methods = ["EV_EBITDA", "PE_MULTIPLE", "EV_REVENUE"]
            rationale_parts.append(
                f"{family.replace('_', ' ').title()} valuation requires mid-cycle margin and normalized reinvestment modeling "
                "within an FCFF DCF framework to avoid peak/trough extrapolation error. Multiples are calibrated against normalized mid-cycle EBITDA."
            )
            
            if latest_ebitda > 0:
                applicable_multiples.append("EV/EBITDA")
            else:
                excluded_methods["EV/EBITDA"] = "Excluded: Reported EBITDA is non-positive; EV/Revenue and cycle-normalized DCF utilized."
                
            if latest_revenue > 0:
                applicable_multiples.append("EV/Revenue")
                
            if latest_net_income > 0:
                applicable_multiples.append("P/E")
            else:
                excluded_methods["P/E"] = "Excluded: Net income is currently non-positive due to cyclical downturn or extraordinary charges."
                
            if latest_book_value > 0:
                applicable_multiples.append("P/B")

        # 4. High-Growth / Asset-Light Tech & SaaS
        elif family == "TECH_SAAS":
            primary_method = "FCFF_DCF"
            secondary_methods = ["EV_EBITDA", "EV_REVENUE", "PE_MULTIPLE"]
            rationale_parts.append(
                "High-operating-leverage Technology / Software valuation uses FCFF DCF driven by long-term operating leverage "
                "and cash conversion, complemented by EV/Revenue (for top-line scaling) and EV/EBITDA (for mature margins)."
            )
            if latest_revenue > 0:
                applicable_multiples.append("EV/Revenue")
            if latest_ebitda > 0:
                applicable_multiples.append("EV/EBITDA")
            else:
                excluded_methods["EV/EBITDA"] = "Excluded: Current reported EBITDA is non-positive."
            if latest_net_income > 0:
                applicable_multiples.append("P/E")
            else:
                excluded_methods["P/E"] = "Excluded: Current net income is negative; standard P/E multiple is economically invalid for loss-making firms."
            if latest_book_value > 0:
                applicable_multiples.append("P/B")


        # 5. Standard Operating Company (Consumer, Pharma, Utilities, Telecom, Auto, Industrial, etc.)
        else:
            primary_method = "FCFF_DCF"
            secondary_methods = ["EV_EBITDA", "PE_MULTIPLE", "EV_REVENUE"]
            rationale_parts.append(
                f"Standard Operating Company ({family.replace('_', ' ').title()}) framework: Enterprise-level Free Cash Flow to Firm (FCFF) DCF "
                "as primary intrinsic model, supplemented by industry peer EV/EBITDA, P/E, and EV/Revenue comparable multiples."
            )
            if latest_ebitda > 0:
                applicable_multiples.append("EV/EBITDA")
            else:
                excluded_methods["EV/EBITDA"] = "Excluded: Current reported EBITDA is non-positive."
                
            if latest_net_income > 0:
                applicable_multiples.append("P/E")
            else:
                excluded_methods["P/E"] = "Excluded: Current net income is negative; standard P/E multiple is economically invalid for loss-making firms."
                
            if latest_revenue > 0:
                applicable_multiples.append("EV/Revenue")
                
            if latest_book_value > 0:
                applicable_multiples.append("P/B")

        if "P/E" in excluded_methods and "PE" not in excluded_methods:
            excluded_methods["PE"] = excluded_methods["P/E"]
        if "EV/EBITDA" in excluded_methods and "EV_EBITDA" not in excluded_methods:
            excluded_methods["EV_EBITDA"] = excluded_methods["EV/EBITDA"]

        return {
            "primary_method": primary_method,
            "secondary_methods": secondary_methods,
            "applicable_multiples": applicable_multiples,
            "excluded_methods": excluded_methods,
            "inapplicable_methods": list(excluded_methods.keys()),
            "method_rationale": " ".join(rationale_parts),
            "requires_cyclical_normalization": requires_cyclical_normalization,
            "requires_sotp": requires_sotp,
            "valuation_family": family,
            "is_financial": is_financial
        }

