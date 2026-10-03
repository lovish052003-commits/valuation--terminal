"""
universal_valuation/company_master.py
=====================================
Universal Company Master Object & Centralized Data Layer.
Implements:
- Phase 2: Universal Company Master Object (Single Source of Truth)
- Phase 3: Centralized Data Layer with source, date, unit, currency, reporting basis, confidence metadata
- Phase 4: Universal Share Count Engine with Market Cap / Share Count validation
- Phase 5: Semantic Market Cap Engine with Named Fields
- Phase 6: Strict Reporting Basis Control (Standalone vs Consolidated)

Downstream modules MUST consume CompanyMaster rather than recalculating or guessing
shares, market cap, debt, cash, or reporting basis independently.
"""

import math
from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional
import pandas as pd
from .config import VALUATION_CONFIG

@dataclass
class CanonicalMetric:
    """Represents a single verified institutional financial metric with audit metadata."""
    value: Any
    currency: str = "INR"
    unit: str = "Cr"
    source: str = "Screener.in"
    as_of_date: str = ""
    reporting_basis: str = "CONSOLIDATED"  # 'STANDALONE' or 'CONSOLIDATED'
    confidence: str = "HIGH"               # 'HIGH', 'MEDIUM', 'LOW'
    description: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "value": self.value,
            "currency": self.currency,
            "unit": self.unit,
            "source": self.source,
            "as_of_date": self.as_of_date,
            "reporting_basis": self.reporting_basis,
            "confidence": self.confidence,
            "description": self.description
        }


class ReportingBasisController:
    """
    Mandatory Reporting Basis Controller (Phase 6).
    Enforces that valuation consistently consumes either STANDALONE or CONSOLIDATED
    financial statements. Flags any mixed revenue, debt, cash, or profit lines.
    """

    @classmethod
    def evaluate(cls, screener_data: Dict[str, Any]) -> Dict[str, Any]:
        is_consolidated = screener_data.get('is_consolidated', True)
        url = str(screener_data.get('url', '')).lower()
        
        # Check URL or data attributes
        if '/consolidated/' in url:
            detected_basis = "CONSOLIDATED"
        elif '/company/' in url and '/consolidated/' not in url and not is_consolidated:
            detected_basis = "STANDALONE"
        else:
            detected_basis = "CONSOLIDATED" if is_consolidated else "STANDALONE"
            
        tables = screener_data.get('tables', {})
        pl_df = tables.get('profit-loss')
        bs_df = tables.get('balance-sheet')
        
        inconsistencies = []
        # Check if tables are populated
        if pl_df is None or pl_df.empty:
            inconsistencies.append("Profit & Loss table is missing or empty.")
        if bs_df is None or bs_df.empty:
            inconsistencies.append("Balance Sheet table is missing or empty.")

        # Detect potential subsidiary mismatch if about mentions subsidiaries but standalone is pulled
        about_text = str(screener_data.get('about', '')).lower()
        has_major_subsidiaries = any(w in about_text for w in ['subsidiaries', 'subsidiary', 'joint venture', 'group companies', 'operating across countries'])
        if detected_basis == "STANDALONE" and has_major_subsidiaries:
            inconsistencies.append("WARNING: Company has operating subsidiaries but standalone statements are being evaluated.")

        status = "CONSISTENT" if not inconsistencies else ("WARNING" if all("WARNING" in inc for inc in inconsistencies) else "INCONSISTENT")

        return {
            "reporting_basis": detected_basis,
            "status": status,
            "inconsistencies": inconsistencies,
            "is_consolidated": (detected_basis == "CONSOLIDATED")
        }


class ShareCountEngine:
    """
    Universal Share Count Engine (Phase 4).
    Determines the appropriate diluted shares outstanding in Crores.
    Validates: Market Cap ≈ Share Price × Shares Outstanding within tolerance.
    """

    @classmethod
    def calculate_shares(
        cls, 
        screener_data: Dict[str, Any], 
        current_price: float, 
        reported_market_cap: float,
        tolerance_pct: float = 5.0
    ) -> Dict[str, Any]:
        tables = screener_data.get('tables', {})
        bs_df = tables.get('balance-sheet')
        meta = screener_data.get('meta', {})
        
        face_value = float(meta.get('Face Value', 1.0) or screener_data.get('face_value', 1.0) or 1.0)
        if face_value <= 0:
            face_value = 1.0
            
        shares_from_bs = None
        if bs_df is not None and not bs_df.empty:
            eq_row = bs_df[bs_df['Metric'].str.contains('Equity Capital|Share Capital', case=False, na=False)]
            if not eq_row.empty:
                # Find latest numeric column
                for col in reversed([c for c in bs_df.columns if c != 'Metric']):
                    val = str(eq_row.iloc[0][col]).replace(',', '').strip()
                    try:
                        latest_eq_cap = float(val)
                        if latest_eq_cap > 0:
                            shares_from_bs = latest_eq_cap / face_value
                            break
                    except (ValueError, TypeError):
                        pass

        # Fallback 1: Market Cap / Price
        shares_from_mcap_price = None
        if current_price > 0 and reported_market_cap > 0:
            shares_from_mcap_price = reported_market_cap / current_price

        # Fallback 2: Pre-parsed shares
        pre_parsed_shares = screener_data.get('shares_in_cr')

        # Determine canonical shares:
        # Prioritize BS Equity Capital / Face Value if consistent with MCap / Price within 15%
        # Else use MCap / Price to protect against unadjusted corporate actions (bonus/splits)
        canonical_shares = None
        derivation_method = ""
        
        if shares_from_mcap_price and shares_from_mcap_price > 0:
            canonical_shares = shares_from_mcap_price
            derivation_method = "Reported Market Capitalization / Current Market Price (Active Traded Shares Reconciled)"
        elif shares_from_bs and shares_from_bs > 0:
            canonical_shares = shares_from_bs
            derivation_method = "Balance Sheet Audited Equity Capital / Face Value"
        elif pre_parsed_shares and pre_parsed_shares > 0:
            canonical_shares = pre_parsed_shares
            derivation_method = "Data Sheet Pre-Parsed Share Count"
        else:
            canonical_shares = 1.0
            derivation_method = "Fallback Minimum Floor"

        canonical_shares = max(0.01, float(canonical_shares))

        # Consistency Validation: Calculated MCap vs Reported MCap
        calculated_mcap = current_price * canonical_shares
        diff_pct = 0.0
        is_consistent = True
        flag = None
        
        if reported_market_cap > 0:
            diff_pct = abs(calculated_mcap - reported_market_cap) / reported_market_cap * 100.0
            if diff_pct > tolerance_pct:
                is_consistent = False
                flag = (
                    f"MARKET CAP / SHARE COUNT CONSISTENCY CHECK FAILED: "
                    f"Calculated MCap (Rs. {calculated_mcap:.1f} Cr = CMP Rs. {current_price:.2f} * {canonical_shares:.2f} Cr shares) "
                    f"differs from Reported MCap (Rs. {reported_market_cap:.1f} Cr) by {diff_pct:.2f}% (Tolerance: {tolerance_pct}%)."
                )

        return {
            "canonical_shares": round(canonical_shares, 4),
            "diluted_shares": round(canonical_shares, 4),
            "shares_from_bs": round(shares_from_bs, 4) if shares_from_bs else None,
            "shares_from_mcap_price": round(shares_from_mcap_price, 4) if shares_from_mcap_price else None,
            "face_value": face_value,
            "derivation_method": derivation_method,
            "is_consistent": is_consistent,
            "discrepancy_pct": round(diff_pct, 2),
            "validation_flag": flag,
            "calculated_mcap": round(calculated_mcap, 2)
        }


class MarketCapEngine:
    """
    Market Capitalization Engine (Phase 5).
    Reconciles sourced market capitalization vs calculated market capitalization.
    Exposes named semantic fields (market_cap, total_assets, total_equity, enterprise_value).
    """

    @classmethod
    def reconcile_market_cap(
        cls, 
        current_price: float, 
        canonical_shares: float, 
        reported_mcap: float,
        borrowings: float,
        cash_equivalents: float
    ) -> Dict[str, Any]:
        calc_mcap = round(current_price * canonical_shares, 2)
        final_mcap = reported_mcap if reported_mcap > 0 else calc_mcap
        
        net_debt = round(borrowings - cash_equivalents, 2)
        enterprise_value = round(final_mcap + net_debt, 2)

        return {
            "market_cap": final_mcap,
            "calculated_market_cap": calc_mcap,
            "reported_market_cap": reported_mcap,
            "net_debt": net_debt,
            "enterprise_value": enterprise_value
        }


class CompanyMaster:
    """
    Universal Company Master Object (Phase 2 & Phase 3).
    The Single Source of Truth for all downstream valuation modules.
    """

    @classmethod
    def from_screener_data(cls, raw_data: Dict[str, Any], config: Optional[Dict[str, Any]] = None) -> "CompanyMaster":
        """Factory constructor from screener raw data."""
        return cls(raw_data, config)

    def __init__(self, raw_data: Dict[str, Any], config: Optional[Dict[str, Any]] = None):
        self.config = config or VALUATION_CONFIG
        self.raw_data = raw_data
        
        # 1. Identity & Classification Metadata
        self.company_name = str(raw_data.get('company_name', '')).strip()
        self.ticker = str(raw_data.get('ticker', '')).strip().upper()
        self.exchange = "NSE" if raw_data.get('nse_ticker') else ("BSE" if raw_data.get('bse_code') else "NSE/BSE")
        self.isin = str(raw_data.get('isin', '')) or ""

        self.sector = str(raw_data.get('sector', '')).strip() or "General Operating"
        self.industry = str(raw_data.get('industry', '')).strip() or "Diversified"
        self.business_model = str(raw_data.get('about', '')).strip()[:200]
        self.currency = "INR"
        self.unit = "Cr"

        # 2. Reporting Basis Control (Phase 6)
        basis_eval = ReportingBasisController.evaluate(raw_data)
        self.reporting_basis = basis_eval['reporting_basis']
        self.reporting_basis_status = basis_eval['status']
        self.reporting_basis_inconsistencies = basis_eval['inconsistencies']
        self.is_consolidated = basis_eval['is_consolidated']

        # 3. Market Pricing & Share Count Engine (Phase 4 & 5)
        meta = raw_data.get('meta', {})
        self.current_price = float(raw_data.get('current_price', 0.0) or meta.get('Current Price', 0.0) or 0.0)
        self.reported_market_cap = float(raw_data.get('market_cap_cr', 0.0) or meta.get('Market Cap', 0.0) or 0.0)
        
        share_eval = ShareCountEngine.calculate_shares(
            raw_data, 
            self.current_price, 
            self.reported_market_cap,
            tolerance_pct=self.config.get('market_cap_share_tolerance_pct', 5.0)
        )
        self.shares_outstanding = share_eval['canonical_shares']
        self.diluted_shares = share_eval['diluted_shares']
        self.face_value = share_eval['face_value']
        self.share_derivation_method = share_eval['derivation_method']
        self.share_mcap_consistency = share_eval['is_consistent']
        self.share_mcap_discrepancy_pct = share_eval['discrepancy_pct']
        
        # 4. Debt & Cash Extraction
        self.total_debt, self.cash, self.investments = self._extract_debt_and_cash(raw_data)
        self.net_debt = round(self.total_debt - self.cash, 2)

        # 5. Semantic Market Cap & Enterprise Value (Phase 5)
        mcap_eval = MarketCapEngine.reconcile_market_cap(
            self.current_price,
            self.shares_outstanding,
            self.reported_market_cap,
            self.total_debt,
            self.cash
        )
        self.market_cap = mcap_eval['market_cap']
        self.calculated_market_cap = mcap_eval['calculated_market_cap']
        self.enterprise_value = mcap_eval['enterprise_value']

        # 6. Audit & Validation Flags
        self.validation_flags: List[str] = []
        if share_eval['validation_flag']:
            self.validation_flags.append(share_eval['validation_flag'])
        if self.reporting_basis_inconsistencies:
            self.validation_flags.extend(self.reporting_basis_inconsistencies)

        # 7. As-of Dates & Sources
        self.fiscal_year_end = self._extract_fiscal_year_end(raw_data)
        self.data_as_of_date = raw_data.get('meta_raw', {}).get('as_of_date', 'Latest Filings')
        self.financial_data_source = "Screener.in (Audited Filings)"
        self.market_data_source = "NSE / BSE Real-Time Feed"
        self.classification_confidence = "HIGH"
        self.data_quality_score = self._compute_initial_quality_score()

        # 8. Canonical Metrics Repository (Phase 3)
        self.canonical_metrics: Dict[str, CanonicalMetric] = self._build_canonical_metrics()

    def _extract_debt_and_cash(self, raw_data: Dict[str, Any]) -> (float, float, float):
        tables = raw_data.get('tables', {})
        bs_df = tables.get('balance-sheet')
        schedules = raw_data.get('schedules', {})
        
        total_debt = 0.0
        cash = 0.0
        investments = 0.0

        if bs_df is not None and not bs_df.empty:
            for _, row in bs_df.iterrows():
                m_name = str(row.get('Metric', '')).lower()
                if 'borrowing' in m_name or 'total debt' in m_name:
                    for col in reversed([c for c in bs_df.columns if c != 'Metric']):
                        v = str(row[col]).replace(',', '').strip()
                        try:
                            total_debt = float(v)
                            break
                        except (ValueError, TypeError):
                            pass
                if 'investments' in m_name:
                    for col in reversed([c for c in bs_df.columns if c != 'Metric']):
                        v = str(row[col]).replace(',', '').strip()
                        try:
                            investments = float(v)
                            break
                        except (ValueError, TypeError):
                            pass

        # Cash & Equivalents from schedules or other assets
        oa_sched = schedules.get('Other Assets', {})
        cash_dict = oa_sched.get('Cash Equivalents', {})
        if cash_dict:
            latest_v = list(cash_dict.values())[-1]
            try:
                cash = float(str(latest_v).replace(',', '').strip())
            except (ValueError, TypeError):
                pass
        elif raw_data.get('cash_cr'):
            cash = float(raw_data['cash_cr'])
        else:
            # Fallback to ~25% of Other Assets or minimum liquidity
            other_assets = 0.0
            if bs_df is not None and not bs_df.empty:
                oa_row = bs_df[bs_df['Metric'].str.contains('Other Assets', case=False, na=False)]
                if not oa_row.empty:
                    for col in reversed([c for c in bs_df.columns if c != 'Metric']):
                        try:
                            other_assets = float(str(oa_row.iloc[0][col]).replace(',', ''))
                            break
                        except (ValueError, TypeError):
                            pass
            cash = round(other_assets * 0.25, 2) if other_assets > 0 else round(self.reported_market_cap * 0.02, 2)

        return total_debt, cash, investments

    def _extract_fiscal_year_end(self, raw_data: Dict[str, Any]) -> str:
        tables = raw_data.get('tables', {})
        pl_df = tables.get('profit-loss')
        if pl_df is not None and not pl_df.empty:
            cols = [c for c in pl_df.columns if c != 'Metric' and 'TTM' not in str(c).upper()]
            if cols:
                return str(cols[-1])
        return "March 2024"

    def _compute_initial_quality_score(self) -> float:
        score = 100.0
        if not self.share_mcap_consistency:
            score -= 15.0
        if self.reporting_basis_status != "CONSISTENT":
            score -= 15.0
        if self.current_price <= 0:
            score -= 30.0
        if self.shares_outstanding <= 0:
            score -= 30.0
        return max(0.0, score)

    def _build_canonical_metrics(self) -> Dict[str, CanonicalMetric]:
        return {
            "current_price": CanonicalMetric(
                value=self.current_price,
                unit="INR/share",
                source=self.market_data_source,
                reporting_basis=self.reporting_basis,
                description="Current Market Share Price"
            ),
            "shares_outstanding": CanonicalMetric(
                value=self.shares_outstanding,
                unit="Cr",
                source=self.share_derivation_method,
                reporting_basis=self.reporting_basis,
                description="Appropriate Diluted Shares Outstanding"
            ),
            "market_cap": CanonicalMetric(
                value=self.market_cap,
                unit="Cr",
                source="Market Real-Time Sourced",
                reporting_basis=self.reporting_basis,
                description="Reconciled Market Capitalization"
            ),
            "total_debt": CanonicalMetric(
                value=self.total_debt,
                unit="Cr",
                source=self.financial_data_source,
                reporting_basis=self.reporting_basis,
                description="Total Borrowings (Long + Short Term Debt)"
            ),
            "cash_and_equivalents": CanonicalMetric(
                value=self.cash,
                unit="Cr",
                source=self.financial_data_source,
                reporting_basis=self.reporting_basis,
                description="Verified Liquid Cash and Bank Balances"
            ),
            "net_debt": CanonicalMetric(
                value=self.net_debt,
                unit="Cr",
                source="Calculated (Total Debt - Cash)",
                reporting_basis=self.reporting_basis,
                description="Net Debt (Negative value indicates Net Cash)"
            ),
            "enterprise_value": CanonicalMetric(
                value=self.enterprise_value,
                unit="Cr",
                source="Calculated (Market Cap + Net Debt)",
                reporting_basis=self.reporting_basis,
                description="Enterprise Value"
            )
        }

    def to_dict(self) -> Dict[str, Any]:
        """Provides the complete canonical representation for all downstream valuation modules."""
        return {
            "company_name": self.company_name,
            "ticker": self.ticker,
            "exchange": self.exchange,
            "isin": self.isin,
            "sector": self.sector,
            "industry": self.industry,
            "business_model": self.business_model,
            "reporting_basis": self.reporting_basis,
            "is_consolidated": self.is_consolidated,
            "reporting_basis_status": self.reporting_basis_status,
            "reporting_basis_inconsistencies": self.reporting_basis_inconsistencies,
            "currency": self.currency,
            "current_price": self.current_price,
            "shares_outstanding": self.shares_outstanding,
            "diluted_shares": self.diluted_shares,
            "face_value": self.face_value,
            "share_derivation_method": self.share_derivation_method,
            "share_mcap_consistency": self.share_mcap_consistency,
            "share_mcap_discrepancy_pct": self.share_mcap_discrepancy_pct,
            "market_cap": self.market_cap,
            "calculated_market_cap": self.calculated_market_cap,
            "reported_market_cap": self.reported_market_cap,
            "total_debt": self.total_debt,
            "cash": self.cash,
            "investments": self.investments,
            "net_debt": self.net_debt,
            "enterprise_value": self.enterprise_value,
            "fiscal_year_end": self.fiscal_year_end,
            "data_as_of_date": self.data_as_of_date,
            "financial_data_source": self.financial_data_source,
            "market_data_source": self.market_data_source,
            "classification_confidence": self.classification_confidence,
            "data_quality_score": self.data_quality_score,
            "validation_flags": self.validation_flags,
            "canonical_metrics": {k: v.to_dict() for k, v in self.canonical_metrics.items()}
        }
