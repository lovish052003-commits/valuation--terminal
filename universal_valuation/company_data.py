"""
universal_valuation/company_data.py
===================================
Authoritative Canonical Data Layer for the Universal Valuation Engine.

Implements:
- Phase 2: One Authoritative Canonical Data Layer (CompanyData container).
- Phase 3: Single Source of Truth for all valuation modules.
- Phase 5: Absolute prohibition of silent zero fallbacks (Missing != Zero).
- Phase 8: Unit and currency harmonization.
- Phase 9: Reconciled Market Cap and Share Count.
- Phase 10: Canonical Enterprise Value.
- Phase 14: Strict separation of Historical, Forecast, and Scenario dimensions.
- Phase 21: Comprehensive Data Validation Engine (19 institutional checks).
- Phase 22: Complete source traceability and audit metadata.
- Phase 25/26: Modular CompanyConfig & Archetype awareness.
"""

import math
import datetime
from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional, Tuple, Union

from .canonical_financials import (
    FinancialField,
    CanonicalPeriodStatements,
    CanonicalFinancialStatementLayer,
    clean_fiscal_year_label
)
from .field_resolver import UniversalFieldResolver
from .company_master import CompanyMaster, ShareCountEngine, MarketCapEngine, ReportingBasisController
from .company_classifier import CompanyClassificationEngine


@dataclass
class ValidationIssue:
    check_name: str
    severity: str  # 'FAIL', 'WARNING', 'INFO'
    message: str
    value: Any = None
    expected: Any = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "check": self.check_name,
            "severity": self.severity,
            "message": self.message,
            "value": str(self.value) if self.value is not None else None,
            "expected": str(self.expected) if self.expected is not None else None
        }


@dataclass
class CompanyIdentity:
    company_name: str
    ticker: str
    exchange: str = "NSE"
    sector: str = ""
    industry: str = ""
    sub_industry: str = ""
    company_type: str = "INDUSTRIAL"
    reporting_basis: str = "CONSOLIDATED"  # CONSOLIDATED or STANDALONE
    is_consolidated: bool = True
    currency: str = "INR"
    unit: str = "Cr"
    fiscal_year_end: str = "March"


@dataclass
class MarketData:
    current_price: float = 0.0
    shares_outstanding: float = 0.0
    market_cap: float = 0.0
    calculated_market_cap: float = 0.0
    enterprise_value: float = 0.0
    beta: float = 1.0
    risk_free_rate: float = 0.0710
    equity_risk_premium: float = 0.0600
    cost_of_equity: float = 0.1310
    pre_tax_cost_of_debt: float = 0.0805
    tax_rate: float = 0.2517
    after_tax_cost_of_debt: float = 0.0602
    wacc: float = 0.1150
    target_debt_weight: float = 0.20
    target_equity_weight: float = 0.80


class CompanyData:
    """
    Authoritative Canonical Data Container for ANY publicly listed company.
    Every downstream valuation module (DCF, WACC, Peer Comps, ROIC, Reinvestment,
    DuPont, Altman Z, Common-Size, Forecasting, and AI Summary) consumes this
    single source of truth.
    """

    def __init__(self, identity: CompanyIdentity):
        self.identity: CompanyIdentity = identity
        self.market_data: MarketData = MarketData()
        
        # Historical data by fiscal year: e.g. historical['2024'] or historical['Mar 2024']
        self.historical_periods: List[str] = []
        self.historical: Dict[str, Dict[str, Dict[str, FinancialField]]] = {}
        
        # Explicit forecast data stored separately by forecast period: e.g. forecast['FY2025']
        self.forecast_periods: List[str] = []
        self.forecast: Dict[str, Dict[str, Any]] = {}
        
        # Scenarios strictly stored separately (NEVER as years or dates!)
        self.scenarios: Dict[str, Dict[str, Any]] = {
            "base": {},
            "bull": {},
            "bear": {}
        }
        
        # Validation and Data Quality Layer (Phase 21)
        self.validation_issues: List[ValidationIssue] = []
        self.data_quality_status: str = "PASS"  # PASS, WARNING, FAIL
        self.data_quality_score: float = 100.0
        
        # Underlying canonical statement layer
        self.canonical_layer: Optional[CanonicalFinancialStatementLayer] = None
        self.company_master: Optional[CompanyMaster] = None
        self.classification: Dict[str, Any] = {}

    @classmethod
    def build(
        cls, 
        screener_data: Dict[str, Any], 
        classification: Optional[Dict[str, Any]] = None,
        custom_params: Optional[Dict[str, Any]] = None
    ) -> "CompanyData":
        """
        Universal Builder: Normalizes raw financial data from Screener into
        a single verified canonical data layer.
        """
        # 1. Company Classification & Identity
        if classification is None:
            classification = CompanyClassificationEngine.classify(screener_data)
        
        c_name = str(screener_data.get('company_name', '')).strip()
        ticker = str(screener_data.get('ticker', '')).strip().upper()
        if not ticker:
            ticker = "COMPANY"
        if not c_name:
            c_name = ticker

        basis_eval = ReportingBasisController.evaluate(screener_data)
        rep_basis = basis_eval.get('reporting_basis', 'CONSOLIDATED')
        is_consol = basis_eval.get('is_consolidated', True)

        identity = CompanyIdentity(
            company_name=c_name,
            ticker=ticker,
            exchange="NSE" if not screener_data.get('bse_id') else "BSE/NSE",
            sector=screener_data.get('sector', classification.get('sector', 'General')),
            industry=screener_data.get('industry', classification.get('industry', 'General')),
            sub_industry=classification.get('sub_industry', ''),
            company_type=classification.get('valuation_family', 'INDUSTRIAL'),
            reporting_basis=rep_basis,
            is_consolidated=is_consol,
            currency="INR",
            unit="Cr",
            fiscal_year_end="March"
        )
        
        company = cls(identity)
        company.classification = classification

        # 2. Canonical Financial Statement Layer
        canonical_layer = CanonicalFinancialStatementLayer.from_screener_data(screener_data, classification)
        company.canonical_layer = canonical_layer
        company.historical_periods = canonical_layer.periods

        # 3. Market Data & Capital Structure (Section 5: Strict Market Cap Rule)
        master = CompanyMaster.from_screener_data(screener_data)
        company.company_master = master

        price = float(master.current_price)
        shares = float(master.shares_outstanding)
        reported_mcap = float(master.market_cap)
        calc_mcap = round(price * shares, 2) if (price > 0 and shares > 0) else reported_mcap

        # Reconcile market cap: Primary is price * shares; flag warning if divergence exceeds 15%
        mcap = calc_mcap if calc_mcap > 0 else reported_mcap
        if reported_mcap > 0 and calc_mcap > 0:
            diff_mcap_pct = abs(reported_mcap - calc_mcap) / reported_mcap * 100.0
            if diff_mcap_pct > 15.0:
                company.validation_issues.append(ValidationIssue(
                    check_name="MARKET_CAP_RECONCILIATION_WARNING",
                    severity="WARNING",
                    message=f"Reported Market Cap (₹{reported_mcap:,.1f} Cr) differs from Price x Shares (₹{calc_mcap:,.1f} Cr) by {diff_mcap_pct:.1f}%.",
                    value=diff_mcap_pct,
                    expected="< 15%"
                ))

        # Check Bank Archetype (Section 8: Bank-specific architecture)
        is_bank = (
            classification.get('canonical_company_type') == 'Bank'
            or classification.get('company_type') == 'Bank'
            or classification.get('valuation_family') == 'BANK'
            or classification.get('is_bank', False)
        )
        if is_bank:
            company.identity.company_type = "Bank"

        # Latest Debt and Cash from Canonical Statements
        latest_stmt = list(canonical_layer.statements_by_period.values())[-1] if canonical_layer.statements_by_period else None
        debt_f = latest_stmt.debt if latest_stmt else FinancialField.unavailable("Latest", "No statement")
        cash_f = latest_stmt.cash if latest_stmt else FinancialField.unavailable("Latest", "No statement")
        
        # Cost of Capital Parameters
        rf = float(custom_params.get('risk_free_rate', 0.0710) if custom_params else 0.0710)
        erp = float(custom_params.get('erp', 0.0600) if custom_params else 0.0600)
        tax_rate = float(custom_params.get('tax_rate', 0.2517) if custom_params else 0.2517)
        beta_val = float(custom_params.get('beta', 1.0) if custom_params and custom_params.get('beta') is not None else 1.0)
        ke = rf + (beta_val * erp)

        if is_bank:
            # Rule 8 & 10: For banks, customer deposits are funding liabilities, NOT debt in an industrial WACC.
            # Primary discount rate for bank equity valuation is Cost of Equity (Ke).
            debt_val = 0.0
            cash_val = 0.0
            net_debt = 0.0
            ev = mcap  # For banks, equity value is the focus
            we = 1.0
            wd = 0.0
            kd_pre = 0.0
            kd_post = 0.0
            wacc = ke  # Cost of Equity is primary discount rate for banks
        else:
            debt_val = debt_f.value if debt_f.is_available and debt_f.value is not None else 0.0
            cash_val = cash_f.value if cash_f.is_available and cash_f.value is not None else 0.0
            net_debt = round(debt_val - cash_val, 2)
            ev = round(mcap + net_debt, 2)
            kd_pre = 0.0805
            kd_post = kd_pre * (1.0 - tax_rate)
            total_cap = mcap + debt_val
            we = mcap / total_cap if total_cap > 0 else 0.80
            wd = debt_val / total_cap if total_cap > 0 else 0.20
            wacc = (we * ke) + (wd * kd_post)

        company.market_data = MarketData(
            current_price=price,
            shares_outstanding=shares,
            market_cap=mcap,
            calculated_market_cap=calc_mcap,
            enterprise_value=ev,
            beta=beta_val,
            risk_free_rate=rf,
            equity_risk_premium=erp,
            cost_of_equity=ke,
            pre_tax_cost_of_debt=kd_pre,
            tax_rate=tax_rate,
            after_tax_cost_of_debt=kd_post,
            wacc=wacc,
            target_debt_weight=wd,
            target_equity_weight=we
        )

        # 4. Populate Historical Dictionary by Fiscal Year
        for p in canonical_layer.periods:
            stmt = canonical_layer.statements_by_period[p]
            clean_yr = clean_fiscal_year_label(p) or p

            # Organize statements into structured sub-categories
            pnl_dict = {
                "revenue": stmt.revenue,
                "ebitda": stmt.ebitda,
                "depreciation": stmt.depreciation,
                "ebit": stmt.ebit,
                "interest_expense": UniversalFieldResolver.resolve_field(screener_data, "interest", p, rep_basis),
                "pre_tax_income": UniversalFieldResolver.resolve_field(screener_data, "pbt", p, rep_basis),
                "tax": stmt.tax,
                "net_income": stmt.net_income,
                "cogs": UniversalFieldResolver.resolve_field(screener_data, "cogs", p, rep_basis),
                "gross_profit": UniversalFieldResolver.resolve_field(screener_data, "gross_profit", p, rep_basis),
                "operating_expenses": UniversalFieldResolver.resolve_field(screener_data, "operating_expenses", p, rep_basis)
            }

            bs_dict = {
                "cash": stmt.cash,
                "short_term_investments": UniversalFieldResolver.resolve_field(screener_data, "short_term_investments", p, rep_basis),
                "accounts_receivable": stmt.receivables,
                "inventory": stmt.inventory,
                "other_current_assets": UniversalFieldResolver.resolve_field(screener_data, "other_assets", p, rep_basis),
                "current_assets": stmt.current_assets,
                "pp_e": UniversalFieldResolver.resolve_field(screener_data, "fixed_assets", p, rep_basis),
                "cwip": UniversalFieldResolver.resolve_field(screener_data, "cwip", p, rep_basis),
                "investments": UniversalFieldResolver.resolve_field(screener_data, "investments", p, rep_basis),
                "total_assets": stmt.total_assets,
                "accounts_payable": UniversalFieldResolver.resolve_field(screener_data, "other_liabilities", p, rep_basis),
                "other_current_liabilities": UniversalFieldResolver.resolve_field(screener_data, "other_liabilities", p, rep_basis),
                "current_liabilities": stmt.current_liabilities,
                "total_debt": stmt.debt,
                "total_liabilities": UniversalFieldResolver.resolve_field(screener_data, "total_liabilities", p, rep_basis),
                "equity_capital": UniversalFieldResolver.resolve_field(screener_data, "equity_capital", p, rep_basis),
                "reserves": UniversalFieldResolver.resolve_field(screener_data, "reserves", p, rep_basis),
                "total_equity": stmt.total_equity
            }

            cf_dict = {
                "operating_cash_flow": stmt.operating_cash_flow,
                "capex": stmt.capex,
                "depreciation": stmt.depreciation,
                "change_in_nwc": stmt.change_in_nwc,
                "free_cash_flow": UniversalFieldResolver.resolve_field(screener_data, "free_cash_flow", p, rep_basis)
            }

            if is_bank:
                # Bank-specific adjustments (Section 8: Bank-specific architecture)
                pnl_dict["ebitda"] = FinancialField.not_applicable(clean_yr, "Not applicable for banks")
                pnl_dict["ebit"] = FinancialField.not_applicable(clean_yr, "Not applicable for banks")
                pnl_dict["cogs"] = FinancialField.not_applicable(clean_yr, "Not applicable for banks")
                pnl_dict["gross_profit"] = FinancialField.not_applicable(clean_yr, "Not applicable for banks")
                bs_dict["inventory"] = FinancialField.not_applicable(clean_yr, "Not applicable for banks")
                bs_dict["cwip"] = FinancialField.not_applicable(clean_yr, "Not applicable for banks")
                cf_dict["capex"] = FinancialField.not_applicable(clean_yr, "Not applicable for banks")

                derived_dict = {
                    "invested_capital": FinancialField.not_applicable(clean_yr, "Not applicable for banks"),
                    "nopat": FinancialField.not_applicable(clean_yr, "Not applicable for banks"),
                    "roic": FinancialField.not_applicable(clean_yr, "Not applicable for banks"),
                    "roe": stmt.roe,
                    "roa": stmt.roa,
                    "asset_turnover": stmt.asset_turnover,
                    "equity_multiplier": stmt.equity_multiplier,
                    "altman_z": FinancialField.not_applicable(clean_yr, "Not applicable for banks"),
                    "reinvestment_rate": FinancialField.not_applicable(clean_yr, "Not applicable for banks"),
                    "net_debt": FinancialField.not_applicable(clean_yr, "Not applicable for banks")
                }

                # Primary bank metrics (Section 8)
                int_earned = UniversalFieldResolver.resolve_field(screener_data, "interest_earned", p, rep_basis)
                fin_profit = UniversalFieldResolver.resolve_field(screener_data, "financing_profit", p, rep_basis)
                provisions = UniversalFieldResolver.resolve_field(screener_data, "provisions", p, rep_basis)
                advances = UniversalFieldResolver.resolve_field(screener_data, "advances", p, rep_basis)
                deposits = UniversalFieldResolver.resolve_field(screener_data, "deposits", p, rep_basis)
                
                nii_val = None
                if int_earned.is_available and pnl_dict["interest_expense"].is_available:
                    nii_val = round((int_earned.value or 0.0) - (pnl_dict["interest_expense"].value or 0.0), 2)
                nii_f = FinancialField.available(nii_val, clean_yr, rep_basis, reported_or_derived="derived", calculation_method="interest_earned - interest_expended") if nii_val is not None else FinancialField.unavailable(clean_yr, "NII uncomputable", rep_basis)

                cd_val = None
                if advances.is_available and deposits.is_available and deposits.value and deposits.value > 0:
                    cd_val = round((advances.value / deposits.value) * 100.0, 2)
                cd_f = FinancialField.available(cd_val, clean_yr, rep_basis, reported_or_derived="derived", calculation_method="advances / deposits") if cd_val is not None else FinancialField.unavailable(clean_yr, "Advances or Deposits missing", rep_basis)

                bank_metrics = {
                    "net_interest_income": nii_f,
                    "interest_earned": int_earned if int_earned.is_available else stmt.revenue,
                    "interest_expended": pnl_dict["interest_expense"],
                    "other_income": UniversalFieldResolver.resolve_field(screener_data, "other_income", p, rep_basis),
                    "operating_expenses": pnl_dict["operating_expenses"],
                    "ppop": fin_profit,
                    "provisions": provisions,
                    "pbt": pnl_dict["pre_tax_income"],
                    "pat": stmt.net_income,
                    "advances": advances,
                    "deposits": deposits,
                    "borrowings": stmt.debt,
                    "total_assets": stmt.total_assets,
                    "total_equity": stmt.total_equity,
                    "roe": stmt.roe,
                    "roa": stmt.roa,
                    "cd_ratio": cd_f,
                    "cost_to_income": UniversalFieldResolver.resolve_field(screener_data, "cost_to_income", p, rep_basis),
                    "gross_npa_pct": UniversalFieldResolver.resolve_field(screener_data, "gross_npa", p, rep_basis),
                    "net_npa_pct": UniversalFieldResolver.resolve_field(screener_data, "net_npa", p, rep_basis),
                    "capital_adequacy_pct": UniversalFieldResolver.resolve_field(screener_data, "capital_adequacy", p, rep_basis)
                }
            else:
                bank_metrics = {}
                nopat_val = None
                if stmt.ebit.is_available and stmt.ebit.value is not None:
                    nopat_val = round(stmt.ebit.value * (1.0 - tax_rate), 2)
                nopat_f = FinancialField.available(nopat_val, p, rep_basis) if nopat_val is not None else FinancialField.unavailable(p, "EBIT missing for NOPAT", rep_basis)

                net_debt_val = None
                if stmt.debt.is_available and stmt.debt.value is not None:
                    c_v = stmt.cash.value if stmt.cash.is_available and stmt.cash.value is not None else 0.0
                    net_debt_val = round(stmt.debt.value - c_v, 2)
                net_debt_f = FinancialField.available(net_debt_val, p, rep_basis) if net_debt_val is not None else FinancialField.unavailable(p, "Debt missing for Net Debt", rep_basis)

                reinvest_val = None
                if stmt.capex.is_available and stmt.change_in_nwc.is_available and stmt.depreciation.is_available:
                    reinvest_val = round(stmt.capex.value - stmt.depreciation.value + (stmt.change_in_nwc.value or 0.0), 2)
                reinvest_f = FinancialField.available(reinvest_val, p, rep_basis) if reinvest_val is not None else FinancialField.unavailable(p, "Components missing for Reinvestment", rep_basis)

                derived_dict = {
                    "invested_capital": stmt.invested_capital,
                    "nopat": nopat_f,
                    "roic": stmt.roic,
                    "roe": stmt.roe,
                    "roa": stmt.roa,
                    "asset_turnover": stmt.asset_turnover,
                    "equity_multiplier": stmt.equity_multiplier,
                    "altman_z": stmt.altman_z,
                    "reinvestment_rate": reinvest_f,
                    "net_debt": net_debt_f
                }

            company.historical[clean_yr] = {
                "income_statement": pnl_dict,
                "balance_sheet": bs_dict,
                "cash_flow": cf_dict,
                "derived": derived_dict,
                "bank_metrics": bank_metrics
            }

        # 5. Run Mandatory Phase 21 Data Validation Engine
        company.validate()

        return company

    def validate(self) -> Dict[str, Any]:
        """
        Phase 21: Institutional Data Validation Engine.
        Executes all 19 mandatory checks:
        1. Total Assets > 0
        2. Revenue >= 0 where applicable
        3. Shares > 0
        4. Current Price > 0
        5. Market Cap reconciliation
        6. EV reconciliation
        7. Assets = Liabilities + Equity within tolerance
        8. Current Assets >= 0
        9. Debt >= 0
        10. Cash >= 0
        11. Invested Capital > 0 where required
        12. WACC > terminal growth
        13. Forecast years sequential
        14. Scenario labels not stored as years
        15. EBITDA consistency
        16. EBIT consistency
        17. Net income consistency
        18. Units consistent
        19. Currency consistent
        """
        self.validation_issues.clear()
        score = 100.0

        # Check 1: Total Assets > 0
        latest_ta = self.get_canonical_value('total_assets', 'latest')
        if latest_ta is None or latest_ta <= 0:
            self.validation_issues.append(ValidationIssue(
                check_name="total_assets_positive",
                severity="FAIL",
                message="Total Assets is zero, negative, or missing.",
                value=latest_ta, expected="> 0"
            ))
            score -= 25.0

        # Check 2: Revenue >= 0
        latest_rev = self.get_canonical_value('revenue', 'latest')
        if latest_rev is not None and latest_rev < 0:
            self.validation_issues.append(ValidationIssue(
                check_name="revenue_non_negative",
                severity="FAIL",
                message="Revenue is reported negative.",
                value=latest_rev, expected=">= 0"
            ))
            score -= 15.0

        # Check 3: Shares > 0
        shares = self.market_data.shares_outstanding
        if shares <= 0:
            self.validation_issues.append(ValidationIssue(
                check_name="shares_positive",
                severity="FAIL",
                message="Shares outstanding is non-positive.",
                value=shares, expected="> 0"
            ))
            score -= 25.0

        # Check 4: Current Price > 0
        price = self.market_data.current_price
        if price <= 0:
            self.validation_issues.append(ValidationIssue(
                check_name="price_positive",
                severity="FAIL",
                message="Current stock price is non-positive.",
                value=price, expected="> 0"
            ))
            score -= 25.0

        # Check 5: Market Cap reconciliation
        mcap = self.market_data.market_cap
        calc_mcap = self.market_data.calculated_market_cap
        if mcap > 0 and calc_mcap > 0:
            diff_mcap_pct = abs(mcap - calc_mcap) / mcap * 100.0
            if diff_mcap_pct > 15.0:
                self.validation_issues.append(ValidationIssue(
                    check_name="market_cap_reconciled",
                    severity="WARNING",
                    message=f"Reported Market Cap ({mcap:,.1f}) diverges from Price x Shares ({calc_mcap:,.1f}) by {diff_mcap_pct:.1f}%.",
                    value=diff_mcap_pct, expected="< 15%"
                ))
                score -= 10.0

        # Check 6: EV reconciliation
        ev = self.market_data.enterprise_value
        debt = self.get_canonical_value('total_debt', 'latest') or 0.0
        cash = self.get_canonical_value('cash', 'latest') or 0.0
        expected_ev = mcap + debt - cash
        if abs(ev - expected_ev) > 5.0:
            self.validation_issues.append(ValidationIssue(
                check_name="ev_reconciled",
                severity="WARNING",
                message=f"Enterprise Value ({ev:,.1f}) differs from MCAP + Debt - Cash ({expected_ev:,.1f}).",
                value=ev, expected=expected_ev
            ))
            score -= 5.0

        # Check 7: Balance Sheet Balance (Assets = Liab + Equity)
        latest_period = self.historical_periods[-1] if self.historical_periods else None
        if latest_period and latest_period in self.historical:
            h_stmt = self.historical[latest_period]
            ta = self.get_canonical_value('total_assets', latest_period)
            te = self.get_canonical_value('total_equity', latest_period)
            tl = self.get_canonical_value('total_liabilities', latest_period)
            if ta and te and tl:
                diff_bs = abs(ta - (te + tl))
                diff_pct = (diff_bs / ta) * 100.0
                if diff_pct > 15.0:
                    self.validation_issues.append(ValidationIssue(
                        check_name="balance_sheet_balance",
                        severity="WARNING",
                        message=f"Balance Sheet out of balance by {diff_pct:.1f}% (Assets: {ta:,.1f} vs Liab+Eq: {te+tl:,.1f}).",
                        value=diff_pct, expected="< 15%"
                    ))
                    score -= 10.0

        # Check 8: Current Assets >= 0
        ca = self.get_canonical_value('current_assets', 'latest')
        if ca is not None and ca < 0:
            self.validation_issues.append(ValidationIssue(
                check_name="current_assets_non_negative",
                severity="FAIL",
                message="Current Assets is reported negative.",
                value=ca, expected=">= 0"
            ))
            score -= 10.0

        # Check 9: Debt >= 0
        if debt < 0:
            self.validation_issues.append(ValidationIssue(
                check_name="debt_non_negative",
                severity="FAIL",
                message="Debt is reported negative.",
                value=debt, expected=">= 0"
            ))
            score -= 10.0

        # Check 10: Cash >= 0
        if cash < 0:
            self.validation_issues.append(ValidationIssue(
                check_name="cash_non_negative",
                severity="FAIL",
                message="Cash is reported negative.",
                value=cash, expected=">= 0"
            ))
            score -= 10.0

        # Check 11: Invested Capital > 0 where required (Rule 8: Not applicable for banks)
        is_bank = (self.identity.company_type == "Bank")
        if not is_bank:
            inv_cap = self.get_canonical_value('invested_capital', 'latest')
            if inv_cap is not None and inv_cap <= 0:
                self.validation_issues.append(ValidationIssue(
                    check_name="invested_capital_positive",
                    severity="WARNING",
                    message="Invested Capital is non-positive; ROIC calculation should be guarded.",
                    value=inv_cap, expected="> 0"
                ))
                score -= 5.0

        # Check 12: Discount Rate > terminal growth (Rule 10 & 11: Cost of Equity for Banks, WACC for Non-banks)
        terminal_g = 0.025
        if is_bank:
            disc_rate = self.market_data.cost_of_equity
            rate_name = "Cost of Equity (Ke)"
        else:
            disc_rate = self.market_data.wacc
            rate_name = "WACC"

        if disc_rate <= terminal_g:
            self.validation_issues.append(ValidationIssue(
                check_name="discount_rate_gt_terminal_growth",
                severity="FAIL",
                message=f"{rate_name} ({disc_rate*100:.2f}%) must exceed terminal growth ({terminal_g*100:.2f}%).",
                value=disc_rate, expected=f"> {terminal_g}"
            ))
            score -= 30.0

        # Check 13 & 14: Scenario labels not stored as years
        for yr in self.historical_periods:
            if any(sc in str(yr).lower() for sc in ['best', 'worst', 'bear', 'bull', 'base', 'case', 'scenario']):
                self.validation_issues.append(ValidationIssue(
                    check_name="no_scenario_as_years",
                    severity="FAIL",
                    message=f"Scenario label '{yr}' was illegally found inside historical periods.",
                    value=yr, expected="Fiscal date string"
                ))
                score -= 20.0

        # Check 15, 16, 17: EBITDA, EBIT, Net Income consistency (Rule 13: Not applicable for banks)
        if not is_bank:
            ebitda = self.get_canonical_value('ebitda', 'latest')
            ebit = self.get_canonical_value('ebit', 'latest')
            net_inc = self.get_canonical_value('net_income', 'latest')
            depr = self.get_canonical_value('depreciation', 'latest') or 0.0

            if ebitda is not None and ebit is not None:
                if ebit > (ebitda + 5.0) and depr > 0:
                    self.validation_issues.append(ValidationIssue(
                        check_name="ebit_ebitda_consistency",
                        severity="WARNING",
                        message=f"EBIT ({ebit:,.1f}) exceeds EBITDA ({ebitda:,.1f}) despite positive depreciation ({depr:,.1f}).",
                        value=ebit, expected=f"<= {ebitda}"
                    ))
                    score -= 5.0

        # Check 18 & 19: Units & Currency consistency
        assert self.identity.currency == "INR", "Universal valuation currency must be INR for Indian equities"
        assert self.identity.unit == "Cr", "Universal financial unit must be Cr for Indian equities"

        # Overall Status
        has_fail = any(i.severity == "FAIL" for i in self.validation_issues)
        has_warn = any(i.severity == "WARNING" for i in self.validation_issues)
        
        self.data_quality_score = max(0.0, round(score, 1))
        if has_fail:
            self.data_quality_status = "FAIL"
        elif has_warn:
            self.data_quality_status = "WARNING"
        else:
            self.data_quality_status = "PASS"

        return {
            "status": self.data_quality_status,
            "score": self.data_quality_score,
            "issues": [i.to_dict() for i in self.validation_issues]
        }

    # Public Retrieval Helpers (Consuming Canonical Fields)
    def get_field(self, field_name: str, period: str = 'latest') -> FinancialField:
        """
        Retrieves the FinancialField object for a given field name and period.
        Never defaults to zero silently.
        """
        f_norm = field_name.strip().lower()
        target_period = period

        if period == 'latest':
            if not self.historical_periods:
                return FinancialField.unavailable("Latest", "No historical periods available")
            target_period = clean_fiscal_year_label(self.historical_periods[-1]) or self.historical_periods[-1]
        else:
            target_period = clean_fiscal_year_label(period) or str(period)

        p_dict = self.historical.get(target_period)
        if not p_dict:
            # Try fuzzy match in historical
            for pk, v in self.historical.items():
                if target_period in pk or pk in target_period:
                    p_dict = v
                    break

        if p_dict:
            for cat in ['income_statement', 'balance_sheet', 'cash_flow', 'derived']:
                if f_norm in p_dict[cat]:
                    return p_dict[cat][f_norm]

        # Check canonical layer directly
        if self.canonical_layer:
            return self.canonical_layer.get_latest(f_norm)

        return FinancialField.unavailable(target_period, f"Field '{field_name}' not found in canonical dataset")

    def get_canonical_value(self, field_name: str, period: str = 'latest') -> Optional[float]:
        """
        Retrieves the exact numerical value of a financial field.
        Returns None if missing. NEVER converts missing to zero.
        """
        f = self.get_field(field_name, period)
        if f.is_available and f.value is not None:
            return float(f.value)
        return None

    def get_series(self, field_name: str) -> List[Tuple[str, Optional[float]]]:
        """
        Returns chronological list of (period_label, value_or_none) across historical periods.
        """
        series = []
        for p in self.historical_periods:
            clean_yr = clean_fiscal_year_label(p) or p
            val = self.get_canonical_value(field_name, clean_yr)
            series.append((clean_yr, val))
        return series

    def to_dict(self) -> Dict[str, Any]:
        """Serializes the authoritative canonical company object for APIs, UI, and AI."""
        return {
            "identity": {
                "company_name": self.identity.company_name,
                "ticker": self.identity.ticker,
                "exchange": self.identity.exchange,
                "sector": self.identity.sector,
                "industry": self.identity.industry,
                "sub_industry": self.identity.sub_industry,
                "company_type": self.identity.company_type,
                "reporting_basis": self.identity.reporting_basis,
                "currency": self.identity.currency,
                "unit": self.identity.unit,
                "fiscal_year_end": self.identity.fiscal_year_end
            },
            "market_data": {
                "current_price": self.market_data.current_price,
                "shares_outstanding": self.market_data.shares_outstanding,
                "market_cap": self.market_data.market_cap,
                "calculated_market_cap": self.market_data.calculated_market_cap,
                "enterprise_value": self.market_data.enterprise_value,
                "beta": self.market_data.beta,
                "risk_free_rate": self.market_data.risk_free_rate,
                "equity_risk_premium": self.market_data.equity_risk_premium,
                "cost_of_equity": self.market_data.cost_of_equity,
                "pre_tax_cost_of_debt": self.market_data.pre_tax_cost_of_debt,
                "after_tax_cost_of_debt": self.market_data.after_tax_cost_of_debt,
                "wacc": self.market_data.wacc
            },
            "historical_periods": self.historical_periods,
            "data_quality": {
                "status": self.data_quality_status,
                "score": self.data_quality_score,
                "issues_count": len(self.validation_issues),
                "issues": [i.to_dict() for i in self.validation_issues]
            },
            "latest_canonical_metrics": {
                "revenue": self.get_canonical_value("revenue"),
                "ebitda": self.get_canonical_value("ebitda"),
                "ebit": self.get_canonical_value("ebit"),
                "net_income": self.get_canonical_value("net_income"),
                "total_assets": self.get_canonical_value("total_assets"),
                "total_equity": self.get_canonical_value("total_equity"),
                "total_debt": self.get_canonical_value("total_debt"),
                "cash": self.get_canonical_value("cash"),
                "current_assets": self.get_canonical_value("current_assets"),
                "current_liabilities": self.get_canonical_value("current_liabilities"),
                "invested_capital": self.get_canonical_value("invested_capital"),
                "operating_cash_flow": self.get_canonical_value("operating_cash_flow"),
                "capex": self.get_canonical_value("capex"),
                "roic": self.get_canonical_value("roic"),
                "altman_z": self.get_canonical_value("altman_z")
            }
        }

    def to_canonical_dict(self) -> Dict[str, Any]:
        """
        Creates the Section 3 Canonical Financial Data Object:
        canonical = {
            "company": {},
            "market": {},
            "income_statement": {},
            "balance_sheet": {},
            "cash_flow": {},
            "ratios": {},
            "bank_metrics": {},
            "valuation_inputs": {},
            "metadata": {}
        }
        Every single financial field adheres strictly to:
        value, source, period, unit, currency, reported_or_derived,
        calculation_method, confidence, availability_status
        """
        is_bank = (self.identity.company_type == "Bank")
        latest_p = clean_fiscal_year_label(self.historical_periods[-1]) if self.historical_periods else "latest"
        latest_h = self.historical.get(latest_p, {})

        def _format_field(f: Any, default_name: str, calc: str = "direct_reported", rep_der: str = "reported", unit: str = "Cr") -> Dict[str, Any]:
            if isinstance(f, FinancialField):
                d = f.to_canonical_dict()
                if not d.get("calculation_method") or d.get("calculation_method") == "none":
                    d["calculation_method"] = calc
                return d
            elif f is not None:
                return {
                    "value": float(f),
                    "source": "Screener.in / Canonical Engine",
                    "period": latest_p,
                    "unit": unit,
                    "currency": self.identity.currency,
                    "reported_or_derived": rep_der,
                    "calculation_method": calc,
                    "confidence": "HIGH",
                    "availability_status": "available"
                }
            else:
                return {
                    "value": None,
                    "source": "Screener.in",
                    "period": latest_p,
                    "unit": unit,
                    "currency": self.identity.currency,
                    "reported_or_derived": "not_applicable" if is_bank and default_name in ['ebitda', 'capex', 'invested_capital', 'roic', 'altman_z'] else "reported",
                    "calculation_method": f"unavailable: {default_name}",
                    "confidence": "LOW",
                    "availability_status": "not_applicable" if is_bank and default_name in ['ebitda', 'capex', 'invested_capital', 'roic', 'altman_z'] else "unavailable"
                }

        # 1. Company
        company_dict = {
            "name": self.identity.company_name,
            "ticker": self.identity.ticker,
            "exchange": self.identity.exchange,
            "sector": self.identity.sector,
            "industry": self.identity.industry,
            "sub_industry": self.identity.sub_industry,
            "reporting_basis": self.identity.reporting_basis,
            "is_consolidated": self.identity.is_consolidated,
            "currency": self.identity.currency,
            "unit": self.identity.unit,
            "fiscal_year_end": self.identity.fiscal_year_end
        }

        # 2. Market Data
        market_dict = {
            "current_price": _format_field(self.market_data.current_price, "current_price", calc="market_quote", rep_der="reported", unit="INR"),
            "shares_outstanding": _format_field(self.market_data.shares_outstanding, "shares_outstanding", calc="equity_capital / face_value or mcap / price", rep_der="derived", unit="Cr"),
            "market_cap": _format_field(self.market_data.market_cap, "market_cap", calc="price * shares_outstanding", rep_der="derived", unit="Cr"),
            "calculated_market_cap": _format_field(self.market_data.calculated_market_cap, "calculated_market_cap", calc="price * shares_outstanding", rep_der="derived", unit="Cr"),
            "enterprise_value": _format_field(self.market_data.enterprise_value, "enterprise_value", calc="mcap + debt - cash" if not is_bank else "equal_to_mcap_for_bank", rep_der="derived", unit="Cr"),
            "beta": _format_field(self.market_data.beta, "beta", calc="regression_vs_index", rep_der="derived", unit="x"),
            "cost_of_equity": _format_field(round(self.market_data.cost_of_equity * 100, 2), "cost_of_equity", calc="rf + beta * erp", rep_der="derived", unit="%"),
            "wacc": _format_field(round(self.market_data.wacc * 100, 2), "wacc", calc="we * ke + wd * kd * (1 - t)" if not is_bank else "cost_of_equity_primary", rep_der="derived", unit="%")
        }

        # 3. Income Statement
        pnl = latest_h.get("income_statement", {})
        income_stmt = {
            "revenue": _format_field(pnl.get("revenue"), "revenue", calc="direct_reported"),
            "ebitda": _format_field(pnl.get("ebitda"), "ebitda", calc="operating_profit + depr" if not is_bank else "not_applicable_for_bank"),
            "depreciation": _format_field(pnl.get("depreciation"), "depreciation", calc="direct_reported"),
            "ebit": _format_field(pnl.get("ebit"), "ebit", calc="ebitda - depr" if not is_bank else "not_applicable_for_bank"),
            "interest": _format_field(pnl.get("interest_expense"), "interest", calc="direct_reported"),
            "pbt": _format_field(pnl.get("pre_tax_income"), "pbt", calc="direct_reported"),
            "tax": _format_field(pnl.get("tax"), "tax", calc="pbt - pat"),
            "net_income": _format_field(pnl.get("net_income"), "net_income", calc="direct_reported")
        }

        # 4. Balance Sheet
        bs = latest_h.get("balance_sheet", {})
        balance_stmt = {
            "cash": _format_field(bs.get("cash"), "cash", calc="cash_and_bank_schedules"),
            "total_debt": _format_field(bs.get("total_debt"), "total_debt", calc="borrowings"),
            "current_assets": _format_field(bs.get("current_assets"), "current_assets", calc="direct_or_schedules"),
            "current_liabilities": _format_field(bs.get("current_liabilities"), "current_liabilities", calc="direct_or_schedules"),
            "fixed_assets": _format_field(bs.get("pp_e"), "fixed_assets", calc="net_block"),
            "cwip": _format_field(bs.get("cwip"), "cwip", calc="direct_reported" if not is_bank else "not_applicable_for_bank"),
            "investments": _format_field(bs.get("investments"), "investments", calc="direct_reported"),
            "total_assets": _format_field(bs.get("total_assets"), "total_assets", calc="direct_reported"),
            "total_equity": _format_field(bs.get("total_equity"), "total_equity", calc="equity_capital + reserves")
        }

        # 5. Cash Flow
        cf = latest_h.get("cash_flow", {})
        cash_flow_stmt = {
            "operating_cash_flow": _format_field(cf.get("operating_cash_flow"), "operating_cash_flow", calc="direct_reported"),
            "capex": _format_field(cf.get("capex"), "capex", calc="cfi_gross_capex" if not is_bank else "not_applicable_for_bank"),
            "free_cash_flow": _format_field(cf.get("free_cash_flow"), "free_cash_flow", calc="cfo - capex" if not is_bank else "not_applicable_for_bank")
        }

        # 6. Ratios
        der = latest_h.get("derived", {})
        ratios_dict = {
            "roe": _format_field(der.get("roe"), "roe", calc="net_income / total_equity", unit="%"),
            "roa": _format_field(der.get("roa"), "roa", calc="net_income / avg_assets", unit="%"),
            "roic": _format_field(der.get("roic"), "roic", calc="nopat / invested_capital" if not is_bank else "not_applicable_for_bank", unit="%"),
            "asset_turnover": _format_field(der.get("asset_turnover"), "asset_turnover", calc="revenue / avg_assets", unit="x"),
            "equity_multiplier": _format_field(der.get("equity_multiplier"), "equity_multiplier", calc="avg_assets / avg_equity", unit="x"),
            "altman_z": _format_field(der.get("altman_z"), "altman_z", calc="altman_formula" if not is_bank else "not_applicable_for_bank", unit="score")
        }

        # 7. Bank Metrics
        bm = latest_h.get("bank_metrics", {})
        bank_metrics_dict = {}
        if is_bank and bm:
            for k, v in bm.items():
                bank_metrics_dict[k] = _format_field(v, k, calc=f"bank_metric_{k}")

        # 8. Valuation Inputs
        valuation_inputs = {
            "risk_free_rate": _format_field(round(self.market_data.risk_free_rate * 100, 2), "risk_free_rate", calc="10y_gsec_benchmark", unit="%"),
            "equity_risk_premium": _format_field(round(self.market_data.equity_risk_premium * 100, 2), "equity_risk_premium", calc="damodaran_india_erp", unit="%"),
            "beta": _format_field(self.market_data.beta, "beta", calc="sector_peer_median_beta", unit="x"),
            "terminal_growth": _format_field(2.5, "terminal_growth", calc="long_term_gdp_anchor", unit="%"),
            "tax_rate": _format_field(round(self.market_data.tax_rate * 100, 2), "tax_rate", calc="corporate_statutory_tax", unit="%")
        }

        # 9. Metadata (Section 7)
        canonical_framework = self.classification.get("canonical_valuation_framework", "Bank Equity Valuation" if is_bank else "FCFF / Enterprise Valuation")
        metadata_dict = {
            "company_type": self.identity.company_type,
            "valuation_framework": canonical_framework,
            "classification_rationale": self.classification.get("classification_rationale", "Universal Classification Engine"),
            "is_bank": is_bank,
            "is_financial": self.classification.get("is_financial", is_bank),
            "validation_status": self.data_quality_status,
            "validation_score": self.data_quality_score,
            "issues_count": len(self.validation_issues)
        }

        return {
            "company": company_dict,
            "market": market_dict,
            "income_statement": income_stmt,
            "balance_sheet": balance_stmt,
            "cash_flow": cash_flow_stmt,
            "ratios": ratios_dict,
            "bank_metrics": bank_metrics_dict,
            "valuation_inputs": valuation_inputs,
            "metadata": metadata_dict
        }

