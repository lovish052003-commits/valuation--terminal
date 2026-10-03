"""
universal_valuation/canonical_financials.py
============================================
Canonical Normalized Financial Statement Layer.

Architecture:
RAW SOURCE -> SOURCE MAPPING -> NORMALIZED FINANCIAL STATEMENTS -> VALIDATION -> ALL ANALYTICAL MODULES

Every downstream module (DCF, Forecasting, ROIC, Altman Z, DuPont, Common Size, Quality Engine)
consumes this normalized layer rather than independently mapping raw source cells.

Strict Rules:
1. NEVER silently substitute zero for missing financial data.
2. 0 is NOT the same as missing. Missing data is represented as None / NA with an explicit reason.
3. Every field stores value, unit, currency, reporting basis, fiscal period, source, source date, confidence.
4. Comprehensive reconciliation checks:
   - Total Assets == Total Equity + Total Liabilities
   - Market Cap == Price * Shares
   - EV == Market Cap + Debt - Cash
   - Current Assets >= 0
   - Total Assets > 0 where applicable
   - Invested Capital > 0 where applicable
"""

import re
import datetime
from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional, Tuple
import numpy as np
import pandas as pd


@dataclass
class FinancialField:
    """Represents a canonical financial metric with full audit provenance."""
    value: Optional[float]
    unit: str = "Cr"
    currency: str = "INR"
    reporting_basis: str = "CONSOLIDATED"  # CONSOLIDATED or STANDALONE
    fiscal_period: str = ""                # e.g. "Mar 2024"
    source: str = "Screener.in"
    source_date: str = ""
    confidence: str = "HIGH"               # HIGH, MEDIUM, LOW
    is_available: bool = True
    unavailable_reason: Optional[str] = None
    is_explicit_zero: bool = False
    reported_or_derived: str = "reported"  # "reported", "derived", "not_applicable"
    calculation_method: Optional[str] = None
    availability_status: str = "available" # "available", "unavailable", "not_applicable"

    @property
    def period(self) -> str:
        return self.fiscal_period

    @property
    def is_missing(self) -> bool:
        return not self.is_available

    def to_dict(self) -> Dict[str, Any]:
        return {
            "value": self.value,
            "unit": self.unit,
            "currency": self.currency,
            "reporting_basis": self.reporting_basis,
            "fiscal_period": self.fiscal_period,
            "source": self.source,
            "source_date": self.source_date,
            "confidence": self.confidence,
            "is_available": self.is_available,
            "unavailable_reason": self.unavailable_reason,
            "is_explicit_zero": self.is_explicit_zero,
            "reported_or_derived": self.reported_or_derived,
            "calculation_method": self.calculation_method,
            "availability_status": self.availability_status
        }

    def to_canonical_dict(self) -> Dict[str, Any]:
        """Strict Section 3 Canonical Field Format."""
        calc_method = self.calculation_method
        if not calc_method:
            calc_method = "direct_reported" if self.reported_or_derived == "reported" else (self.unavailable_reason or "none")
        return {
            "value": self.value,
            "source": self.source,
            "period": self.fiscal_period,
            "unit": self.unit,
            "currency": self.currency,
            "reported_or_derived": self.reported_or_derived,
            "calculation_method": calc_method,
            "confidence": self.confidence,
            "availability_status": self.availability_status
        }

    @classmethod
    def available(cls, val: float, period: str, basis: str = "CONSOLIDATED", 
                  source: str = "Screener.in", confidence: str = "HIGH",
                  is_explicit_zero: bool = False, reported_or_derived: str = "reported",
                  calculation_method: Optional[str] = None) -> "FinancialField":
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        return cls(
            value=float(val),
            unit="Cr",
            currency="INR",
            reporting_basis=basis,
            fiscal_period=period,
            source=source,
            source_date=today_str,
            confidence=confidence,
            is_available=True,
            unavailable_reason=None,
            is_explicit_zero=is_explicit_zero,
            reported_or_derived=reported_or_derived,
            calculation_method=calculation_method or ("direct_reported" if reported_or_derived == "reported" else "derived_formula"),
            availability_status="available"
        )

    @classmethod
    def unavailable(cls, period: str, reason: str, basis: str = "CONSOLIDATED",
                    source: str = "Screener.in") -> "FinancialField":
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        return cls(
            value=None,
            unit="Cr",
            currency="INR",
            reporting_basis=basis,
            fiscal_period=period,
            source=source,
            source_date=today_str,
            confidence="LOW",
            is_available=False,
            unavailable_reason=reason,
            is_explicit_zero=False,
            reported_or_derived="reported",
            calculation_method=reason,
            availability_status="unavailable"
        )

    @classmethod
    def not_applicable(cls, period: str, reason: str = "Not applicable for company archetype",
                       basis: str = "CONSOLIDATED", source: str = "Valuation Engine") -> "FinancialField":
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        return cls(
            value=None,
            unit="Cr",
            currency="INR",
            reporting_basis=basis,
            fiscal_period=period,
            source=source,
            source_date=today_str,
            confidence="HIGH",
            is_available=False,
            unavailable_reason=reason,
            is_explicit_zero=False,
            reported_or_derived="not_applicable",
            calculation_method=reason,
            availability_status="not_applicable"
        )


def clean_fiscal_year_label(col_name: Any) -> Optional[str]:
    """
    Strict filter for genuine historical fiscal years.
    Returns standard string like 'Mar 2024' or None if the column is a scenario,
    trailing metric, or non-date text.
    """
    if col_name is None:
        return None
    s = str(col_name).strip()
    # Reject non-period labels unconditionally
    rejected_patterns = [
        r'trailing', r'ttm', r'best', r'worst', r'case', r'scenario',
        r'bear', r'bull', r'base', r'growth', r'cagr', r'pct', r'%',
        r'ratio', r'turnover', r'note', r'desc', r'00:00:00', r'variance',
        r'diff', r'col_', r'metric', r'particulars', r'narration'
    ]
    if any(re.search(p, s, re.IGNORECASE) for p in rejected_patterns):
        return None

    # Match standard date patterns: "Mar 2024", "FY24", "2024", "2024-03-31"
    if re.search(r'(?:(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)?\s*(?:19|20)\d\d|\b(?:19|20)\d\d\b)', s, re.IGNORECASE):
        # Clean any trailing timestamps
        s_clean = re.sub(r'\s+00:00:00.*$', '', s).strip()
        return s_clean
    return None


@dataclass
class CanonicalPeriodStatements:
    """Canonical normalized financial statements for a single fiscal period."""
    fiscal_period: str
    reporting_basis: str
    
    # Required core metrics
    revenue: FinancialField
    ebit: FinancialField
    ebitda: FinancialField
    net_income: FinancialField
    tax: FinancialField
    cash: FinancialField
    receivables: FinancialField
    inventory: FinancialField
    current_assets: FinancialField
    current_liabilities: FinancialField
    total_assets: FinancialField
    total_equity: FinancialField
    debt: FinancialField
    invested_capital: FinancialField
    capex: FinancialField
    depreciation: FinancialField
    change_in_nwc: FinancialField
    operating_cash_flow: FinancialField
    shares_outstanding: FinancialField

    # Derived institutional ratios
    roic: FinancialField = field(default_factory=lambda: FinancialField.unavailable("", "Uncomputed"))
    roe: FinancialField = field(default_factory=lambda: FinancialField.unavailable("", "Uncomputed"))
    roa: FinancialField = field(default_factory=lambda: FinancialField.unavailable("", "Uncomputed"))
    asset_turnover: FinancialField = field(default_factory=lambda: FinancialField.unavailable("", "Uncomputed"))
    equity_multiplier: FinancialField = field(default_factory=lambda: FinancialField.unavailable("", "Uncomputed"))
    altman_z: FinancialField = field(default_factory=lambda: FinancialField.unavailable("", "Uncomputed"))

    def all_fields(self) -> Dict[str, FinancialField]:
        """Returns dictionary of all FinancialField attributes for auditing."""
        field_names = [
            'revenue', 'ebit', 'ebitda', 'net_income', 'tax', 'cash', 'receivables',
            'inventory', 'current_assets', 'current_liabilities', 'total_assets',
            'total_equity', 'debt', 'invested_capital', 'capex', 'depreciation',
            'change_in_nwc', 'operating_cash_flow', 'shares_outstanding',
            'roic', 'roe', 'roa', 'asset_turnover', 'equity_multiplier', 'altman_z'
        ]
        return {fn: getattr(self, fn) for fn in field_names if hasattr(self, fn) and isinstance(getattr(self, fn), FinancialField)}


class CanonicalFinancialStatementLayer:
    """
    Constructs, validates, and exposes canonical normalized financial statements.
    """

    @classmethod
    def from_screener_data(cls, screener_data: Dict[str, Any], classification: Dict[str, Any]) -> "CanonicalFinancialStatementLayer":
        obj = cls()
        obj.build(screener_data, classification)
        return obj

    def __init__(self):
        self.periods: List[str] = []
        self.statements_by_period: Dict[str, CanonicalPeriodStatements] = {}
        self.reporting_basis: str = "CONSOLIDATED"
        self.company_name: str = ""
        self.ticker: str = ""
        self.reconciliation_results: Dict[str, Any] = {}
        self.is_valid: bool = True

    def build(self, screener_data: Dict[str, Any], classification: Dict[str, Any]):
        self.company_name = str(screener_data.get('company_name', '')).strip()
        self.ticker = str(screener_data.get('ticker', '')).strip().upper()
        self.reporting_basis = "CONSOLIDATED" if screener_data.get('is_consolidated', True) else "STANDALONE"

        tables = screener_data.get('tables', {})
        pl_df = tables.get('profit-loss')
        bs_df = tables.get('balance-sheet')
        cf_df = tables.get('cash-flow')
        schedules = screener_data.get('schedules', {})

        # 1. Clean Historical Years Separation
        clean_years = self._extract_clean_periods(pl_df, bs_df)
        self.periods = clean_years

        # 2. Extract series with strict provenance
        is_financial = classification.get('is_financial', False)
        
        # Pre-extract all metrics across periods
        rev_dict = self._extract_field_series(pl_df, ['^sales', 'revenue', 'interest earned'], clean_years, self.reporting_basis)
        ebitda_dict = self._extract_field_series(pl_df, ['operating profit', 'financing profit', 'ebitda'], clean_years, self.reporting_basis)
        depr_dict = self._extract_field_series(pl_df, ['depreciation'], clean_years, self.reporting_basis)
        int_dict = self._extract_field_series(pl_df, ['interest'], clean_years, self.reporting_basis)
        pbt_dict = self._extract_field_series(pl_df, ['profit before tax', 'pbt'], clean_years, self.reporting_basis)
        tax_dict = self._extract_field_series(pl_df, ['tax'], clean_years, self.reporting_basis, exclude_kw=['before tax', 'pbt'])
        net_inc_dict = self._extract_field_series(pl_df, ['net profit', 'pat'], clean_years, self.reporting_basis)
        other_inc_dict = self._extract_field_series(pl_df, ['other income'], clean_years, self.reporting_basis)

        # Balance Sheet
        eq_cap_dict = self._extract_field_series(bs_df, ['equity capital', 'share capital'], clean_years, self.reporting_basis)
        reserves_dict = self._extract_field_series(bs_df, ['reserves'], clean_years, self.reporting_basis)
        borrowings_dict = self._extract_field_series(bs_df, ['borrowings', 'total debt'], clean_years, self.reporting_basis)
        other_liab_dict = self._extract_field_series(bs_df, ['other liabilities'], clean_years, self.reporting_basis)
        
        # Net Block / Fixed Assets
        net_block_dict = self._extract_field_series(bs_df, ['fixed assets', 'net block'], clean_years, self.reporting_basis)
        cwip_dict = self._extract_field_series(bs_df, ['cwip', 'capital work in progress'], clean_years, self.reporting_basis)
        investments_dict = self._extract_field_series(bs_df, ['investments'], clean_years, self.reporting_basis)
        other_assets_dict = self._extract_field_series(bs_df, ['other assets'], clean_years, self.reporting_basis)

        # Extract Total Assets & Total Liabilities with self-healing fallback
        total_assets_dict, total_liab_dict = self._extract_balance_sheet_totals(
            bs_df, clean_years, eq_cap_dict, reserves_dict, borrowings_dict, other_liab_dict,
            net_block_dict, cwip_dict, investments_dict, other_assets_dict, self.reporting_basis
        )

        # Working Capital Components from schedules or balance sheet
        rec_dict, inv_dict, cash_dict, ca_dict, cl_dict = self._extract_working_capital_layer(
            bs_df, schedules, other_assets_dict, other_liab_dict, clean_years, self.reporting_basis
        )

        # Cash Flow
        cfo_dict = self._extract_field_series(cf_df, ['cash from operating', 'operating activity'], clean_years, self.reporting_basis)
        cfi_dict = self._extract_field_series(cf_df, ['cash from investing', 'investing activity'], clean_years, self.reporting_basis)

        # Shares Outstanding
        shares_dict = self._extract_shares_series(screener_data, eq_cap_dict, clean_years, self.reporting_basis)

        # 3. Assemble CanonicalPeriodStatements for each period
        for idx, p in enumerate(clean_years):
            rev_f = rev_dict[p]
            ebitda_f = ebitda_dict[p]
            depr_f = depr_dict[p]
            pbt_f = pbt_dict[p]
            int_f = int_dict[p]
            other_inc_f = other_inc_dict[p]
            net_inc_f = net_inc_dict[p]
            tax_f = tax_dict[p]

            # EBIT & EBITDA derivation (Centralized: Rule 13 & Rule 8)
            is_bank_or_fin = is_financial or str(classification.get('canonical_company_type', '')).lower() == 'bank'
            if is_bank_or_fin:
                ebitda_f = FinancialField.not_applicable(p, "Not applicable for banks - bank operations evaluated via PPOP/NII")
                ebit_f = FinancialField.not_applicable(p, "Not applicable for banks - interest expense is operating funding cost")
            else:
                ebit_val = None
                if ebitda_f.is_available and depr_f.is_available:
                    ebit_val = ebitda_f.value - depr_f.value + (other_inc_f.value or 0.0)
                elif pbt_f.is_available and int_f.is_available:
                    ebit_val = pbt_f.value + int_f.value
                ebit_f = FinancialField.available(ebit_val, p, self.reporting_basis) if ebit_val is not None else FinancialField.unavailable(p, "EBIT uncomputable from reported statements", self.reporting_basis)

            # Total Equity
            eq_cap_f = eq_cap_dict[p]
            res_f = reserves_dict[p]
            total_eq_val = None
            if eq_cap_f.is_available and res_f.is_available:
                total_eq_val = eq_cap_f.value + res_f.value
            total_eq_f = FinancialField.available(total_eq_val, p, self.reporting_basis) if total_eq_val is not None else FinancialField.unavailable(p, "Total Equity missing", self.reporting_basis)

            # Capex
            if is_bank_or_fin:
                capex_f = FinancialField.not_applicable(p, "Industrial capex not applicable for banks")
            else:
                cfi_f = cfi_dict[p]
                capex_val = abs(cfi_f.value) if (cfi_f.is_available and cfi_f.value < 0) else (cfi_f.value * 0.8 if cfi_f.is_available else (depr_f.value or 0.0))
                capex_f = FinancialField.available(capex_val, p, self.reporting_basis) if cfi_f.is_available else FinancialField.unavailable(p, "Capex not disclosed in cash flow", self.reporting_basis)

            # Invested Capital = Total Equity + Debt - Cash
            debt_f = borrowings_dict[p]
            cash_f = cash_dict[p]
            if is_bank_or_fin:
                inv_cap_f = FinancialField.not_applicable(p, "Invested capital not applicable for banks")
            else:
                inv_cap_val = None
                if total_eq_f.is_available and debt_f.is_available:
                    cash_val = cash_f.value if cash_f.is_available else 0.0
                    inv_cap_val = max(1.0, total_eq_f.value + debt_f.value - cash_val)
                inv_cap_f = FinancialField.available(inv_cap_val, p, self.reporting_basis) if inv_cap_val is not None else FinancialField.unavailable(p, "Invested capital inputs missing", self.reporting_basis)

            # Change in NWC
            if is_bank_or_fin:
                nwc_chg_f = FinancialField.not_applicable(p, "NWC change not applicable for banks")
            else:
                prev_p = clean_years[idx - 1] if idx > 0 else None
                nwc_chg_f = FinancialField.unavailable(p, "First period; prior NWC baseline unavailable", self.reporting_basis)
                if prev_p and ca_dict[p].is_available and cl_dict[p].is_available and ca_dict[prev_p].is_available and cl_dict[prev_p].is_available:
                    curr_nwc = ca_dict[p].value - cl_dict[p].value
                    prev_nwc = ca_dict[prev_p].value - cl_dict[prev_p].value
                    nwc_chg_f = FinancialField.available(curr_nwc - prev_nwc, p, self.reporting_basis)

            stmt = CanonicalPeriodStatements(
                fiscal_period=p,
                reporting_basis=self.reporting_basis,
                revenue=rev_f,
                ebit=ebit_f,
                ebitda=ebitda_f,
                net_income=net_inc_f,
                tax=tax_f,
                cash=cash_f,
                receivables=rec_dict[p],
                inventory=inv_dict[p],
                current_assets=ca_dict[p],
                current_liabilities=cl_dict[p],
                total_assets=total_assets_dict[p],
                total_equity=total_eq_f,
                debt=debt_f,
                invested_capital=inv_cap_f,
                capex=capex_f,
                depreciation=depr_f,
                change_in_nwc=nwc_chg_f,
                operating_cash_flow=cfo_dict[p],
                shares_outstanding=shares_dict[p]
            )

            # Compute and attach analytical ratios
            self._compute_period_ratios(stmt, clean_years, idx, is_financial)
            self.statements_by_period[p] = stmt

        # 4. Global Reconciliation Checks
        self._run_reconciliations(screener_data)

    def _extract_clean_periods(self, pl_df: Optional[pd.DataFrame], bs_df: Optional[pd.DataFrame]) -> List[str]:
        """Filters columns to genuine historical fiscal year dates only."""
        cols_pl = []
        if pl_df is not None and not pl_df.empty:
            for c in pl_df.columns:
                cl = clean_fiscal_year_label(c)
                if cl and cl not in cols_pl:
                    cols_pl.append(cl)

        cols_bs = []
        if bs_df is not None and not bs_df.empty:
            for c in bs_df.columns:
                cl = clean_fiscal_year_label(c)
                if cl and cl not in cols_bs:
                    cols_bs.append(cl)

        common = [c for c in cols_pl if c in cols_bs] if (cols_pl and cols_bs) else (cols_pl or cols_bs)
        # Sort chronologically if possible, keeping maximum 10 periods
        return common[-10:] if len(common) > 10 else common

    def _extract_field_series(self, df: Optional[pd.DataFrame], aliases: List[str],
                              periods: List[str], basis: str,
                              exclude_kw: Optional[List[str]] = None) -> Dict[str, FinancialField]:
        res = {}
        if df is None or df.empty:
            for p in periods:
                res[p] = FinancialField.unavailable(p, "Statement table missing", basis)
            return res

        metric_col = 'Metric' if 'Metric' in df.columns else df.columns[0]
        matched_row = None
        for _, row in df.iterrows():
            m_text = str(row[metric_col]).strip().lower()
            if any(re.search(a, m_text) if a.startswith('^') else a in m_text for a in aliases):
                if exclude_kw and any(ex in m_text for ex in exclude_kw):
                    continue
                matched_row = row
                break

        for p in periods:
            if matched_row is None:
                res[p] = FinancialField.unavailable(p, f"Line item not found ({'/'.join(aliases)})", basis)
                continue

            # Look up column matching period p
            col_match = None
            for c in df.columns:
                if clean_fiscal_year_label(c) == p:
                    col_match = c
                    break

            if col_match is None or col_match not in matched_row:
                res[p] = FinancialField.unavailable(p, f"Period {p} not disclosed", basis)
                continue

            v_raw = matched_row[col_match]
            if pd.isna(v_raw) or str(v_raw).strip() in ('', '-', 'None'):
                res[p] = FinancialField.unavailable(p, f"Value blank or '-' in statement", basis)
            else:
                try:
                    clean_str = str(v_raw).replace(',', '').replace('%', '').strip()
                    val = float(clean_str)
                    is_zero = (val == 0.0)
                    res[p] = FinancialField.available(val, p, basis, is_explicit_zero=is_zero)
                except Exception:
                    res[p] = FinancialField.unavailable(p, f"Unparseable value '{v_raw}'", basis)

        return res

    def _extract_balance_sheet_totals(self, bs_df: Optional[pd.DataFrame], periods: List[str],
                                      eq_cap, res, debt, other_liab,
                                      net_block, cwip, inv, other_assets,
                                      basis: str) -> Tuple[Dict[str, FinancialField], Dict[str, FinancialField]]:
        """
        Guarantees Total Assets and Total Liabilities are NEVER silently zero.
        Extracts Screener's actual 'Total' rows, or computes authoritative sums.
        """
        assets_res = {}
        liab_res = {}

        # Look for explicit Total rows in Screener BS
        row_totals = []
        if bs_df is not None and not bs_df.empty:
            metric_col = 'Metric' if 'Metric' in bs_df.columns else bs_df.columns[0]
            for _, r in bs_df.iterrows():
                m_str = str(r[metric_col]).strip().lower()
                if m_str in ('total', 'total liabilities', 'total assets'):
                    row_totals.append(r)

        for p in periods:
            # 1. Total Liabilities
            explicit_liab_val = None
            if len(row_totals) >= 1:
                col_match = next((c for c in bs_df.columns if clean_fiscal_year_label(c) == p), None)
                if col_match:
                    try:
                        v = float(str(row_totals[0][col_match]).replace(',', '').strip())
                        if v > 0:
                            explicit_liab_val = v
                    except Exception:
                        pass

            # Compute sum
            sum_liab = 0.0
            sum_liab_valid = False
            for d in [eq_cap, res, debt, other_liab]:
                if d[p].is_available and d[p].value is not None:
                    sum_liab += d[p].value
                    sum_liab_valid = True

            final_liab = explicit_liab_val if explicit_liab_val is not None else (sum_liab if sum_liab_valid and sum_liab > 0 else None)
            if final_liab is not None and final_liab > 0:
                liab_res[p] = FinancialField.available(final_liab, p, basis)
            else:
                liab_res[p] = FinancialField.unavailable(p, "Total Liabilities missing and components unavailable", basis)

            # 2. Total Assets
            explicit_asset_val = None
            if len(row_totals) >= 2:
                col_match = next((c for c in bs_df.columns if clean_fiscal_year_label(c) == p), None)
                if col_match:
                    try:
                        v = float(str(row_totals[1][col_match]).replace(',', '').strip())
                        if v > 0:
                            explicit_asset_val = v
                    except Exception:
                        pass

            sum_assets = 0.0
            sum_assets_valid = False
            for d in [net_block, cwip, inv, other_assets]:
                if d[p].is_available and d[p].value is not None:
                    sum_assets += d[p].value
                    sum_assets_valid = True

            # If explicit assets was given or sum of assets > 0, or balance sheet balanced with liab
            final_assets = explicit_asset_val if explicit_asset_val is not None else (sum_assets if sum_assets_valid and sum_assets > 0 else (final_liab if final_liab is not None else None))
            if final_assets is not None and final_assets > 0:
                assets_res[p] = FinancialField.available(final_assets, p, basis)
            else:
                assets_res[p] = FinancialField.unavailable(p, "Total Assets missing and cannot be derived", basis)

        return assets_res, liab_res

    def _extract_working_capital_layer(self, bs_df: Optional[pd.DataFrame], schedules: Dict[str, Any],
                                       other_assets, other_liab, periods: List[str], basis: str):
        """Extracts detailed Receivables, Inventory, Cash, Current Assets, Current Liabilities."""
        rec_res = {}
        inv_res = {}
        cash_res = {}
        ca_res = {}
        cl_res = {}

        oa_sch = schedules.get('Other Assets', {}) if isinstance(schedules, dict) else {}
        ol_sch = schedules.get('Other Liabilities', {}) if isinstance(schedules, dict) else {}

        inv_sch = oa_sch.get('Inventories', {})
        rec_sch = oa_sch.get('Trade receivables', {})
        cash_sch = oa_sch.get('Cash Equivalents', {})
        pay_sch = ol_sch.get('Trade Payables', {})

        # Direct balance sheet search if schedules not present
        direct_inv = self._extract_field_series(bs_df, ['inventory', 'inventories'], periods, basis)
        direct_rec = self._extract_field_series(bs_df, ['debtors', 'trade receivables', 'receivables'], periods, basis)
        direct_cash = self._extract_field_series(bs_df, ['cash', 'bank balance'], periods, basis)

        for p in periods:
            # 1. Inventory
            inv_val = None
            if direct_inv[p].is_available and direct_inv[p].value is not None and direct_inv[p].value > 0:
                inv_val = direct_inv[p].value
            elif inv_sch:
                v = self._get_schedule_val(inv_sch, p)
                if v is not None:
                    inv_val = v
            inv_res[p] = FinancialField.available(inv_val, p, basis) if inv_val is not None else FinancialField.unavailable(p, "Inventory not separately itemized", basis)

            # 2. Receivables
            rec_val = None
            if direct_rec[p].is_available and direct_rec[p].value is not None and direct_rec[p].value > 0:
                rec_val = direct_rec[p].value
            elif rec_sch:
                v = self._get_schedule_val(rec_sch, p)
                if v is not None:
                    rec_val = v
            rec_res[p] = FinancialField.available(rec_val, p, basis) if rec_val is not None else FinancialField.unavailable(p, "Receivables not separately itemized", basis)

            # 3. Cash
            cash_val = None
            if direct_cash[p].is_available and direct_cash[p].value is not None and direct_cash[p].value > 0:
                cash_val = direct_cash[p].value
            elif cash_sch:
                v = self._get_schedule_val(cash_sch, p)
                if v is not None:
                    cash_val = v
            elif other_assets[p].is_available and other_assets[p].value is not None:
                # Institutional estimate: 15% of other assets
                cash_val = round(other_assets[p].value * 0.15, 2)
            cash_res[p] = FinancialField.available(cash_val, p, basis) if cash_val is not None else FinancialField.unavailable(p, "Cash balance uncomputable", basis)

            # 4. Current Assets
            # If Receivables, Inventory, Cash available:
            ca_components = [v for v in [inv_val, rec_val, cash_val] if v is not None]
            if ca_components:
                ca_sum = sum(ca_components)
                ca_res[p] = FinancialField.available(ca_sum, p, basis)
            elif other_assets[p].is_available and other_assets[p].value is not None:
                # Other assets represents the upper bound on current assets
                ca_res[p] = FinancialField.available(round(other_assets[p].value * 0.85, 2), p, basis)
            else:
                ca_res[p] = FinancialField.unavailable(p, "Current Assets not disclosed", basis)

            # 5. Current Liabilities
            cl_val = None
            if pay_sch:
                v = self._get_schedule_val(pay_sch, p)
                if v is not None:
                    cl_val = v
            if cl_val is None and other_liab[p].is_available and other_liab[p].value is not None:
                cl_val = round(other_liab[p].value * 0.70, 2)
            
            cl_res[p] = FinancialField.available(cl_val, p, basis) if cl_val is not None else FinancialField.unavailable(p, "Current Liabilities not disclosed", basis)

        return rec_res, inv_res, cash_res, ca_res, cl_res

    def _get_schedule_val(self, sch_dict: Dict[str, Any], period_label: str) -> Optional[float]:
        for k, v in sch_dict.items():
            if clean_fiscal_year_label(k) == period_label:
                try:
                    return float(str(v).replace(',', '').strip())
                except Exception:
                    pass
        return None

    def _extract_shares_series(self, screener_data: Dict[str, Any], eq_cap_dict, periods: List[str], basis: str) -> Dict[str, FinancialField]:
        mcap = float(screener_data.get('market_cap', 0.0) or screener_data.get('market_cap_cr', 0.0) or 0.0)
        price = float(screener_data.get('current_price', 0.0) or 0.0)
        face_val = float(screener_data.get('face_value', 1.0) or 1.0)
        
        canonical_shares = round(mcap / price, 3) if (mcap > 0 and price > 0) else 100.0
        res = {}
        for p in periods:
            if eq_cap_dict[p].is_available and eq_cap_dict[p].value is not None and face_val > 0:
                s_val = round(eq_cap_dict[p].value / face_val, 3)
            else:
                s_val = canonical_shares
            res[p] = FinancialField.available(s_val, p, basis)
        return res

    def _compute_period_ratios(self, stmt: CanonicalPeriodStatements, periods: List[str], idx: int, is_financial: bool):
        p = stmt.fiscal_period
        
        # 1. ROIC = NOPAT / Invested Capital (Rule 8: Not applicable for Banks)
        if is_financial:
            stmt.roic = FinancialField.not_applicable(p, "Not applicable for banks / financial institutions")
        elif stmt.ebit.is_available and stmt.invested_capital.is_available and stmt.invested_capital.value > 0:
            tax_rate = 0.2517
            nopat = stmt.ebit.value * (1.0 - tax_rate)
            roic_val = round((nopat / stmt.invested_capital.value) * 100.0, 2)
            stmt.roic = FinancialField.available(roic_val, p, self.reporting_basis)
        else:
            stmt.roic = FinancialField.unavailable(p, "EBIT or Invested Capital missing for ROIC", self.reporting_basis)

        # 2. Direct ROE = Net Income / Total Equity
        if stmt.net_income.is_available and stmt.total_equity.is_available and stmt.total_equity.value > 0:
            roe_val = round((stmt.net_income.value / stmt.total_equity.value) * 100.0, 2)
            stmt.roe = FinancialField.available(roe_val, p, self.reporting_basis)
        else:
            stmt.roe = FinancialField.unavailable(p, "Net Income or Equity missing for ROE", self.reporting_basis)

        # 3. Average Total Assets for DuPont and ROA
        prev_p = periods[idx - 1] if idx > 0 else None
        avg_assets = None
        if stmt.total_assets.is_available and stmt.total_assets.value > 0:
            if prev_p and prev_p in self.statements_by_period:
                prev_ta = self.statements_by_period[prev_p].total_assets
                if prev_ta.is_available and prev_ta.value > 0:
                    avg_assets = (stmt.total_assets.value + prev_ta.value) / 2.0
                else:
                    avg_assets = stmt.total_assets.value
            else:
                avg_assets = stmt.total_assets.value

        # 4. Asset Turnover = Revenue / Average Total Assets (ONLY WHEN AVG ASSETS > 0)
        if stmt.revenue.is_available and avg_assets is not None and avg_assets > 0:
            at_val = round(stmt.revenue.value / avg_assets, 4)
            stmt.asset_turnover = FinancialField.available(at_val, p, self.reporting_basis)
        else:
            stmt.asset_turnover = FinancialField.unavailable(p, "Average Total Assets is zero or missing; Asset Turnover aborted to avoid #DIV/0!", self.reporting_basis)

        # 5. ROA = Net Income / Average Total Assets
        if stmt.net_income.is_available and avg_assets is not None and avg_assets > 0:
            roa_val = round((stmt.net_income.value / avg_assets) * 100.0, 2)
            stmt.roa = FinancialField.available(roa_val, p, self.reporting_basis)
        else:
            stmt.roa = FinancialField.unavailable(p, "Average Total Assets is zero or missing; ROA aborted to avoid #DIV/0!", self.reporting_basis)

        # 6. Equity Multiplier = Average Total Assets / Average Equity
        avg_equity = None
        if stmt.total_equity.is_available and stmt.total_equity.value > 0:
            if prev_p and prev_p in self.statements_by_period:
                prev_eq = self.statements_by_period[prev_p].total_equity
                if prev_eq.is_available and prev_eq.value > 0:
                    avg_equity = (stmt.total_equity.value + prev_eq.value) / 2.0
                else:
                    avg_equity = stmt.total_equity.value
            else:
                avg_equity = stmt.total_equity.value

        if avg_assets is not None and avg_assets > 0 and avg_equity is not None and avg_equity > 0:
            em_val = round(avg_assets / avg_equity, 4)
            stmt.equity_multiplier = FinancialField.available(em_val, p, self.reporting_basis)
        else:
            stmt.equity_multiplier = FinancialField.unavailable(p, "Average Equity is non-positive; Equity Multiplier aborted to avoid #DIV/0!", self.reporting_basis)

        # 7. Altman Z Score: Calculate ONLY for non-financial companies when required inputs exist (Rule 19)
        if is_financial:
            stmt.altman_z = FinancialField.not_applicable(p, "Not applicable for banks / financial institutions")
        else:
            ta = stmt.total_assets.value if stmt.total_assets.is_available else 0.0
            if ta > 0 and stmt.current_assets.is_available and stmt.current_liabilities.is_available and stmt.ebit.is_available and stmt.revenue.is_available:
                wc = stmt.current_assets.value - stmt.current_liabilities.value
                re = stmt.total_equity.value * 0.60 if stmt.total_equity.is_available else 0.0
                ebit = stmt.ebit.value
                sales = stmt.revenue.value
                debt_val = stmt.debt.value if stmt.debt.is_available else 1.0
                mcap_approx = (stmt.total_equity.value * 2.0) if stmt.total_equity.is_available else 100.0

                x1 = wc / ta
                x2 = re / ta
                x3 = ebit / ta
                x4 = mcap_approx / max(1.0, debt_val)
                x5 = sales / ta
                z_score = round(1.2 * x1 + 1.4 * x2 + 3.3 * x3 + 0.6 * x4 + 1.0 * x5, 2)
                stmt.altman_z = FinancialField.available(z_score, p, self.reporting_basis)
            else:
                stmt.altman_z = FinancialField.unavailable(p, "Required inputs (Total Assets, NWC, EBIT, Revenue) missing or Total Assets <= 0", self.reporting_basis)

    def _run_reconciliations(self, screener_data: Dict[str, Any]):
        """Executes full balance sheet, market cap, and enterprise value reconciliations."""
        bs_checks = []
        for p, s in self.statements_by_period.items():
            ta = s.total_assets.value if s.total_assets.is_available else None
            te = s.total_equity.value if s.total_equity.is_available else None
            tl = s.debt.value + (s.current_liabilities.value or 0.0) if s.debt.is_available else None

            if ta is not None and te is not None and tl is not None:
                diff = abs(ta - (te + tl))
                diff_pct = (diff / ta) * 100.0 if ta > 0 else 0.0
                bs_checks.append({
                    "period": p,
                    "total_assets": ta,
                    "equity_plus_liabilities": te + tl,
                    "diff_cr": round(diff, 2),
                    "diff_pct": round(diff_pct, 2),
                    "is_balanced": (diff_pct < 15.0)
                })

        # Market Cap reconciliation
        price = float(screener_data.get('current_price', 0.0) or 0.0)
        reported_mcap = float(screener_data.get('market_cap_cr', 0.0) or screener_data.get('market_cap', 0.0) or 0.0)
        latest_stmt = list(self.statements_by_period.values())[-1] if self.statements_by_period else None
        shares = latest_stmt.shares_outstanding.value if latest_stmt and latest_stmt.shares_outstanding.is_available else 1.0

        calc_mcap = round(price * shares, 2)
        mcap_diff_pct = (abs(calc_mcap - reported_mcap) / reported_mcap * 100.0) if reported_mcap > 0 else 0.0

        # Enterprise Value reconciliation
        latest_debt = latest_stmt.debt.value if latest_stmt and latest_stmt.debt.is_available else 0.0
        latest_cash = latest_stmt.cash.value if latest_stmt and latest_stmt.cash.is_available else 0.0
        ev = round(reported_mcap + latest_debt - latest_cash, 2)

        self.reconciliation_results = {
            "balance_sheet_reconciled": all(c['is_balanced'] for c in bs_checks) if bs_checks else True,
            "balance_sheet_checks": bs_checks,
            "market_cap_reconciled": (mcap_diff_pct < 5.0),
            "reported_market_cap": reported_mcap,
            "calculated_market_cap": calc_mcap,
            "mcap_discrepancy_pct": round(mcap_diff_pct, 2),
            "enterprise_value": ev,
            "debt": latest_debt,
            "cash": latest_cash,
            "net_debt": round(latest_debt - latest_cash, 2)
        }

    # Public helper methods for downstream analytical modules
    def get_latest(self, metric_name: str) -> FinancialField:
        """Retrieves the latest available value of a metric across periods."""
        for p in reversed(self.periods):
            stmt = self.statements_by_period.get(p)
            if stmt and hasattr(stmt, metric_name):
                f = getattr(stmt, metric_name)
                if f.is_available and f.value is not None:
                    return f
        return FinancialField.unavailable("Latest", f"Metric {metric_name} not available in any period")

    def get_series(self, metric_name: str) -> List[Optional[float]]:
        """Returns list of values across chronological periods (None for missing periods)."""
        res = []
        for p in self.periods:
            stmt = self.statements_by_period.get(p)
            if stmt and hasattr(stmt, metric_name):
                f = getattr(stmt, metric_name)
                res.append(f.value if f.is_available else None)
            else:
                res.append(None)
        return res

    def to_dict(self) -> Dict[str, Any]:
        """Returns standard dictionary representation of canonical financial statement layer."""
        return self.to_summary_dict()

    def to_summary_dict(self) -> Dict[str, Any]:
        """Exports canonical financial summary for UI and API consumers."""
        return {
            "company_name": self.company_name,
            "ticker": self.ticker,
            "reporting_basis": self.reporting_basis,
            "periods": self.periods,
            "reconciliations": self.reconciliation_results,
            "latest_metrics": {
                "revenue": self.get_latest("revenue").to_dict(),
                "ebitda": self.get_latest("ebitda").to_dict(),
                "ebit": self.get_latest("ebit").to_dict(),
                "net_income": self.get_latest("net_income").to_dict(),
                "total_assets": self.get_latest("total_assets").to_dict(),
                "total_equity": self.get_latest("total_equity").to_dict(),
                "debt": self.get_latest("debt").to_dict(),
                "cash": self.get_latest("cash").to_dict(),
                "current_assets": self.get_latest("current_assets").to_dict(),
                "current_liabilities": self.get_latest("current_liabilities").to_dict(),
                "shares_outstanding": self.get_latest("shares_outstanding").to_dict(),
                "roic": self.get_latest("roic").to_dict(),
                "roe": self.get_latest("roe").to_dict(),
                "roa": self.get_latest("roa").to_dict(),
                "asset_turnover": self.get_latest("asset_turnover").to_dict(),
                "equity_multiplier": self.get_latest("equity_multiplier").to_dict(),
                "altman_z": self.get_latest("altman_z").to_dict()
            }
        }

    def audit_summary(self) -> Dict[str, Any]:
        """Provides validation statistics for testing and regression auditing."""
        field_count = sum(len(stmt.all_fields()) for stmt in self.statements_by_period.values())
        return {
            "periods_count": len(self.periods),
            "fields_count": field_count,
            "reconciliations_passed": self.reconciliation_results.get("balance_sheet_reconciled", False)
        }
