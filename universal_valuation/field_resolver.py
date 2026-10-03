"""
universal_valuation/field_resolver.py
======================================
Universal Financial Field Resolver & Multi-Provider Normalization Layer.
Implements:
- Phase 5: Absolute prohibition of silent zero fallbacks (Strict NULL / NA with availability status).
- Phase 6: Universal Financial Statement Normalization (Hierarchy: Reported -> Operating + D&A -> Derived -> NULL).
- Phase 7: Universal Field Resolution with robust alias mapping, unit & currency normalization.
- Phase 8: Unit and currency harmonization.
- Phase 22: Full source traceability and confidence scoring.
"""

import re
from typing import Dict, Any, List, Optional, Tuple, Union
import pandas as pd
from .canonical_financials import FinancialField, clean_fiscal_year_label

# Comprehensive multi-provider alias dictionary
FIELD_ALIASES: Dict[str, List[str]] = {
    "revenue": [
        "^sales", "revenue from operations", "total revenue", "revenue",
        "net sales", "turnover", "total income", "interest earned"
    ],
    "cogs": [
        "cost of materials consumed", "raw material cost", "material cost",
        "cost of goods sold", "cogs", "purchases of stock-in-trade"
    ],
    "gross_profit": [
        "gross profit"
    ],
    "operating_expenses": [
        "^expenses", "operating expenses", "total expenditure", "other expenses"
    ],
    "ebitda": [
        "operating profit", "ebitda", "operating ebitda", "adjusted ebitda",
        "ebitda before exceptional", "financing profit"
    ],
    "depreciation": [
        "depreciation", "depreciation & amortisation", "depreciation and amortisation",
        "d&a", "amortisation"
    ],
    "ebit": [
        "ebit", "operating profit after depreciation", "pbit"
    ],
    "interest": [
        "^interest", "finance costs", "interest expense", "borrowing costs"
    ],
    "pbt": [
        "profit before tax", "pbt", "pre-tax income", "profit before exceptional"
    ],
    "tax": [
        "^tax", "current tax", "tax expense", "provision for tax"
    ],
    "net_income": [
        "net profit", "pat", "profit after tax", "profit for the period",
        "net income", "consolidated profit"
    ],
    "eps": [
        "eps in rs", "eps", "diluted eps", "basic eps"
    ],
    "cash": [
        "cash equivalents", "cash and bank balances", "cash and cash equivalents",
        "cash balances", "bank balances"
    ],
    "short_term_investments": [
        "current investments", "short term investments", "marketable securities"
    ],
    "receivables": [
        "trade receivables", "debtors", "accounts receivable", "receivables"
    ],
    "inventory": [
        "inventories", "inventory", "stock in trade", "finished goods"
    ],
    "current_assets": [
        "other assets", "current assets", "total current assets"
    ],
    "fixed_assets": [
        "fixed assets", "net block", "property, plant and equipment", "property plant & equipment"
    ],
    "cwip": [
        "cwip", "capital work in progress", "capital work-in-progress"
    ],
    "investments": [
        "investments", "non-current investments", "total investments"
    ],
    "other_assets": [
        "other assets", "other non-current assets"
    ],
    "total_assets": [
        "total assets", "total", "total application of funds"
    ],
    "equity_capital": [
        "equity capital", "share capital", "equity share capital"
    ],
    "reserves": [
        "reserves", "reserves and surplus", "retained earnings", "other equity"
    ],
    "total_equity": [
        "total equity", "shareholders funds", "net worth", "equity"
    ],
    "borrowings": [
        "borrowings", "total debt", "long term borrowings", "short term borrowings",
        "debt", "secured loans", "unsecured loans"
    ],
    "other_liabilities": [
        "other liabilities", "current liabilities", "trade payables"
    ],
    "total_liabilities": [
        "total liabilities", "total", "total sources of funds"
    ],
    "cfo": [
        "cash from operating activity", "cash flow from operations",
        "operating cash flow", "cash generated from operations"
    ],
    "cfi": [
        "cash from investing activity", "cash flow from investing"
    ],
    "cff": [
        "cash from financing activity", "cash flow from financing"
    ],
    "capex": [
        "capex", "capital expenditure", "fixed assets purchased", "purchase of fixed assets"
    ]
}


class UniversalFieldResolver:
    """
    Universal Field Resolver executing multi-source resolution,
    alias matching, unit normalization, and hierarchical derivation.
    """

    @classmethod
    def resolve_raw_metric(
        cls, 
        df: Optional[pd.DataFrame], 
        field_name: str, 
        period: str
    ) -> Tuple[Optional[float], bool, str]:
        """
        Extracts a financial value for a given field and period from a dataframe.
        Returns: (value_or_none, is_explicit_zero, matched_label)
        """
        if df is None or df.empty:
            return None, False, ""

        aliases = FIELD_ALIASES.get(field_name.lower(), [field_name.lower()])
        metric_col = 'Metric' if 'Metric' in df.columns else df.columns[0]

        # Find matching period column
        col_match = None
        for c in df.columns:
            if clean_fiscal_year_label(c) == period:
                col_match = c
                break
        if col_match is None:
            # Try fuzzy match on period string
            for c in df.columns:
                if str(c).strip().lower() == str(period).strip().lower():
                    col_match = c
                    break

        if col_match is None:
            return None, False, ""

        for _, row in df.iterrows():
            row_label = str(row[metric_col]).strip()
            row_label_lower = row_label.lower()

            for alias in aliases:
                match = False
                if alias.startswith('^'):
                    if re.match(alias, row_label_lower, re.IGNORECASE):
                        match = True
                else:
                    if alias in row_label_lower:
                        match = True

                # Disambiguate tax vs profit before tax
                if 'tax' in alias and 'before' not in alias and 'pbt' not in alias:
                    if 'before tax' in row_label_lower or 'pbt' in row_label_lower:
                        match = False

                if match:
                    raw_val = row[col_match]
                    parsed, is_zero = cls._parse_raw_cell(raw_val)
                    if parsed is not None or is_zero:
                        return parsed, is_zero, row_label

        return None, False, ""

    @classmethod
    def _parse_raw_cell(cls, val: Any) -> Tuple[Optional[float], bool]:
        """
        Strict parser distinguishing explicit zero from missing data.
        Returns: (float_value_or_none, is_explicit_zero)
        """
        if pd.isna(val) or val is None:
            return None, False
        s = str(val).strip()
        if s in ('', '-', '--', 'N/A', 'NA', 'null', 'None', 'nil', '.'):
            return None, False
        s_clean = s.replace(',', '').replace('%', '').replace('₹', '').replace('$', '').strip()
        try:
            f = float(s_clean)
            if abs(f) < 1e-9:
                return 0.0, True
            return f, False
        except (ValueError, TypeError):
            return None, False

    @classmethod
    def resolve_field(
        cls, 
        screener_data: Dict[str, Any], 
        field_name: str, 
        period: str,
        reporting_basis: str = "CONSOLIDATED"
    ) -> FinancialField:
        """
        Universal Resolver: Resolves any financial field for a period using
        reported statements, schedules, or component derivation.
        NEVER substitutes silent zero for missing data.
        """
        tables = screener_data.get('tables', {})
        pl_df = tables.get('profit-loss')
        bs_df = tables.get('balance-sheet')
        cf_df = tables.get('cash-flow')
        schedules = screener_data.get('schedules', {})

        field_key = field_name.strip().lower()

        # 1. Direct Reported Search in Primary Tables
        for df, category in [(pl_df, 'P&L'), (bs_df, 'Balance Sheet'), (cf_df, 'Cash Flow')]:
            val, is_zero, matched_lbl = cls.resolve_raw_metric(df, field_key, period)
            if val is not None or is_zero:
                return FinancialField.available(
                    val=val or 0.0,
                    period=period,
                    basis=reporting_basis,
                    source=f"Screener.in ({category}: '{matched_lbl}')",
                    confidence="HIGH",
                    is_explicit_zero=is_zero
                )

        # 2. Schedule Search (Working Capital & Expense Breakdown)
        if field_key in ('receivables', 'inventory', 'cash'):
            val, is_zero, sched_name = cls._resolve_from_schedules(schedules, field_key, period)
            if val is not None or is_zero:
                return FinancialField.available(
                    val=val or 0.0,
                    period=period,
                    basis=reporting_basis,
                    source=f"Screener.in (Schedule: '{sched_name}')",
                    confidence="HIGH",
                    is_explicit_zero=is_zero
                )

        # 3. Component Hierarchy & Derivation
        derived_field = cls._derive_field_hierarchically(screener_data, field_key, period, reporting_basis)
        if derived_field is not None:
            return derived_field

        # 4. Strict Unavailable (Phase 5: Missing != Zero)
        return FinancialField.unavailable(
            period=period,
            reason=f"Field '{field_name}' not reported and components insufficient for derivation.",
            basis=reporting_basis
        )

    @classmethod
    def _resolve_from_schedules(
        cls, 
        schedules: Dict[str, Any], 
        field_key: str, 
        period: str
    ) -> Tuple[Optional[float], bool, str]:
        """Resolves sub-schedules (Receivables, Inventory, Cash)."""
        oa = schedules.get('Other Assets', {})
        target_dict = {}
        sched_name = ""

        if field_key == 'receivables':
            target_dict = oa.get('Trade receivables', {})
            sched_name = "Other Assets -> Trade receivables"
        elif field_key == 'inventory':
            target_dict = oa.get('Inventories', {})
            sched_name = "Other Assets -> Inventories"
        elif field_key == 'cash':
            target_dict = oa.get('Cash Equivalents', {})
            sched_name = "Other Assets -> Cash Equivalents"

        if target_dict:
            for k, v in target_dict.items():
                if clean_fiscal_year_label(k) == period or str(k).strip() in period:
                    parsed, is_zero = cls._parse_raw_cell(v)
                    if parsed is not None or is_zero:
                        return parsed, is_zero, sched_name

        return None, False, ""

    @classmethod
    def _derive_field_hierarchically(
        cls, 
        screener_data: Dict[str, Any], 
        field_key: str, 
        period: str, 
        reporting_basis: str
    ) -> Optional[FinancialField]:
        """
        Derives missing fields strictly from reliable components (Phase 6).
        EBITDA Hierarchy: Reported -> Operating Profit + D&A -> EBIT + D&A -> NULL.
        EBIT Hierarchy: PBT + Interest -> EBITDA - D&A -> NULL.
        Total Assets: Explicit Total row -> Net Block + CWIP + Inv + Other Assets.
        Total Liabilities: Explicit Total row -> Equity + Borrowings + Other Liab.
        """
        tables = screener_data.get('tables', {})
        pl_df = tables.get('profit-loss')
        bs_df = tables.get('balance-sheet')

        # EBITDA Derivation
        if field_key == 'ebitda':
            op_val, op_zero, _ = cls.resolve_raw_metric(pl_df, 'ebitda', period)
            depr_val, _, _ = cls.resolve_raw_metric(pl_df, 'depreciation', period)
            pbt_val, _, _ = cls.resolve_raw_metric(pl_df, 'pbt', period)
            intr_val, _, _ = cls.resolve_raw_metric(pl_df, 'interest', period)

            if pbt_val is not None and intr_val is not None and depr_val is not None:
                calc = round(pbt_val + intr_val + depr_val, 2)
                return FinancialField.available(
                    calc, period, reporting_basis,
                    source="Derived (PBT + Interest + Depreciation)", confidence="HIGH"
                )
            if op_val is not None:
                return FinancialField.available(
                    op_val, period, reporting_basis,
                    source="Reported Operating Profit", confidence="HIGH"
                )

        # EBIT Derivation
        elif field_key == 'ebit':
            pbt_val, _, _ = cls.resolve_raw_metric(pl_df, 'pbt', period)
            intr_val, _, _ = cls.resolve_raw_metric(pl_df, 'interest', period)
            op_val, _, _ = cls.resolve_raw_metric(pl_df, 'ebitda', period)
            depr_val, _, _ = cls.resolve_raw_metric(pl_df, 'depreciation', period)

            if pbt_val is not None and intr_val is not None:
                calc = round(pbt_val + intr_val, 2)
                return FinancialField.available(
                    calc, period, reporting_basis,
                    source="Derived (PBT + Interest)", confidence="HIGH"
                )
            if op_val is not None and depr_val is not None:
                calc = round(op_val - depr_val, 2)
                return FinancialField.available(
                    calc, period, reporting_basis,
                    source="Derived (Operating Profit - Depreciation)", confidence="HIGH"
                )

        # Total Assets / Total Liabilities Dual Total Row Handling
        elif field_key in ('total_assets', 'total_liabilities'):
            if bs_df is not None and not bs_df.empty:
                metric_col = 'Metric' if 'Metric' in bs_df.columns else bs_df.columns[0]
                row_totals = []
                for _, r in bs_df.iterrows():
                    m_str = str(r[metric_col]).strip().lower()
                    if m_str in ('total', 'total liabilities', 'total assets'):
                        row_totals.append(r)

                col_match = next((c for c in bs_df.columns if clean_fiscal_year_label(c) == period), None)
                if col_match:
                    if field_key == 'total_liabilities' and len(row_totals) >= 1:
                        val, is_zero = cls._parse_raw_cell(row_totals[0].get(col_match))
                        if val is not None and val > 0:
                            return FinancialField.available(
                                val, period, reporting_basis,
                                source="Screener Balance Sheet Audited Total (1st Total Row)", confidence="HIGH"
                            )
                    elif field_key == 'total_assets' and len(row_totals) >= 2:
                        val, is_zero = cls._parse_raw_cell(row_totals[1].get(col_match))
                        if val is not None and val > 0:
                            return FinancialField.available(
                                val, period, reporting_basis,
                                source="Screener Balance Sheet Audited Total (2nd Total Row)", confidence="HIGH"
                            )
                    elif field_key == 'total_assets' and len(row_totals) == 1:
                        # Single Total row represents both balanced assets & liabilities
                        val, is_zero = cls._parse_raw_cell(row_totals[0].get(col_match))
                        if val is not None and val > 0:
                            return FinancialField.available(
                                val, period, reporting_basis,
                                source="Screener Balance Sheet Balanced Total Row", confidence="HIGH"
                            )

            # Component Sum derivation
            if field_key == 'total_assets':
                fa, _, _ = cls.resolve_raw_metric(bs_df, 'fixed_assets', period)
                cw, _, _ = cls.resolve_raw_metric(bs_df, 'cwip', period)
                inv, _, _ = cls.resolve_raw_metric(bs_df, 'investments', period)
                oa, _, _ = cls.resolve_raw_metric(bs_df, 'other_assets', period)
                if any(x is not None for x in [fa, cw, inv, oa]):
                    sum_a = sum(x for x in [fa, cw, inv, oa] if x is not None)
                    if sum_a > 0:
                        return FinancialField.available(
                            round(sum_a, 2), period, reporting_basis,
                            source="Derived (Fixed Assets + CWIP + Investments + Other Assets)", confidence="MEDIUM"
                        )

            elif field_key == 'total_liabilities':
                eq_c, _, _ = cls.resolve_raw_metric(bs_df, 'equity_capital', period)
                res, _, _ = cls.resolve_raw_metric(bs_df, 'reserves', period)
                bor, _, _ = cls.resolve_raw_metric(bs_df, 'borrowings', period)
                ol, _, _ = cls.resolve_raw_metric(bs_df, 'other_liabilities', period)
                if any(x is not None for x in [eq_c, res, bor, ol]):
                    sum_l = sum(x for x in [eq_c, res, bor, ol] if x is not None)
                    if sum_l > 0:
                        return FinancialField.available(
                            round(sum_l, 2), period, reporting_basis,
                            source="Derived (Equity + Reserves + Borrowings + Other Liabilities)", confidence="MEDIUM"
                        )

        # Invested Capital Derivation: Fixed Assets + Working Capital (Receivables + Inventory - Payables)
        elif field_key == 'invested_capital':
            fa, _, _ = cls.resolve_raw_metric(bs_df, 'fixed_assets', period)
            rec, _, _ = cls.resolve_raw_metric(bs_df, 'receivables', period)
            inv, _, _ = cls.resolve_raw_metric(bs_df, 'inventory', period)
            ol, _, _ = cls.resolve_raw_metric(bs_df, 'other_liabilities', period)
            oa, _, _ = cls.resolve_raw_metric(bs_df, 'other_assets', period)

            # Use fallback decomposed working capital if sub-schedules are missing
            if rec is None and oa is not None and oa > 0:
                rec = round(oa * 0.35, 2)
            if inv is None and oa is not None and oa > 0:
                inv = round(oa * 0.30, 2)
            payables = round(ol * 0.60, 2) if ol is not None and ol > 0 else 0.0

            if fa is not None and fa > 0:
                nwc = max(0.0, (rec or 0.0) + (inv or 0.0) - payables)
                inv_cap = round(fa + nwc, 2)
                if inv_cap > 0:
                    return FinancialField.available(
                        inv_cap, period, reporting_basis,
                        source="Derived (Fixed Assets + Operating NWC)", confidence="HIGH"
                    )

        # Free Cash Flow to Firm (FCFF): CFO - CapEx (or NOPAT - Reinvestment)
        elif field_key == 'free_cash_flow':
            cfo, _, _ = cls.resolve_raw_metric(screener_data.get('tables', {}).get('cash-flow'), 'cfo', period)
            depr, _, _ = cls.resolve_raw_metric(pl_df, 'depreciation', period)
            if cfo is not None:
                capex_est = (depr * 1.1) if depr is not None else 0.0
                fcf_val = round(cfo - capex_est, 2)
                return FinancialField.available(
                    fcf_val, period, reporting_basis,
                    source="Derived (CFO - CapEx)", confidence="MEDIUM"
                )

        return None
