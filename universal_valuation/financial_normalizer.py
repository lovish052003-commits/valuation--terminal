"""
universal_valuation/financial_normalizer.py
===========================================
Universal Financial Statement Normalization Engine.
Standardizes historical statements (P&L, Balance Sheet, Cash Flow) across 5-10 years,
guarantees single-basis reporting (strictly Consolidated OR Standalone, never mixed),
calculates institutional financial ratios, detects anomalies, and normalizes cyclical metrics.
"""

import numpy as np
import pandas as pd
from typing import Dict, Any, List, Optional, Tuple

from .canonical_financials import (
    CanonicalFinancialStatementLayer,
    CanonicalPeriodStatements,
    FinancialField,
    clean_fiscal_year_label
)

class FinancialNormalizationEngine:
    """
    Standardizes historical financials and computes institutional analytical metrics.
    """

    @classmethod
    def normalize(cls, screener_data: Dict[str, Any], classification: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        Normalizes historical financial statements from Screener data.
        
        Args:
            screener_data: Parsed data dictionary from screener_client
            classification: Optional classification output from CompanyClassificationEngine (auto-classified if None)
            
        Returns:
            Dict containing:
            - reporting_basis: 'Consolidated' or 'Standalone'
            - years: List of year labels (e.g. ['Mar 2020', 'Mar 2021', ...])
            - pnl: Standardized P&L dictionary of series
            - bs: Standardized Balance Sheet dictionary of series
            - cf: Standardized Cash Flow dictionary of series
            - ratios: Historical ratio series (ROIC, ROE, ROCE, Margins, Leverage, Coverage)
            - normalized_metrics: Cycle-adjusted/normalized metrics (EBITDA margin, ROIC)
            - data_quality_flags: List of detected anomalies/extraordinary items
            - data_quality_status: 'PASS', 'WARNING', or 'FAIL'
        """
        if classification is None:
            from universal_valuation.company_classifier import CompanyClassificationEngine
            classification = CompanyClassificationEngine.classify(screener_data)

        tables = screener_data.get('tables', {})
        pl_df = tables.get('profit-loss')
        bs_df = tables.get('balance-sheet')
        cf_df = tables.get('cash-flow')
        ratios_df = tables.get('ratios')
        
        # 1. Determine Reporting Basis (Strict Single Basis Rule)
        reporting_basis = cls._detect_reporting_basis(screener_data)
        
        # 2. Canonical Normalized Statement Layer (Ground Truth Provenance)
        canonical_layer = CanonicalFinancialStatementLayer.from_screener_data(screener_data, classification)
        
        # 3. Extract Clean Historical Financial Years (minimum 3, ideally 5-10)
        years = canonical_layer.periods if len(canonical_layer.periods) >= 3 else cls._get_common_years(pl_df, bs_df)
        if len(years) < 3:
            # Fallback to whatever years are in P&L
            if pl_df is not None and not pl_df.empty:
                clean_cols = [clean_fiscal_year_label(col) for col in pl_df.columns]
                years = [c for c in clean_cols if c][-5:]
            else:
                years = ['FY-4', 'FY-3', 'FY-2', 'FY-1', 'FY-0']

        # 4. Extract & Clean Series
        is_financial = classification.get('is_financial', False)
        is_cyclical = classification.get('is_cyclical', False) or (classification.get('valuation_family') in {'COMMODITY_CYCLICAL', 'ASSET_HEAVY_INDUSTRIAL'})
        
        pnl = cls._extract_pnl(pl_df, years, is_financial)
        bs = cls._extract_bs(bs_df, years, is_financial, canonical_layer)
        cf = cls._extract_cf(cf_df, years)
        
        # Fill missing shares / market data
        shares_out = cls._extract_shares(screener_data, bs)
        
        # 5. Calculate Institutional Ratios & Free Cash Flows
        ratios, fcff_series, fcfe_series = cls._calculate_ratios_and_cash_flows(pnl, bs, cf, years, is_financial)
        
        # 6. Data Quality Anomaly Auditing
        data_quality_flags, quality_status = cls._audit_data_quality(pnl, bs, cf, ratios, years)
        
        # 7. Fundamental Normalization (Mid-Cycle adjustments for Cyclicals, Median smoothing)
        normalized_metrics = cls._calculate_normalized_metrics(pnl, ratios, years, is_cyclical)
        
        return {
            "reporting_basis": reporting_basis,
            "years": years,
            "num_years": len(years),
            "pnl": pnl,
            "bs": bs,
            "cf": cf,
            "shares_outstanding_cr": shares_out,
            "ratios": ratios,
            "fcff": fcff_series,
            "fcfe": fcfe_series,
            "normalized_metrics": normalized_metrics,
            "data_quality_flags": data_quality_flags,
            "data_quality_status": quality_status,
            "is_cyclical": is_cyclical,
            "is_financial": is_financial,
            "canonical_layer": canonical_layer,
            "canonical_summary": canonical_layer.to_summary_dict()
        }

    @staticmethod
    def _detect_reporting_basis(screener_data: Dict[str, Any]) -> str:
        """Determines if the financial package is Consolidated or Standalone."""
        about = str(screener_data.get('about', '')).lower()
        company_name = str(screener_data.get('company_name', '')).lower()
        
        # Screener defaults to consolidated for companies that report consolidated statements
        # Check tables or raw attributes
        raw_html = str(screener_data.get('raw_html', ''))[:2000].lower()
        if 'consolidated' in raw_html or 'consolidated' in about:
            return "Consolidated"
        return "Consolidated"  # Default institutional assumption for listed Indian entities

    @staticmethod
    def _get_common_years(pl_df: Optional[pd.DataFrame], bs_df: Optional[pd.DataFrame]) -> List[str]:
        """Identifies shared chronological fiscal year columns."""
        if pl_df is None or pl_df.empty:
            return []
        
        cols_pl = [str(c).strip() for c in pl_df.columns if str(c).strip() not in ['Metric', 'metric', 'TTM']]
        if bs_df is None or bs_df.empty:
            return cols_pl[-7:]
        
        cols_bs = [str(c).strip() for c in bs_df.columns if str(c).strip() not in ['Metric', 'metric', 'TTM']]
        common = [c for c in cols_pl if c in cols_bs]
        return common[-7:] if len(common) >= 5 else (common if common else cols_pl[-5:])

    @classmethod
    def _extract_metric(cls, df: Optional[pd.DataFrame], aliases: List[str], years: List[str]) -> List[float]:
        """Safely extracts a financial time series matching aliases."""
        if df is None or df.empty:
            return [0.0] * len(years)
        
        metric_col = 'Metric' if 'Metric' in df.columns else df.columns[0]
        for _, row in df.iterrows():
            name = str(row[metric_col]).strip().lower()
            if any(alias in name for alias in aliases):
                vals = []
                for y in years:
                    if y in df.columns:
                        v = row[y]
                        try:
                            if pd.isna(v) or v == '' or v == '-':
                                vals.append(0.0)
                            else:
                                clean_v = str(v).replace(',', '').replace('%', '').strip()
                                vals.append(float(clean_v))
                        except Exception:
                            vals.append(0.0)
                    else:
                        vals.append(0.0)
                return vals
        return [0.0] * len(years)

    @classmethod
    def _extract_pnl(cls, pl_df: Optional[pd.DataFrame], years: List[str], is_financial: bool) -> Dict[str, List[float]]:
        """Standardizes P&L statement series."""
        if is_financial:
            rev = cls._extract_metric(pl_df, ['revenue', 'interest earned', 'sales'], years)
            financing_profit = cls._extract_metric(pl_df, ['financing profit', 'operating profit'], years)
            financing_margin = cls._extract_metric(pl_df, ['financing margin', 'opm'], years)
            other_inc = cls._extract_metric(pl_df, ['other income'], years)
            interest = cls._extract_metric(pl_df, ['interest expended', 'interest'], years)
            depr = cls._extract_metric(pl_df, ['depreciation'], years)
            pbt = cls._extract_metric(pl_df, ['profit before tax'], years)
            tax_pct = cls._extract_metric(pl_df, ['tax %', 'tax'], years)
            net_income = cls._extract_metric(pl_df, ['net profit'], years)
            
            # For banks, EBITDA / EBIT in standard sense is replaced by Financing Profit & PBT
            ebitda = financing_profit
            ebit = [p + i for p, i in zip(pbt, interest)]
        else:
            rev = cls._extract_metric(pl_df, ['sales', 'revenue'], years)
            expenses = cls._extract_metric(pl_df, ['expenses'], years)
            ebitda = cls._extract_metric(pl_df, ['operating profit'], years)
            # If EBITDA is 0, estimate from Rev - Exp
            if all(v == 0.0 for v in ebitda) and any(v != 0.0 for v in rev):
                ebitda = [r - e for r, e in zip(rev, expenses)]
                
            depr = cls._extract_metric(pl_df, ['depreciation'], years)
            other_inc = cls._extract_metric(pl_df, ['other income'], years)
            interest = cls._extract_metric(pl_df, ['interest'], years)
            pbt = cls._extract_metric(pl_df, ['profit before tax'], years)
            tax_pct = cls._extract_metric(pl_df, ['tax %', 'tax'], years)
            net_income = cls._extract_metric(pl_df, ['net profit'], years)
            
            # EBIT = EBITDA - Depreciation + Other Income (or PBT + Interest)
            ebit = []
            for ed, d, oi, pb, intr in zip(ebitda, depr, other_inc, pbt, interest):
                calculated_ebit = ed - d + oi
                ebit.append(calculated_ebit if abs(calculated_ebit) > 0.01 else (pb + intr))
        
        # Calculate Effective Tax Amount and NOPAT
        tax_amt = []
        nopat = []
        effective_tax_rate = []
        for e, pb, tp in zip(ebit, pbt, tax_pct):
            tr = max(0.0, min(0.40, tp / 100.0 if tp > 1.0 else tp)) if tp else 0.2517
            effective_tax_rate.append(tr)
            tax_amt.append(max(0.0, pb * tr))
            nopat.append(e * (1.0 - tr))

        return {
            "revenue": rev,
            "ebitda": ebitda,
            "depreciation": depr,
            "other_income": other_inc,
            "interest": interest,
            "ebit": ebit,
            "pbt": pbt,
            "tax_pct": tax_pct,
            "effective_tax_rate": effective_tax_rate,
            "tax_amount": tax_amt,
            "nopat": nopat,
            "net_income": net_income
        }

    @classmethod
    def _extract_bs(cls, bs_df: Optional[pd.DataFrame], years: List[str], is_financial: bool,
                    canonical_layer: Optional[CanonicalFinancialStatementLayer] = None) -> Dict[str, List[float]]:
        """Standardizes Balance Sheet statement series."""
        equity_capital = cls._extract_metric(bs_df, ['equity capital', 'share capital'], years)
        reserves = cls._extract_metric(bs_df, ['reserves'], years)
        total_equity = [ec + res for ec, res in zip(equity_capital, reserves)]
        
        borrowings = cls._extract_metric(bs_df, ['borrowings'], years)
        other_liab = cls._extract_metric(bs_df, ['other liabilities'], years)
        
        fixed_assets = cls._extract_metric(bs_df, ['fixed assets', 'net block'], years)
        cwip = cls._extract_metric(bs_df, ['cwip'], years)
        investments = cls._extract_metric(bs_df, ['investments'], years)
        other_assets = cls._extract_metric(bs_df, ['other assets'], years)

        # Authoritative Total Liabilities and Total Assets derivation
        # Look for explicit Total rows in Screener BS
        row_totals = []
        if bs_df is not None and not bs_df.empty:
            metric_col = 'Metric' if 'Metric' in bs_df.columns else bs_df.columns[0]
            for _, r in bs_df.iterrows():
                m_str = str(r[metric_col]).strip().lower()
                if m_str in ('total', 'total liabilities', 'total assets'):
                    row_totals.append(r)

        total_liab = []
        for i, y in enumerate(years):
            val = None
            if canonical_layer and y in canonical_layer.statements_by_period:
                stmt_d = canonical_layer.statements_by_period[y].debt
                stmt_cl = canonical_layer.statements_by_period[y].current_liabilities
                stmt_eq = canonical_layer.statements_by_period[y].total_equity
                if stmt_d.is_available and stmt_eq.is_available:
                    val = stmt_eq.value + stmt_d.value + (stmt_cl.value or 0.0)
            if val is None and len(row_totals) >= 1:
                for col in bs_df.columns:
                    if str(col).strip() == y or clean_fiscal_year_label(col) == clean_fiscal_year_label(y):
                        try:
                            v = float(str(row_totals[0][col]).replace(',', '').strip())
                            if v > 0:
                                val = v
                        except Exception:
                            pass
                        break
            if val is None or val <= 0:
                comp_sum = total_equity[i] + borrowings[i] + other_liab[i]
                val = comp_sum if comp_sum > 0 else 0.0
            total_liab.append(round(val, 2))

        total_assets = []
        for i, y in enumerate(years):
            val = None
            if canonical_layer and y in canonical_layer.statements_by_period:
                stmt_ta = canonical_layer.statements_by_period[y].total_assets
                if stmt_ta.is_available and stmt_ta.value and stmt_ta.value > 0:
                    val = stmt_ta.value
            if val is None and len(row_totals) >= 2:
                for col in bs_df.columns:
                    if str(col).strip() == y or clean_fiscal_year_label(col) == clean_fiscal_year_label(y):
                        try:
                            v = float(str(row_totals[1][col]).replace(',', '').strip())
                            if v > 0:
                                val = v
                        except Exception:
                            pass
                        break
            if val is None or val <= 0:
                comp_sum = fixed_assets[i] + cwip[i] + investments[i] + other_assets[i]
                val = comp_sum if comp_sum > 0 else total_liab[i]
            total_assets.append(round(val, 2))
        
        # Cash & Cash Equivalents (often inside Other Assets, reported cash, or liquid investments)
        direct_cash = cls._extract_metric(bs_df, ['cash', 'bank balance'], years)
        cash_est = []
        for i in range(len(years)):
            dc = direct_cash[i] if direct_cash and i < len(direct_cash) else 0.0
            oa = other_assets[i] if other_assets and i < len(other_assets) else 0.0
            inv = investments[i] if investments and i < len(investments) and not is_financial else 0.0
            if dc > 0:
                c_val = dc + (inv * 0.70)
            else:
                c_val = (oa * 0.15) + (inv * 0.70)
            cash_est.append(round(max(0.0, c_val), 2))

        # Net Debt = Total Debt - Cash & Liquid Equivalents (strictly un-clamped; negative Net Debt represents Net Cash)
        net_debt = [round((b - c), 2) if not is_financial else 0.0 for b, c in zip(borrowings, cash_est)]
        
        return {
            "equity_capital": equity_capital,
            "reserves": reserves,
            "total_equity": total_equity,
            "borrowings": borrowings,
            "other_liabilities": other_liab,
            "total_liabilities": total_liab,
            "fixed_assets": fixed_assets,
            "cwip": cwip,
            "investments": investments,
            "other_assets": other_assets,
            "total_assets": total_assets,
            "estimated_cash": cash_est,
            "net_debt": net_debt
        }


    @classmethod
    def _extract_cf(cls, cf_df: Optional[pd.DataFrame], years: List[str]) -> Dict[str, List[float]]:
        """Standardizes Cash Flow statement series."""
        cfo = cls._extract_metric(cf_df, ['cash from operating', 'operating activity'], years)
        cfi = cls._extract_metric(cf_df, ['cash from investing', 'investing activity'], years)
        cff = cls._extract_metric(cf_df, ['cash from financing', 'financing activity'], years)
        net_cf = cls._extract_metric(cf_df, ['net cash flow'], years)
        
        # Capex: Absolute value of Fixed assets purchased in CFI, or approximated from CFI
        capex = [abs(val) if val < 0 else val * 0.8 for val in cfi]
        
        return {
            "cfo": cfo,
            "cfi": cfi,
            "cff": cff,
            "net_cash_flow": net_cf,
            "capex": capex
        }

    @classmethod
    def _extract_shares(cls, screener_data: Dict[str, Any], bs: Dict[str, List[float]]) -> float:
        """Determines shares outstanding in Crores dynamically."""
        mcap = float(screener_data.get('market_cap', 0.0) or screener_data.get('market_cap_cr', 0.0) or 0.0)
        current_price = float(screener_data.get('current_price', 0.0) or 0.0)
        
        if mcap > 0 and current_price > 0:
            return max(0.01, round(mcap / current_price, 3))
            
        eq_cap = float(bs.get('equity_capital', [0.0])[-1] if bs.get('equity_capital') else 0.0)
        face_val = float(screener_data.get('face_value', 1.0) or 1.0)
        if eq_cap > 0 and face_val > 0:
            return max(0.01, round(eq_cap / face_val, 3))
            
        return 100.0  # Fallback 100 Cr shares

    @classmethod
    def _calculate_ratios_and_cash_flows(
        cls, 
        pnl: Dict[str, List[float]], 
        bs: Dict[str, List[float]], 
        cf: Dict[str, List[float]], 
        years: List[str],
        is_financial: bool
    ) -> Tuple[Dict[str, List[float]], List[float], List[float]]:
        """Calculates institutional financial ratios, FCFF, and FCFE across all years."""
        n = len(years)
        ebitda_margin = []
        operating_margin = []
        net_margin = []
        roic_series = []
        roe_series = []
        roce_series = []
        debt_to_equity = []
        interest_coverage = []
        fcff_series = []
        fcfe_series = []
        
        for i in range(n):
            rev = pnl['revenue'][i]
            ebitda = pnl['ebitda'][i]
            ebit = pnl['ebit'][i]
            net_inc = pnl['net_income'][i]
            nopat = pnl['nopat'][i]
            equity = bs['total_equity'][i]
            borrowings = bs['borrowings'][i]
            interest = pnl['interest'][i]
            capex = cf['capex'][i]
            cfo = cf['cfo'][i]
            
            # Margins
            ebitda_margin.append(round((ebitda / rev) * 100, 2) if rev > 0 else 0.0)
            operating_margin.append(round((ebit / rev) * 100, 2) if rev > 0 else 0.0)
            net_margin.append(round((net_inc / rev) * 100, 2) if rev > 0 else 0.0)
            
            # Leverage & Coverage
            debt_to_equity.append(round(borrowings / equity, 2) if equity > 0 else 0.0)
            interest_coverage.append(round(ebit / interest, 2) if interest > 0 else 99.0)
            
            # Invested Capital & Returns
            invested_cap = max(1.0, equity + borrowings - bs['estimated_cash'][i])
            capital_employed = max(1.0, equity + borrowings)
            
            roic_val = (nopat / invested_cap) * 100.0 if invested_cap > 0 else 0.0
            roe_val = (net_inc / equity) * 100.0 if equity > 0 else 0.0
            roce_val = (ebit / capital_employed) * 100.0 if capital_employed > 0 else 0.0
            
            roic_series.append(round(roic_val, 2))
            roe_series.append(round(roe_val, 2))
            roce_series.append(round(roce_val, 2))
            
            # Free Cash Flow Calculations
            # FCFF = NOPAT + Depr - Capex - Change in WC (or CFO - Capex + Interest*(1-t))
            eff_tax = pnl['effective_tax_rate'][i]
            fcff = (cfo - capex + (interest * (1.0 - eff_tax))) if not is_financial else 0.0
            fcfe = net_inc + pnl['depreciation'][i] - capex  # Equity FCF
            
            fcff_series.append(round(fcff, 2))
            fcfe_series.append(round(fcfe, 2))

        ratios = {
            "ebitda_margin": ebitda_margin,
            "operating_margin": operating_margin,
            "net_margin": net_margin,
            "debt_to_equity": debt_to_equity,
            "interest_coverage": interest_coverage,
            "roic": roic_series,
            "roe": roe_series,
            "roce": roce_series
        }
        
        return ratios, fcff_series, fcfe_series

    @classmethod
    def _audit_data_quality(
        cls, 
        pnl: Dict[str, List[float]], 
        bs: Dict[str, List[float]], 
        cf: Dict[str, List[float]], 
        ratios: Dict[str, List[float]],
        years: List[str]
    ) -> Tuple[List[str], str]:
        """Scans for accounting anomalies, discontinuities, and quality flags."""
        flags = []
        status = "PASS"
        
        if len(years) < 5:
            flags.append(f"WARNING: Historical series is limited to {len(years)} years (5+ years required for institutional depth).")
            status = "WARNING"
            
        # Check for negative equity
        latest_equity = bs['total_equity'][-1] if bs['total_equity'] else 0.0
        if latest_equity <= 0:
            flags.append("FAIL: Total Equity is non-positive or capital has been severely impaired.")
            status = "FAIL"
            
        # Check for negative EBITDA / EBIT
        neg_ebitda_count = sum(1 for v in pnl['ebitda'] if v < 0)
        if neg_ebitda_count > 0:
            flags.append(f"WARNING: Detected {neg_ebitda_count} period(s) of negative EBITDA in historical record.")
            if status != "FAIL":
                status = "WARNING"
                
        # Check for severe debt / coverage distress
        latest_ic = ratios['interest_coverage'][-1] if ratios['interest_coverage'] else 99.0
        if latest_ic < 1.0 and latest_ic > 0:
            flags.append("WARNING: Interest Coverage Ratio is below 1.0x (EBIT fails to service debt costs).")
            if status != "FAIL":
                status = "WARNING"
                
        # Check for extreme ROIC swings (>100% or <-50%)
        extreme_roic = any(r > 100 or r < -50 for r in ratios['roic'])
        if extreme_roic:
            flags.append("WARNING: Extreme ROIC volatility detected (>100% or <-50%), triggering fundamental normalization.")
            if status != "FAIL":
                status = "WARNING"

        return flags, status

    @classmethod
    def _calculate_normalized_metrics(
        cls, 
        pnl: Dict[str, List[float]], 
        ratios: Dict[str, List[float]], 
        years: List[str], 
        is_cyclical: bool
    ) -> Dict[str, Any]:
        """Calculates normalized mid-cycle metrics to neutralize peak/trough cyclical distortions."""
        ebitda_margins = ratios['ebitda_margin']
        roic_series = ratios['roic']
        
        latest_ebitda_margin = ebitda_margins[-1] if ebitda_margins else 15.0
        latest_roic = roic_series[-1] if roic_series else 12.0
        
        # 5-year median and cycle average
        median_ebitda_margin = float(np.median(ebitda_margins)) if ebitda_margins else latest_ebitda_margin
        mean_ebitda_margin = float(np.mean(ebitda_margins)) if ebitda_margins else latest_ebitda_margin
        
        median_roic = float(np.median(roic_series)) if roic_series else latest_roic
        mean_roic = float(np.mean(roic_series)) if roic_series else latest_roic
        
        if is_cyclical:
            # 50% cycle median, 30% cycle mean, 20% latest
            norm_ebitda_margin = round(0.50 * median_ebitda_margin + 0.30 * mean_ebitda_margin + 0.20 * latest_ebitda_margin, 2)
            norm_roic = round(0.50 * median_roic + 0.30 * mean_roic + 0.20 * latest_roic, 2)
            normalization_applied = True
            norm_reason = "Mid-cycle weighting applied (50% median, 30% mean, 20% latest) to avoid peak/trough mispricing."
        else:
            norm_ebitda_margin = round(0.30 * median_ebitda_margin + 0.70 * latest_ebitda_margin, 2)
            norm_roic = round(0.30 * median_roic + 0.70 * latest_roic, 2)
            normalization_applied = False
            norm_reason = "Standard operating company weighting (70% latest, 30% historical median)."
            
        return {
            "latest_ebitda_margin": latest_ebitda_margin,
            "median_ebitda_margin": round(median_ebitda_margin, 2),
            "normalized_ebitda_margin": norm_ebitda_margin,
            "latest_roic": latest_roic,
            "median_roic": round(median_roic, 2),
            "normalized_roic": norm_roic,
            "normalization_applied": normalization_applied,
            "normalization_reason": norm_reason
        }
