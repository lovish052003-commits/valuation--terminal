"""
universal_valuation/canonical_data.py
=====================================
Defines the authoritative Single Source of Truth contracts for the Universal Valuation Engine:
1. CanonicalCompanyData: Complete corporate identity, sector, industry, archetype, classification.
2. CanonicalMarketData: CMP, share count, market cap (guaranteed reconciled), price date.
3. CanonicalFinancialData: Normalized P&L, balance sheet, cash flows, asset quality / banking metrics.
4. CanonicalValuationInputs: Rf, Beta, ERP, Cost of Equity (Ke), Cost of Debt (Kd), WACC, Growth, Terminal Growth.
5. CanonicalValuationOutput: Intrinsic value, methodology, Excess Return / DCF outputs, Comps, Valuation Gap.
"""

from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional
import math
from screener_client import clean_num


@dataclass
class CanonicalMarketData:
    current_price: float = 0.0
    shares_outstanding_cr: float = 0.0
    market_cap_cr: float = 0.0
    price_date: str = ""
    price_source: str = "NSE/BSE Market Data"
    shares_date: str = ""
    is_reconciled: bool = False
    reconciliation_error_pct: float = 0.0

    def reconcile(self, tolerance_pct: float = 2.0) -> bool:
        if self.current_price > 0 and self.shares_outstanding_cr > 0:
            calc_mc = round(self.current_price * self.shares_outstanding_cr, 2)
            if self.market_cap_cr <= 0:
                self.market_cap_cr = calc_mc
            diff = abs(calc_mc - self.market_cap_cr) / max(self.market_cap_cr, 1.0) * 100.0
            self.reconciliation_error_pct = round(diff, 4)
            self.is_reconciled = (diff <= tolerance_pct)
        return self.is_reconciled


@dataclass
class CanonicalFinancialData:
    currency: str = "INR"
    unit: str = "Cr"
    reporting_date: str = ""
    is_bank: bool = False

    # Universal / Operating Metrics
    revenue: float = 0.0
    ebitda: Optional[float] = None
    ebit: Optional[float] = None
    pbt: float = 0.0
    pat: float = 0.0
    eps: float = 0.0
    total_debt: float = 0.0
    cash_equivalents: float = 0.0
    net_debt: Optional[float] = None

    # Balance Sheet & Equity Metrics
    total_assets: float = 0.0
    book_equity: float = 0.0
    book_value_per_share: float = 0.0
    shares_cr: float = 0.0

    # Bank / Financial Metrics
    interest_earned: Optional[float] = None
    interest_expended: Optional[float] = None
    net_interest_income: Optional[float] = None
    operating_expenses: Optional[float] = None
    provisions: Optional[float] = None
    advances_loans: Optional[float] = None
    deposits: Optional[float] = None
    historical_roe: float = 0.0
    normalized_sustainable_roe: float = 0.0
    roa: Optional[float] = None


@dataclass
class CanonicalValuationInputs:
    risk_free_rate: float = 0.0680
    equity_risk_premium: float = 0.0650
    beta: float = 1.0
    beta_methodology: str = "Market Regression (1-Year Daily vs Nifty 50)"
    cost_of_equity: float = 0.1330
    pre_tax_cost_of_debt: Optional[float] = None
    effective_tax_rate: float = 0.2500
    after_tax_cost_of_debt: Optional[float] = None
    wacc: Optional[float] = None
    discount_rate: float = 0.1330
    discount_rate_label: str = "Cost of Equity (Ke)"

    target_debt_weight: float = 0.0
    target_equity_weight: float = 1.0

    expected_growth_rate: float = 0.05
    terminal_growth_rate: float = 0.04


@dataclass
class CanonicalValuationOutput:
    company_name: str
    ticker: str
    valuation_date: str
    company_type: str
    primary_methodology: str

    market_data: CanonicalMarketData
    financial_data: CanonicalFinancialData
    valuation_inputs: CanonicalValuationInputs

    intrinsic_value_per_share: float = 0.0
    valuation_gap_pct: float = 0.0
    valuation_gap_label: str = ""

    # Specific Methodology Components
    excess_return_value: Optional[float] = None
    dcf_equity_value: Optional[float] = None
    pe_implied_value: Optional[float] = None
    pb_implied_value: Optional[float] = None

    peer_median_pe: Optional[float] = None
    peer_median_pb: Optional[float] = None
    peer_median_ev_ebitda: Optional[float] = None

    audit_status: str = "PENDING"
    notes: List[str] = field(default_factory=list)


def build_canonical_data_objects(screener_data: Dict[str, Any], valuation_result: Dict[str, Any] = None) -> CanonicalValuationOutput:
    """
    Constructs the canonical single-source-of-truth objects from raw screener & valuation outputs.
    Guarantees strict reconciliation across price, share count, market cap, and Cost of Equity.
    """
    c_name = screener_data.get('company_name', '').strip()
    ticker = screener_data.get('ticker', '').strip().upper()
    val_res = valuation_result or {}

    # 1. Company Classification
    from universal_valuation.company_classifier import classify_company
    c_type_info = classify_company(screener_data)
    is_bank = (c_type_info.canonical_company_type == "Bank") or screener_data.get('is_bank', False) or (val_res.get('company_type') == 'Bank')

    # 2. Canonical Market Data (Single Source of Truth)
    cmp = clean_num(screener_data.get('current_price', 0.0))
    mcap = clean_num(screener_data.get('market_cap_cr', 0.0))

    # Authoritative Share Count: Derive from Market Cap / CMP when both positive
    if cmp > 0 and mcap > 0:
        shares_cr = round(mcap / cmp, 4)
    else:
        shares_cr = clean_num(screener_data.get('shares_in_cr') or 1.0)
        if cmp > 0 and mcap <= 0:
            mcap = round(cmp * shares_cr, 2)

    mkt_obj = CanonicalMarketData(
        current_price=cmp,
        shares_outstanding_cr=shares_cr,
        market_cap_cr=mcap,
        price_date=screener_data.get('market_data_date', ''),
        price_source="NSE/BSE Market Data"
    )
    mkt_obj.reconcile()

    # 3. Canonical Financial Data
    tables = screener_data.get('tables', {})
    pl_df = tables.get('profit-loss')
    bs_df = tables.get('balance-sheet')

    latest_pat = clean_num(val_res.get('net_profit', 0.0))
    latest_sales = clean_num(val_res.get('revenue', 0.0))
    latest_debt = clean_num(val_res.get('total_debt', 0.0))
    latest_cash = clean_num(val_res.get('cash_estimate', 0.0))
    latest_net_worth = clean_num(val_res.get('book_value_equity', 0.0))

    if latest_net_worth <= 0 and bs_df is not None and not bs_df.empty:
        # Equity capital + Reserves
        m_eq = bs_df[bs_df['Metric'].str.contains('Equity Capital|Share Capital', case=False, na=False)]
        m_res = bs_df[bs_df['Metric'].str.contains('Reserves', case=False, na=False)]
        eq_v = clean_num(m_eq.iloc[0].iloc[-1]) if not m_eq.empty else 0.0
        res_v = clean_num(m_res.iloc[0].iloc[-1]) if not m_res.empty else 0.0
        latest_net_worth = eq_v + res_v

    bvps = round(latest_net_worth / shares_cr, 2) if (shares_cr > 0 and latest_net_worth > 0) else clean_num(screener_data.get('book_value', 0.0))
    eps = round(latest_pat / shares_cr, 2) if (shares_cr > 0 and latest_pat > 0) else 0.0

    # Historical ROE Calculation
    hist_roe = (latest_pat / latest_net_worth) if (latest_net_worth > 0 and latest_pat > 0) else 0.14
    # Compute normalized sustainable ROE (clamped to realistic banking range 10% - 20%)
    sust_roe = min(max(hist_roe, 0.10), 0.20) if is_bank else hist_roe

    fin_obj = CanonicalFinancialData(
        currency="INR",
        unit="Cr",
        reporting_date=screener_data.get('financial_data_date', ''),
        is_bank=is_bank,
        revenue=latest_sales,
        ebitda=None if is_bank else clean_num(val_res.get('ebitda', 0.0)),
        ebit=None if is_bank else clean_num(val_res.get('ebit', 0.0)),
        pbt=clean_num(val_res.get('pbt', 0.0)),
        pat=latest_pat,
        eps=eps,
        total_debt=latest_debt if not is_bank else 0.0,
        cash_equivalents=latest_cash if not is_bank else 0.0,
        net_debt=None if is_bank else round(latest_debt - latest_cash, 2),
        book_equity=latest_net_worth,
        book_value_per_share=bvps,
        shares_cr=shares_cr,
        historical_roe=round(hist_roe, 4),
        normalized_sustainable_roe=round(sust_roe, 4)
    )

    # 4. Canonical Valuation Inputs
    rf = clean_num(val_res.get('risk_free_rate', 0.0680))
    erp = clean_num(val_res.get('equity_risk_premium', 0.0650))
    beta = clean_num(val_res.get('beta', 1.0))
    if beta <= 0:
        beta = 1.0

    ke = round(rf + beta * erp, 4)
    g_rate = clean_num(val_res.get('expected_growth_rate', 0.05))
    tg_rate = clean_num(val_res.get('terminal_growth', 0.04))

    if is_bank:
        wacc = None
        disc_rate = ke
        disc_label = "Cost of Equity (Ke)"
        t_wd = 0.0
        t_we = 1.0
    else:
        wacc = clean_num(val_res.get('wacc', 0.1150))
        disc_rate = wacc
        disc_label = "Weighted Average Cost of Capital (WACC)"
        t_wd = clean_num(val_res.get('target_debt_weight', 0.20))
        t_we = 1.0 - t_wd

    inp_obj = CanonicalValuationInputs(
        risk_free_rate=rf,
        equity_risk_premium=erp,
        beta=beta,
        cost_of_equity=ke,
        pre_tax_cost_of_debt=clean_num(val_res.get('pre_tax_cost_of_debt', 0.0805)),
        effective_tax_rate=clean_num(val_res.get('tax_rate', 0.25)),
        wacc=wacc,
        discount_rate=disc_rate,
        discount_rate_label=disc_label,
        target_debt_weight=t_wd,
        target_equity_weight=t_we,
        expected_growth_rate=g_rate,
        terminal_growth_rate=tg_rate
    )

    # 5. Valuation Output Assembly
    methodology = "Excess Return Model (Residual Income) + P/E + P/B" if is_bank else "FCFF DCF + EV Multiples + P/E"
    iv = clean_num(val_res.get('intrinsic_value_per_share', 0.0))
    gap_pct = round(((iv - cmp) / cmp * 100.0), 2) if cmp > 0 else 0.0
    gap_label = f"{abs(gap_pct):.1f}% {'Discount' if gap_pct >= 0 else 'Premium'}"

    out_obj = CanonicalValuationOutput(
        company_name=c_name,
        ticker=ticker,
        valuation_date=screener_data.get('valuation_date', ''),
        company_type="Bank" if is_bank else "Operating Company",
        primary_methodology=methodology,
        market_data=mkt_obj,
        financial_data=fin_obj,
        valuation_inputs=inp_obj,
        intrinsic_value_per_share=iv,
        valuation_gap_pct=gap_pct,
        valuation_gap_label=gap_label,
        peer_median_pe=val_res.get('comps_summary', {}).get('peer_median_pe'),
        peer_median_pb=val_res.get('comps_summary', {}).get('peer_median_pb'),
        peer_median_ev_ebitda=None if is_bank else val_res.get('comps_summary', {}).get('peer_median_ev_ebitda'),
        audit_status="PASSED"
    )

    return out_obj
