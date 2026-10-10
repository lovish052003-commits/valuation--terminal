import os, sys
sys.path.insert(0, os.path.abspath('.'))
from screener_client import load_offline_company_data

def evaluate_exact_model(screener_data):
    # From screener_data
    current_price = screener_data.get('current_price', 0.0)
    shares_cr = screener_data.get('shares_in_cr', 0.0)
    
    # 1. Base Operating EBIT (Intrinsic Valuation L38)
    # From Raw FS or tables: Operating Profit - Depreciation or EBIT
    tables = screener_data.get('tables', {})
    pl_df = tables.get('profit-loss')
    bs_df = tables.get('balance-sheet')
    
    # Extract latest EBIT
    m_op = pl_df[pl_df['Metric'].str.contains('Operating Profit', case=False, na=False)]
    m_dep = pl_df[pl_df['Metric'].str.contains('Depreciation', case=False, na=False)]
    latest_op = float(m_op.iloc[0].iloc[-1]) if not m_op.empty else 0.0
    latest_dep = float(m_dep.iloc[0].iloc[-1]) if not m_dep.empty else 0.0
    base_ebit = latest_op - latest_dep
    
    # Tax rate
    tax_rate = 0.2517
    
    # Invested Capital (Intrinsic Valuation L37)
    # Fixed Assets + Net Working Capital
    m_fa = bs_df[bs_df['Metric'].str.contains('Fixed Assets|Net Block', case=False, na=False)]
    latest_fa = float(m_fa.iloc[0].iloc[-1]) if not m_fa.empty else 0.0
    m_oa = bs_df[bs_df['Metric'].str.contains('Other Assets', case=False, na=False)]
    latest_oa = float(m_oa.iloc[0].iloc[-1]) if not m_oa.empty else 0.0
    m_ol = bs_df[bs_df['Metric'].str.contains('Other Liabilities', case=False, na=False)]
    latest_ol = float(m_ol.iloc[0].iloc[-1]) if not m_ol.empty else 0.0
    working_cap = max(0.0, latest_oa - latest_ol)
    invested_capital = max(1.0, latest_fa + working_cap)
    
    # After-Tax ROIC (Intrinsic Valuation L40)
    roic = (base_ebit * (1 - tax_rate)) / invested_capital
    
    # WACC
    # Beta
    beta = 0.70 if 'fmcg' in screener_data.get('sector', '').lower() else 0.85
    rf = 0.068
    erp = 0.065
    ke = rf + beta * erp
    
    m_debt = bs_df[bs_df['Metric'].str.contains('Borrowings', case=False, na=False)]
    debt = float(m_debt.iloc[0].iloc[-1]) if not m_debt.empty else 0.0
    mcap = current_price * shares_cr
    we = mcap / (mcap + debt) if (mcap + debt) > 0 else 0.99
    wd = debt / (mcap + debt) if (mcap + debt) > 0 else 0.01
    kd_post_tax = 0.078 * (1 - tax_rate)
    wacc = we * ke + wd * kd_post_tax
    
    # Terminal Growth
    terminal_g = 0.04
    
    # Terminal ROIC (DCF D22 = MAX(ROIC, WACC))
    terminal_roic = max(roic, wacc)
    terminal_reinvest = terminal_g / terminal_roic
    
    # Year 1 Reinvestment Rate (Intrinsic Valuation L55 median of reinvestment rates)
    # In Nestlé, Year 1 reinvest was 0.65% (0.0065)
    # Let's see: terminal reinvest was 3.78% (0.0378)
    # If base reinvest is small:
    # Schedule:
    print(f"Base EBIT: {base_ebit:.2f}, Invested Cap: {invested_capital:.2f}, ROIC: {roic*100:.2f}%, WACC: {wacc*100:.2f}%, Terminal ROIC: {terminal_roic*100:.2f}%, Terminal Reinvest: {terminal_reinvest*100:.2f}%")

for t in ['NESTLEIND', 'ITC']:
    d = load_offline_company_data(t)
    print(f"\n--- {t} ---")
    evaluate_exact_model(d)
