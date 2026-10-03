"""
scratch/adversarial_audit_probe.py
==================================
Adversarial Probe across 8 Materially Different Company Types:
1. Normal Industrial Company (ITC / LT)
2. Bank (SBIN)
3. NBFC (BAJFINANCE)
4. IT Company (INFY)
5. Pharma Company (SUNPHARMA)
6. Cyclical Commodity Company (TATASTEEL)
7. Conglomerate (ADANIENT)
8. Loss-Making Growth Company (Synthetic Turnaround)
"""

import os
import sys
import traceback
import pandas as pd
import numpy as np

ws_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ws_root not in sys.path:
    sys.path.insert(0, ws_root)

import valuation_engine
from screener_client import fetch_company_data

def probe_company(name, data):
    print("\n" + "="*80)
    print(f"PROBING: {name}")
    print("="*80)
    try:
        res = valuation_engine.calculate_valuation(data)
        cm = res.get('company_master', {})
        clf = res.get('company_classification', {})
        meth = res.get('valuation_methods', {})
        qual = res.get('model_quality', {})
        dcf = res.get('universal_engine', {}).get('dcf', {})
        coc = res.get('universal_engine', {}).get('cost_of_capital', {})
        peers = res.get('universal_engine', {}).get('peers', {})
        recon = res.get('universal_engine', {}).get('reconciliation', {})
        
        print(f"  Company Name     : {cm.get('company_name')}")
        print(f"  Ticker           : {cm.get('ticker')}")
        print(f"  Sector / Industry: {clf.get('sector')} / {clf.get('industry')}")
        print(f"  Company Type     : {clf.get('company_type')}")
        print(f"  Valuation Family : {clf.get('valuation_family')}")
        print(f"  Primary Method   : {meth.get('primary_method')}")
        print(f"  Shares Master    : {cm.get('shares_outstanding')} Cr")
        print(f"  Shares Engine    : {res.get('shares_cr')} Cr")
        print(f"  Market Cap       : Rs. {cm.get('market_cap')} Cr")
        print(f"  Current Price    : Rs. {cm.get('current_price')}")
        print(f"  Total Debt / Cash: Rs. {cm.get('total_debt')} Cr / Rs. {cm.get('cash')} Cr")
        print(f"  Net Debt         : Rs. {cm.get('net_debt')} Cr")
        print(f"  Reporting Basis  : {cm.get('reporting_basis')}")
        print(f"  WACC / Ke        : {coc.get('wacc')}% (Ke: {coc.get('cost_of_equity')}%)")
        print(f"  Beta             : {coc.get('beta')}")
        print(f"  Terminal Growth  : {dcf.get('terminal_growth')}%")
        print(f"  Intrinsic DCF    : Rs. {res.get('intrinsic_value_per_share')}")
        if 'sotp' in res.get('universal_engine', {}) and res['universal_engine']['sotp']:
            sotp = res['universal_engine']['sotp']
            print(f"  Intrinsic SOTP   : Rs. {sotp.get('implied_value_per_share')}")
        print(f"  Valuation Gap    : {recon.get('valuation_gap', {}).get('gap_label')}")
        print(f"  Peer Count       : {len(peers.get('peers', []))}")
        print(f"  Quality Status   : {qual.get('overall_status')} | Fails: {qual.get('fail_count')} | Warns: {qual.get('warning_count')}")
        
        # Check for failures
        if qual.get('fail_count', 0) > 0:
            print("  FAILURES DETECTED:")
            for t in qual.get('tests', []):
                if t['status'] == 'FAIL':
                    print(f"    [FAIL] Test {t['test_id']} ({t['test_name']}): {t['details']}")
                    
        # Check for warnings
        if qual.get('warning_count', 0) > 0:
            print("  WARNINGS DETECTED:")
            for t in qual.get('tests', []):
                if t['status'] == 'WARNING':
                    print(f"    [WARN] Test {t['test_id']} ({t['test_name']}): {t['details']}")

        # Deep Consistency Verifications
        # 1. Share count match
        if abs(cm.get('shares_outstanding', 0) - res.get('shares_cr', 0)) > 0.01:
            print(f"  [ANOMALY] Share count mismatch: Master {cm.get('shares_outstanding')} vs Engine {res.get('shares_cr')}")
            
        # 2. Market Cap ≈ Price * Shares
        calc_mcap = round(cm.get('current_price', 0) * cm.get('shares_outstanding', 0), 2)
        diff_pct = abs(calc_mcap - cm.get('market_cap', 0)) / max(cm.get('market_cap', 1), 1.0) * 100.0
        if diff_pct > 5.0:
            print(f"  [ANOMALY] Market Cap discrepancy > 5%: Calc {calc_mcap} vs Reported {cm.get('market_cap')} ({diff_pct:.2f}%)")

        # 3. Method suitability
        if clf.get('is_financial') and 'FCFF' in str(meth.get('primary_method')):
            print("  [CRITICAL ANOMALY] Financial institution assigned FCFF DCF!")
            
        # 4. Loss making multiples
        pnl = res.get('universal_engine', {}).get('financial_normalization', {}).get('pnl', {})
        ni = pnl.get('net_income', [0.0])[-1] if pnl.get('net_income') else 0.0
        if ni <= 0 and 'P/E' in meth.get('applicable_multiples', []):
            print("  [CRITICAL ANOMALY] Negative net income assigned P/E multiple!")

    except Exception as e:
        print(f"  [CRASH] {e}")
        traceback.print_exc()

def run_probe():
    tickers = [
        ("1. Industrial Company", "ITC"),
        ("2. Bank", "SBIN"),
        ("3. NBFC", "BAJAJHFL"),
        ("4. IT Company", "INFY"),
        ("5. Pharma Company", "SUNPHARMA"),
        ("6. Cyclical Commodity Company", "TATASTEEL"),
        ("7. Conglomerate", "ADANIENT"),
    ]
    
    for label, tick in tickers:
        try:
            data = fetch_company_data(tick)
            probe_company(f"{label} ({tick})", data)
        except Exception as e:
            print(f"Error fetching {tick}: {e}")

    # 8. Loss-Making Growth Company (Synthetic Turnaround)
    loss_maker_data = {
        'company_name': 'AeroCloud Systems Ltd',
        'ticker': 'AEROCLOUD',
        'sector': 'Technology',
        'industry': 'Cloud & SaaS',
        'about': 'AeroCloud provides next-generation autonomous flight orchestration software, investing aggressively in R&D ahead of enterprise monetization.',
        'current_price': 350.0,
        'market_cap': 7000.0,
        'market_cap_cr': 7000.0,
        'face_value': 10.0,
        'tables': {
            'profit-loss': pd.DataFrame({
                'Metric': ['Sales', 'Expenses', 'Operating Profit', 'Depreciation', 'Interest', 'Profit before tax', 'Net profit', 'EPS in Rs'],
                'Mar 2021': [300, 420, -120, 20, 5, -145, -145, -7.25],
                'Mar 2022': [600, 750, -150, 30, 8, -188, -188, -9.40],
                'Mar 2023': [1200, 1400, -200, 45, 12, -257, -257, -12.85],
                'Mar 2024': [2200, 2350, -150, 60, 15, -225, -225, -11.25],
            }),
            'balance-sheet': pd.DataFrame({
                'Metric': ['Equity Capital', 'Reserves', 'Borrowings', 'Other Liabilities', 'Total Liabilities', 'Fixed Assets', 'Other Assets', 'Total Assets'],
                'Mar 2021': [200, 500, 50, 40, 790, 150, 640, 790],
                'Mar 2022': [200, 1100, 80, 70, 1450, 280, 1170, 1450],
                'Mar 2023': [200, 1900, 120, 110, 2330, 450, 1880, 2330],
                'Mar 2024': [200, 2700, 160, 150, 3210, 650, 2560, 3210],
            }),
            'cash-flow': pd.DataFrame({
                'Metric': ['Cash from Operating Activity', 'Cash from Investing Activity', 'Cash from Financing Activity', 'Net Cash Flow'],
                'Mar 2021': [-80, -120, 400, 200],
                'Mar 2022': [-110, -180, 600, 310],
                'Mar 2023': [-130, -240, 900, 530],
                'Mar 2024': [-90, -310, 850, 450],
            }),
            'ratios': pd.DataFrame({
                'Metric': ['Debtor Days', 'Inventory Days', 'Days Payable', 'Cash Conversion Cycle', 'Working Capital Days'],
                'Mar 2021': [25, 5, 35, -5, 10],
                'Mar 2022': [26, 7, 36, -3, 12],
                'Mar 2023': [24, 6, 34, -4, 11],
                'Mar 2024': [22, 5, 32, -5, 9],
            })
        }
    }
    probe_company("8. Loss-Making Growth Company (AEROCLOUD)", loss_maker_data)

if __name__ == '__main__':
    run_probe()
