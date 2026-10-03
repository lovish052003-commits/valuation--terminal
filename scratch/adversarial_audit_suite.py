"""
scratch/adversarial_audit_suite.py
==================================
Comprehensive Universal Adversarial Audit Suite.
Tests the valuation engine across 18+ company archetypes and corner cases:
 1. Steel Company (Commodity Cyclical)
 2. Cement Company (Asset Heavy Industrial)
 3. Pharma Company (Life Sciences / IP)
 4. IT Company (Tech SaaS / Services)
 5. FMCG Company (Consumer Staples)
 6. Automobile Company (OEM Discretionary)
 7. NBFC (Non-Banking Financial Intermediation)
 8. Commercial Bank (Deposit-Taking Lending)
 9. Insurance Company (Underwriting Float / Solvency)
10. Loss-Making Company (Negative Net Income)
11. Highly Leveraged Company (Debt Overhang / Distress)
12. Cash-Rich Company (Net Cash Surplus Credit)
13. Cyclical Company (Mid-Cycle Margin & ROIC Normalization)
14. Company with Incomplete Data (Missing Balance Sheet / Cash Flow rows)
15. Company with Only 3 Years of History (Limited Disclosure / Short Horizon)
16. Company with Negative EBITDA (Operating Loss / Multiple Suppression)
17. Company with Negative Net Income (P/E Multiple Suppression)
18. Distressed Company with Negative Book Value (Regulatory Capital Deficit)
19. Adversarial Boundary: WACC <= Terminal Growth (StrictGordon Singularity Abort)
20. Adversarial Identity: Target "Tata Steel Ltd" vs Candidate "Tata Steel Limited" (Self-Peer Zero Contamination)
"""

import os
import sys
import unittest
import pandas as pd
import numpy as np

# Ensure workspace root is in path
ws_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ws_root not in sys.path:
    sys.path.insert(0, ws_root)

from universal_valuation import (
    CompanyClassificationEngine,
    ValuationMethodSelector,
    FinancialNormalizationEngine,
    ForecastEngine,
    CostOfCapitalEngine,
    SustainableROICEngine,
    PeerSelectionEngine,
    PeerSelectionError,
    DCFEngine,
    ValuationReconciliationEngine,
    ValuationConfidenceEngine,
    ModelQualityEngine
)
from screener_client import fetch_company_data
import valuation_engine

class UniversalAdversarialAuditSuite(unittest.TestCase):

    def setUp(self):
        print("\n" + "="*80)

    # -------------------------------------------------------------------------
    # TEST 1: Steel Company (Commodity Cyclical)
    # -------------------------------------------------------------------------
    def test_01_steel_company(self):
        print("[TEST 1] Steel Company Archetype (TATASTEEL / JSW)")
        data = fetch_company_data('TATASTEEL')
        c = CompanyClassificationEngine.classify(data)
        self.assertEqual(c['valuation_family'], 'COMMODITY_CYCLICAL')
        self.assertTrue(c['is_cyclical'])
        self.assertFalse(c['is_financial'])
        
        m = ValuationMethodSelector.select_methods(c, data)
        self.assertEqual(m['primary_method'], 'FCFF_DCF')
        self.assertTrue(m['requires_cyclical_normalization'])
        
        fn = FinancialNormalizationEngine.normalize(data, c)
        self.assertTrue(fn['normalized_metrics']['normalization_applied'])
        print(f"  -> Steel mid-cycle normalized EBITDA margin: {fn['normalized_metrics']['normalized_ebitda_margin']}%")

    # -------------------------------------------------------------------------
    # TEST 2: Cement Company (Asset-Heavy Industrial / Cyclical)
    # -------------------------------------------------------------------------
    def test_02_cement_company(self):
        print("[TEST 2] Cement Company Archetype (Synthetic / Sector)")
        cement_data = {
            'company_name': 'Ambuja Cements Limited',
            'ticker': 'AMBUJACEM',
            'sector': 'Cement',
            'industry': 'Cement & Construction Materials',
            'about': 'Ambuja Cements is a leading manufacturer of cement and clinker across India.',
            'current_price': 580.0,
            'market_cap': 140000.0,
            'tables': {
                'profit-loss': pd.DataFrame({
                    'Metric': ['Sales', 'Expenses', 'Operating Profit', 'Depreciation', 'Profit before tax', 'Net profit'],
                    'Mar 2021': [24000, 19000, 5000, 1100, 4200, 3100],
                    'Mar 2022': [28000, 22000, 6000, 1200, 5100, 3800],
                    'Mar 2023': [32000, 26000, 6000, 1300, 5000, 3700],
                    'Mar 2024': [33000, 26500, 6500, 1400, 5400, 4000],
                }),
                'balance-sheet': pd.DataFrame({
                    'Metric': ['Equity Capital', 'Reserves', 'Borrowings', 'Fixed Assets', 'Other Assets'],
                    'Mar 2021': [400, 30000, 500, 22000, 10000],
                    'Mar 2022': [400, 33000, 400, 24000, 11000],
                    'Mar 2023': [400, 36000, 300, 26000, 12000],
                    'Mar 2024': [400, 39000, 200, 28000, 13000],
                })
            }
        }
        c = CompanyClassificationEngine.classify(cement_data)
        self.assertEqual(c['valuation_family'], 'COMMODITY_CYCLICAL')
        self.assertTrue(c['is_cyclical'])
        m = ValuationMethodSelector.select_methods(c, cement_data)
        self.assertTrue(m['requires_cyclical_normalization'])
        print(f"  -> Cement correctly identified as {c['valuation_family']} with cyclical normalization required.")

    # -------------------------------------------------------------------------
    # TEST 3: Pharma Company (Life Sciences / Healthcare)
    # -------------------------------------------------------------------------
    def test_03_pharma_company(self):
        print("[TEST 3] Pharma Company Archetype (SUNPHARMA)")
        data = fetch_company_data('SUNPHARMA')
        c = CompanyClassificationEngine.classify(data)
        self.assertEqual(c['valuation_family'], 'PHARMA_HEALTHCARE')
        self.assertFalse(c['is_financial'])
        self.assertFalse(c['is_cyclical'])
        m = ValuationMethodSelector.select_methods(c, data)
        self.assertEqual(m['primary_method'], 'FCFF_DCF')
        print(f"  -> Pharma correctly classified as {c['valuation_family']}, primary method: {m['primary_method']}")

    # -------------------------------------------------------------------------
    # TEST 4: IT Company (Tech SaaS / IP Services)
    # -------------------------------------------------------------------------
    def test_04_it_company(self):
        print("[TEST 4] IT Company Archetype (INFY)")
        data = fetch_company_data('INFY')
        c = CompanyClassificationEngine.classify(data)
        self.assertEqual(c['valuation_family'], 'TECH_SAAS')
        self.assertFalse(c['is_financial'])
        fn = FinancialNormalizationEngine.normalize(data, c)
        # Verify IT has negative net debt (Net Cash)
        net_debt = fn['bs']['net_debt'][-1]
        self.assertLessEqual(net_debt, 0.0, f"Expected INFY to have Net Cash (Net Debt <= 0), got: {net_debt}")
        print(f"  -> IT Company classified as {c['valuation_family']} with Net Cash: Rs. {abs(net_debt):.1f} Cr")

    # -------------------------------------------------------------------------
    # TEST 5: FMCG Company (Consumer Staples)
    # -------------------------------------------------------------------------
    def test_05_fmcg_company(self):
        print("[TEST 5] FMCG Company Archetype (NESTLEIND / ITC)")
        data = fetch_company_data('NESTLEIND')
        c = CompanyClassificationEngine.classify(data)
        self.assertEqual(c['valuation_family'], 'CONSUMER_FMCG')
        self.assertFalse(c['is_financial'])
        print(f"  -> FMCG correctly classified as {c['valuation_family']}")

    # -------------------------------------------------------------------------
    # TEST 6: Automobile Company (OEM Discretionary)
    # -------------------------------------------------------------------------
    def test_06_automobile_company(self):
        print("[TEST 6] Automobile Company Archetype (TATAMOTORS / FORCEMOT)")
        data = fetch_company_data('TATAMOTORS')
        c = CompanyClassificationEngine.classify(data)
        self.assertEqual(c['valuation_family'], 'AUTO_ANCILLARY')
        self.assertTrue(c['is_cyclical'])
        print(f"  -> Auto OEM classified as {c['valuation_family']}")

    # -------------------------------------------------------------------------
    # TEST 7: NBFC (Non-Banking Financial Intermediation)
    # -------------------------------------------------------------------------
    def test_07_nbfc_company(self):
        print("[TEST 7] NBFC Archetype (Spread Lending / No Deposits)")
        nbfc_data = {
            'company_name': 'Bajaj Finance Limited',
            'ticker': 'BAJFINANCE',
            'sector': 'Financial Services',
            'industry': 'NBFC / Lending',
            'about': 'Bajaj Finance is a leading non-banking financial company offering consumer finance and commercial lending.',
            'current_price': 6800.0,
            'market_cap': 420000.0,
            'tables': {
                'profit-loss': pd.DataFrame({
                    'Metric': ['Interest Earned', 'Financing Profit', 'Profit before tax', 'Net profit'],
                    'Mar 2022': [27000, 9000, 9500, 7000],
                    'Mar 2023': [35000, 12000, 15000, 11500],
                    'Mar 2024': [45000, 15000, 19000, 14400],
                }),
                'balance-sheet': pd.DataFrame({
                    'Metric': ['Equity Capital', 'Reserves', 'Borrowings', 'Advances'],
                    'Mar 2022': [120, 43000, 160000, 190000],
                    'Mar 2023': [120, 54000, 210000, 240000],
                    'Mar 2024': [120, 74000, 290000, 320000],
                })
            }
        }
        c = CompanyClassificationEngine.classify(nbfc_data)
        self.assertEqual(c['valuation_family'], 'NBFC')
        self.assertTrue(c['is_financial'])
        m = ValuationMethodSelector.select_methods(c, nbfc_data)
        self.assertEqual(m['primary_method'], 'EXCESS_RETURN')
        self.assertIn('FCFF_DCF', m['excluded_methods'])
        print(f"  -> NBFC classified as {c['valuation_family']}, FCFF strictly excluded: {m['excluded_methods']['FCFF_DCF'][:45]}...")

    # -------------------------------------------------------------------------
    # TEST 8: Bank (Commercial Banking / Deposits)
    # -------------------------------------------------------------------------
    def test_08_bank_company(self):
        print("[TEST 8] Commercial Bank Archetype (SBIN / HDFCBANK)")
        data = fetch_company_data('SBIN')
        c = CompanyClassificationEngine.classify(data)
        self.assertEqual(c['valuation_family'], 'BANK')
        self.assertTrue(c['is_financial'])
        m = ValuationMethodSelector.select_methods(c, data)
        self.assertEqual(m['primary_method'], 'EXCESS_RETURN')
        self.assertIn('EV_EBITDA', m['excluded_methods'])
        
        # Test Excess Return DCF
        fn = FinancialNormalizationEngine.normalize(data, c)
        coc = CostOfCapitalEngine.calculate(data, c, fn)
        r = SustainableROICEngine.compute(fn, c, base_growth=12.0, wacc=coc['discount_rate'])
        dcf = DCFEngine.calculate_dcf(fn, c, {}, r, coc)
        self.assertTrue(dcf['valid'])
        self.assertEqual(dcf['valuation_model'], 'EXCESS_RETURN_MODEL')
        print(f"  -> Bank successfully valued via Excess Return Model: Intrinsic Rs. {dcf['intrinsic_value_per_share']}/share")

    # -------------------------------------------------------------------------
    # TEST 9: Insurance Company (Underwriting / Premium Float)
    # -------------------------------------------------------------------------
    def test_09_insurance_company(self):
        print("[TEST 9] Insurance Company Archetype")
        ins_data = {
            'company_name': 'SBI Life Insurance Company Limited',
            'ticker': 'SBILIFE',
            'sector': 'Financial Services',
            'industry': 'Life Insurance',
            'about': 'SBI Life Insurance provides life insurance and general annuity policies across India.',
            'current_price': 1800.0,
            'market_cap': 180000.0,
            'tables': {
                'profit-loss': pd.DataFrame({
                    'Metric': ['Premiums Earned', 'Other Income', 'Profit before tax', 'Net profit'],
                    'Mar 2022': [58000, 2000, 1900, 1500],
                    'Mar 2023': [67000, 2500, 2200, 1700],
                    'Mar 2024': [81000, 3000, 2400, 1900],
                }),
                'balance-sheet': pd.DataFrame({
                    'Metric': ['Equity Capital', 'Reserves', 'Policy Liabilities', 'Total Assets'],
                    'Mar 2022': [1000, 11000, 250000, 270000],
                    'Mar 2023': [1000, 12500, 300000, 320000],
                    'Mar 2024': [1000, 14000, 360000, 385000],
                })
            }
        }
        c = CompanyClassificationEngine.classify(ins_data)
        self.assertEqual(c['valuation_family'], 'INSURANCE')
        self.assertTrue(c['is_financial'])
        m = ValuationMethodSelector.select_methods(c, ins_data)
        self.assertEqual(m['primary_method'], 'EXCESS_RETURN')
        self.assertIn('EV_EBITDA', m['excluded_methods'])
        print(f"  -> Insurance classified as {c['valuation_family']}, EV/EBITDA banned.")

    # -------------------------------------------------------------------------
    # TEST 10: Loss-Making Company (Negative Net Income)
    # -------------------------------------------------------------------------
    def test_10_loss_making_company(self):
        print("[TEST 10] Loss-Making Company (Negative Net Income)")
        loss_data = {
            'company_name': 'Turnaround Industrial Ltd',
            'ticker': 'TURNAROUND',
            'sector': 'Capital Goods',
            'industry': 'Engineering',
            'current_price': 45.0,
            'market_cap': 4500.0,
            'tables': {
                'profit-loss': pd.DataFrame({
                    'Metric': ['Sales', 'Expenses', 'Operating Profit', 'Depreciation', 'Profit before tax', 'Net profit'],
                    'Mar 2022': [2000, 1800, 200, 150, -50, -80],
                    'Mar 2023': [2200, 2000, 200, 160, -20, -50],
                    'Mar 2024': [2500, 2250, 250, 170, -10, -40],
                }),
                'balance-sheet': pd.DataFrame({
                    'Metric': ['Equity Capital', 'Reserves', 'Borrowings', 'Fixed Assets', 'Other Assets'],
                    'Mar 2022': [100, 800, 600, 900, 600],
                    'Mar 2023': [100, 750, 580, 880, 620],
                    'Mar 2024': [100, 710, 550, 860, 650],
                })
            }
        }
        c = CompanyClassificationEngine.classify(loss_data)
        m = ValuationMethodSelector.select_methods(c, loss_data)
        # P/E MUST be excluded for loss-making company
        self.assertIn('P/E', m['excluded_methods'])
        self.assertNotIn('P/E', m['applicable_multiples'])
        print(f"  -> P/E successfully excluded for loss-making company: {m['excluded_methods']['P/E']}")

    # -------------------------------------------------------------------------
    # TEST 11: Highly Leveraged Company (Debt Overhang / Distress)
    # -------------------------------------------------------------------------
    def test_11_highly_leveraged_company(self):
        print("[TEST 11] Highly Leveraged Company (Debt > Enterprise Value)")
        leveraged_data = {
            'company_name': 'Heavy Debt Infrastructure Ltd',
            'ticker': 'HEAVYDEBT',
            'sector': 'Infrastructure',
            'industry': 'Roads & EPC',
            'current_price': 12.0,
            'market_cap': 600.0,
            'tables': {
                'profit-loss': pd.DataFrame({
                    'Metric': ['Sales', 'Expenses', 'Operating Profit', 'Depreciation', 'Interest', 'Profit before tax', 'Net profit'],
                    'Mar 2022': [1500, 1300, 200, 100, 180, -80, -80],
                    'Mar 2023': [1600, 1380, 220, 110, 190, -80, -80],
                    'Mar 2024': [1700, 1450, 250, 120, 200, -70, -70],
                }),
                'balance-sheet': pd.DataFrame({
                    'Metric': ['Equity Capital', 'Reserves', 'Borrowings', 'Fixed Assets', 'Other Assets'],
                    'Mar 2022': [50, 200, 2500, 1800, 1000],
                    'Mar 2023': [50, 120, 2700, 1850, 1050],
                    'Mar 2024': [50, 50, 3000, 1900, 1200],  # 3000 Cr debt vs 600 Cr MCap
                })
            }
        }
        c = CompanyClassificationEngine.classify(leveraged_data)
        fn = FinancialNormalizationEngine.normalize(leveraged_data, c)
        f = ForecastEngine.generate_forecast(fn, c)
        coc = CostOfCapitalEngine.calculate(leveraged_data, c, fn)
        r = SustainableROICEngine.compute(fn, c, base_growth=f['base_growth'], wacc=coc['discount_rate'])
        dcf = DCFEngine.calculate_dcf(fn, c, f, r, coc)
        
        # Must execute without crash and identify capital distress
        self.assertTrue(dcf['valid'])
        self.assertIsNotNone(dcf.get('tv_warning'))
        self.assertIn("CAPITAL DISTRESS", dcf['tv_warning'])
        print(f"  -> Distress warning generated properly: {dcf['tv_warning']}")

    # -------------------------------------------------------------------------
    # TEST 12: Cash-Rich Company (Surplus Cash > Total Debt)
    # -------------------------------------------------------------------------
    def test_12_cash_rich_company(self):
        print("[TEST 12] Cash-Rich Company (Net Cash Surplus Added to Equity)")
        cash_rich_data = {
            'company_name': 'Cash Fortress Software Ltd',
            'ticker': 'CASHFORT',
            'sector': 'Technology',
            'industry': 'Software & SaaS',
            'current_price': 1500.0,
            'market_cap': 150000.0,
            'tables': {
                'profit-loss': pd.DataFrame({
                    'Metric': ['Sales', 'Expenses', 'Operating Profit', 'Depreciation', 'Profit before tax', 'Net profit'],
                    'Mar 2022': [20000, 14000, 6000, 800, 5400, 4100],
                    'Mar 2023': [23000, 16000, 7000, 900, 6300, 4800],
                    'Mar 2024': [26000, 18000, 8000, 1000, 7200, 5500],
                }),
                'balance-sheet': pd.DataFrame({
                    'Metric': ['Equity Capital', 'Reserves', 'Borrowings', 'Investments', 'Other Assets'],
                    'Mar 2022': [100, 30000, 500, 12000, 18000],
                    'Mar 2023': [100, 35000, 400, 15000, 20000],
                    'Mar 2024': [100, 40000, 300, 20000, 25000], # Cash & Inv = ~17,750 Cr vs Borrowings = 300 Cr
                })
            }
        }
        c = CompanyClassificationEngine.classify(cash_rich_data)
        fn = FinancialNormalizationEngine.normalize(cash_rich_data, c)
        net_debt = fn['bs']['net_debt'][-1]
        self.assertLess(net_debt, 0, f"Expected negative Net Debt (Net Cash), got: {net_debt}")
        
        f = ForecastEngine.generate_forecast(fn, c)
        coc = CostOfCapitalEngine.calculate(cash_rich_data, c, fn)
        r = SustainableROICEngine.compute(fn, c, base_growth=f['base_growth'], wacc=coc['discount_rate'])
        dcf = DCFEngine.calculate_dcf(fn, c, f, r, coc)
        
        ev = dcf['enterprise_value_cr']
        eq = dcf['equity_value_cr']
        # Equity value must exceed Enterprise Value because Net Debt is negative (Net Cash)!
        self.assertGreater(eq, ev, f"For cash-rich firm, Equity Value ({eq}) must exceed Enterprise Value ({ev})!")
        print(f"  -> Full Cash credit verified: EV = Rs. {ev:.1f} Cr, Net Debt = Rs. {net_debt:.1f} Cr, Equity Value = Rs. {eq:.1f} Cr (EV + Net Cash)")

    # -------------------------------------------------------------------------
    # TEST 13: Cyclical Company (Mid-Cycle Margin & Normalization)
    # -------------------------------------------------------------------------
    def test_13_cyclical_company(self):
        print("[TEST 13] Cyclical Company Mid-Cycle Smoothing")
        # Volatile margins: 35% -> 8% -> 28% -> 12%
        cyclical_data = {
            'company_name': 'Mining & Minerals Cyclical Ltd',
            'ticker': 'MINECYC',
            'sector': 'Commodities',
            'industry': 'Mining & Metals',
            'current_price': 180.0,
            'market_cap': 36000.0,
            'tables': {
                'profit-loss': pd.DataFrame({
                    'Metric': ['Sales', 'Expenses', 'Operating Profit', 'Depreciation', 'Profit before tax', 'Net profit'],
                    'Mar 2021': [10000, 6500, 3500, 500, 3000, 2200],  # 35% margin (supercycle peak)
                    'Mar 2022': [11000, 10120, 880, 520, 360, 260],    # 8% margin (trough)
                    'Mar 2023': [13000, 9360, 3640, 550, 3090, 2300],  # 28% margin
                    'Mar 2024': [12000, 10560, 1440, 580, 860, 640],   # 12% margin (latest)
                }),
                'balance-sheet': pd.DataFrame({
                    'Metric': ['Equity Capital', 'Reserves', 'Borrowings', 'Fixed Assets', 'Other Assets'],
                    'Mar 2021': [200, 8000, 3000, 6000, 4500],
                    'Mar 2022': [200, 8200, 3500, 6400, 4800],
                    'Mar 2023': [200, 10400, 2800, 7000, 5200],
                    'Mar 2024': [200, 11000, 2600, 7200, 5500],
                })
            }
        }
        c = CompanyClassificationEngine.classify(cyclical_data)
        fn = FinancialNormalizationEngine.normalize(cyclical_data, c)
        norm_margin = fn['normalized_metrics']['normalized_ebitda_margin']
        latest_margin = fn['normalized_metrics']['latest_ebitda_margin']
        # Normalized margin should be blended (approx 20%), NOT just the latest 12%
        self.assertNotEqual(norm_margin, latest_margin)
        self.assertTrue(fn['normalized_metrics']['normalization_applied'])
        print(f"  -> Mid-cycle normalization applied: Latest margin = {latest_margin}%, Mid-cycle normalized = {norm_margin}%")

    # -------------------------------------------------------------------------
    # TEST 14: Company with Incomplete Data
    # -------------------------------------------------------------------------
    def test_14_incomplete_data(self):
        print("[TEST 14] Company with Incomplete Data (Missing CF and Balance Sheet items)")
        incomplete_data = {
            'company_name': 'Sparse Disclosures Ltd',
            'ticker': 'SPARSE',
            'current_price': 100.0,
            'market_cap': 1000.0,
            'tables': {
                'profit-loss': pd.DataFrame({
                    'Metric': ['Sales', 'Expenses', 'Operating Profit', 'Net profit'],
                    'Mar 2022': [500, 400, 100, 60],
                    'Mar 2023': [600, 480, 120, 75],
                    'Mar 2024': [700, 550, 150, 95],
                }),
                'balance-sheet': pd.DataFrame() # Missing balance sheet
            }
        }
        c = CompanyClassificationEngine.classify(incomplete_data)
        fn = FinancialNormalizationEngine.normalize(incomplete_data, c)
        f = ForecastEngine.generate_forecast(fn, c)
        coc = CostOfCapitalEngine.calculate(incomplete_data, c, fn)
        r = SustainableROICEngine.compute(fn, c, base_growth=f['base_growth'], wacc=coc['discount_rate'])
        dcf = DCFEngine.calculate_dcf(fn, c, f, r, coc)
        self.assertTrue(dcf['valid'])
        self.assertGreater(dcf['intrinsic_value_per_share'], 0.0)
        print(f"  -> Robust fallback executed: Intrinsic Rs. {dcf['intrinsic_value_per_share']}/share without crash.")

    # -------------------------------------------------------------------------
    # TEST 15: Company with Only 3 Years of History
    # -------------------------------------------------------------------------
    def test_15_short_history_3yr(self):
        print("[TEST 15] Company with Only 3 Years of History (Recent IPO)")
        ipo_data = {
            'company_name': 'New Tech IPO Ltd',
            'ticker': 'NEWTECH',
            'sector': 'Technology',
            'industry': 'Software',
            'current_price': 400.0,
            'market_cap': 4000.0,
            'tables': {
                'profit-loss': pd.DataFrame({
                    'Metric': ['Sales', 'Expenses', 'Operating Profit', 'Depreciation', 'Profit before tax', 'Net profit'],
                    'Mar 2022': [300, 240, 60, 10, 50, 38],
                    'Mar 2023': [450, 350, 100, 15, 85, 64],
                    'Mar 2024': [650, 480, 170, 20, 150, 112],
                }),
                'balance-sheet': pd.DataFrame({
                    'Metric': ['Equity Capital', 'Reserves', 'Borrowings', 'Fixed Assets', 'Other Assets'],
                    'Mar 2022': [10, 200, 20, 50, 180],
                    'Mar 2023': [10, 260, 15, 70, 215],
                    'Mar 2024': [10, 370, 10, 90, 300],
                })
            }
        }
        c = CompanyClassificationEngine.classify(ipo_data)
        fn = FinancialNormalizationEngine.normalize(ipo_data, c)
        self.assertEqual(fn['num_years'], 3)
        self.assertIn("WARNING: Historical series is limited to 3 years", fn['data_quality_flags'][0])
        f = ForecastEngine.generate_forecast(fn, c)
        # CAGR calculation must not raise IndexError
        self.assertIsNotNone(f['base_growth'])
        print(f"  -> 3-Year IPO correctly handled: 2Y/3Y CAGR = {f['base_growth']}%, Quality status: {fn['data_quality_status']}")

    # -------------------------------------------------------------------------
    # TEST 16: Company with Negative EBITDA
    # -------------------------------------------------------------------------
    def test_16_negative_ebitda_company(self):
        print("[TEST 16] Company with Negative EBITDA (Multiple Suppression & Scenario Safety)")
        neg_ebitda_data = {
            'company_name': 'Distressed Steel Producer Ltd',
            'ticker': 'DISTSTEEL',
            'sector': 'Metals',
            'industry': 'Steel',
            'current_price': 15.0,
            'market_cap': 750.0,
            'tables': {
                'profit-loss': pd.DataFrame({
                    'Metric': ['Sales', 'Expenses', 'Operating Profit', 'Depreciation', 'Profit before tax', 'Net profit'],
                    'Mar 2022': [1000, 1150, -150, 80, -230, -230],
                    'Mar 2023': [1100, 1280, -180, 85, -265, -265],
                    'Mar 2024': [1200, 1400, -200, 90, -290, -290],
                }),
                'balance-sheet': pd.DataFrame({
                    'Metric': ['Equity Capital', 'Reserves', 'Borrowings', 'Fixed Assets', 'Other Assets'],
                    'Mar 2022': [50, 400, 600, 700, 350],
                    'Mar 2023': [50, 200, 750, 680, 320],
                    'Mar 2024': [50, 50, 900, 650, 350],
                })
            }
        }
        c = CompanyClassificationEngine.classify(neg_ebitda_data)
        m = ValuationMethodSelector.select_methods(c, neg_ebitda_data)
        # EV/EBITDA and P/E MUST be suppressed
        self.assertIn('EV/EBITDA', m['excluded_methods'])
        self.assertIn('P/E', m['excluded_methods'])
        
        fn = FinancialNormalizationEngine.normalize(neg_ebitda_data, c)
        f = ForecastEngine.generate_forecast(fn, c)
        # Verify Bear EBITDA margin does NOT invert into positive
        bear_margin = f['scenarios']['Bear']['ebitda_margin_pct'][0]
        base_margin = f['scenarios']['Base']['ebitda_margin_pct'][0]
        self.assertLess(bear_margin, base_margin, f"Bear margin ({bear_margin}%) must be lower than Base ({base_margin}%)!")
        print(f"  -> Negative EBITDA handled: EV/EBITDA excluded; Base margin = {base_margin}%, Bear margin = {bear_margin}% (no inversion).")

    # -------------------------------------------------------------------------
    # TEST 17: Company with Negative Net Income
    # -------------------------------------------------------------------------
    def test_17_negative_net_income(self):
        print("[TEST 17] Company with Negative Net Income (P/E Suppression)")
        neg_ni_data = {
            'company_name': 'Clean Energy Growth Ltd',
            'ticker': 'CLEANGROWTH',
            'sector': 'Renewables',
            'industry': 'Solar & Wind',
            'current_price': 50.0,
            'market_cap': 5000.0,
            'tables': {
                'profit-loss': pd.DataFrame({
                    'Metric': ['Sales', 'Expenses', 'Operating Profit', 'Depreciation', 'Interest', 'Profit before tax', 'Net profit'],
                    'Mar 2022': [1000, 750, 250, 150, 150, -50, -50],
                    'Mar 2023': [1400, 1000, 400, 220, 220, -40, -40],
                    'Mar 2024': [1900, 1350, 550, 300, 300, -50, -50],
                }),
                'balance-sheet': pd.DataFrame({
                    'Metric': ['Equity Capital', 'Reserves', 'Borrowings', 'Fixed Assets', 'Other Assets'],
                    'Mar 2022': [100, 1000, 1500, 2100, 500],
                    'Mar 2023': [100, 950, 2200, 2800, 450],
                    'Mar 2024': [100, 900, 3000, 3500, 500],
                })
            }
        }
        c = CompanyClassificationEngine.classify(neg_ni_data)
        m = ValuationMethodSelector.select_methods(c, neg_ni_data)
        self.assertIn('P/E', m['excluded_methods'])
        self.assertIn('EV/EBITDA', m['applicable_multiples'])
        print(f"  -> P/E excluded due to negative earnings; EV/EBITDA preserved (operating profit is positive Rs. 550 Cr).")

    # -------------------------------------------------------------------------
    # TEST 18: Distressed Bank with Negative Book Value
    # -------------------------------------------------------------------------
    def test_18_negative_book_value_bank(self):
        print("[TEST 18] Distressed Bank with Negative Book Value (Regulatory Insolvency)")
        distressed_bank = {
            'company_name': 'Insolvent Cooperative Bank Ltd',
            'ticker': 'INSOLVENTBK',
            'sector': 'Banking',
            'industry': 'Commercial Banks',
            'about': 'Insolvent Bank has experienced massive NPA write-downs exhausting regulatory capital.',
            'current_price': 5.0,
            'market_cap': 250.0,
            'tables': {
                'profit-loss': pd.DataFrame({
                    'Metric': ['Interest Earned', 'Financing Profit', 'Profit before tax', 'Net profit'],
                    'Mar 2022': [5000, -500, -800, -800],
                    'Mar 2023': [4500, -900, -1200, -1200],
                    'Mar 2024': [4000, -1500, -2000, -2000],
                }),
                'balance-sheet': pd.DataFrame({
                    'Metric': ['Equity Capital', 'Reserves', 'Deposits', 'Advances'],
                    'Mar 2022': [50, -400, 35000, 28000],
                    'Mar 2023': [50, -1600, 32000, 24000],
                    'Mar 2024': [50, -3600, 28000, 18000], # Total Equity = -3550 Cr
                })
            }
        }
        c = CompanyClassificationEngine.classify(distressed_bank)
        fn = FinancialNormalizationEngine.normalize(distressed_bank, c)
        coc = CostOfCapitalEngine.calculate(distressed_bank, c, fn)
        r = SustainableROICEngine.compute(fn, c, base_growth=5.0, wacc=coc['discount_rate'])
        dcf = DCFEngine.calculate_dcf(fn, c, {}, r, coc)
        
        # Must safely abort with capital deficit warning, zero equity value, no crash
        self.assertFalse(dcf['valid'])
        self.assertEqual(dcf['intrinsic_value_per_share'], 0.0)
        self.assertIn("Regulatory Capital Deficit", dcf['error'])
        print(f"  -> Regulatory capital deficit detected properly: {dcf['error']}")

    # -------------------------------------------------------------------------
    # TEST 19: Adversarial Boundary: WACC <= Terminal Growth
    # -------------------------------------------------------------------------
    def test_19_wacc_terminal_growth_singularity(self):
        print("[TEST 19] Boundary Stress: WACC <= Terminal Growth (Gordon Denominator <= 0)")
        data = fetch_company_data('TATASTEEL')
        c = CompanyClassificationEngine.classify(data)
        fn = FinancialNormalizationEngine.normalize(data, c)
        f = ForecastEngine.generate_forecast(fn, c)
        # Force WACC = 5.0% and Terminal Growth = 5.0%
        bad_coc = {'wacc': 5.0, 'discount_rate': 5.0, 'is_financial': False, 'wacc_valid': False}
        r = SustainableROICEngine.compute(fn, c, base_growth=f['base_growth'], wacc=5.0, terminal_growth=5.0)
        
        dcf = DCFEngine.calculate_dcf(fn, c, f, r, bad_coc, terminal_growth=5.0)
        self.assertFalse(dcf['valid'])
        self.assertIn("WACC (5.00%) <= Terminal Growth (5.00%)", dcf['error'])
        print(f"  -> Model safely aborted without division by zero: {dcf['error']}")

    # -------------------------------------------------------------------------
    # TEST 20: Self-Peer Contamination Variant Tests
    # -------------------------------------------------------------------------
    def test_20_self_peer_contamination_variants(self):
        print("[TEST 20] Adversarial Self-Peer Variants (Ltd, Limited, Punctuation, Acronyms)")
        
        screener_data = {
            'company_name': 'Tata Steel Ltd.',
            'ticker': 'TATASTEEL',
            'peers': [
                {'Name': 'Tata Steel Limited', 'ticker': '500470', 'P/E': 14.2, 'EV / EBITDA': 7.5}, # Must be excluded!
                {'Name': 'JSW Steel Limited', 'ticker': 'JSWSTEEL', 'P/E': 15.5, 'EV / EBITDA': 8.2},
                {'Name': 'Jindal Steel & Power Ltd', 'ticker': 'JINDALSTEL', 'P/E': 12.0, 'EV / EBITDA': 6.8},
                {'Name': 'Steel Authority of India Ltd', 'ticker': 'SAIL', 'P/E': 10.5, 'EV / EBITDA': 5.5},
            ]
        }
        c = CompanyClassificationEngine.classify(screener_data)
        fn = FinancialNormalizationEngine.normalize(screener_data, c)
        m = ValuationMethodSelector.select_methods(c, screener_data)
        
        peers_result = PeerSelectionEngine.process_peers(screener_data, c, fn, m)
        peer_names = [p['name'].lower() for p in peers_result['peers']]
        peer_tickers = [p['ticker'].upper() for p in peers_result['peers']]
        
        # Verify Tata Steel Limited was strictly excluded
        self.assertNotIn('tata steel limited', peer_names)
        self.assertNotIn('500470', peer_tickers)
        self.assertIn('JSWSTEEL', peer_tickers)
        print(f"  -> Self-peer 'Tata Steel Limited' (500470) strictly excluded from 'Tata Steel Ltd.' (TATASTEEL). Peer count: {peers_result['peer_count']}")


if __name__ == '__main__':
    unittest.main()
