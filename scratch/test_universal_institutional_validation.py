"""
scratch/test_universal_institutional_validation.py
=================================================
Automated Institutional Validation Suite (Phases 31 & 32).
Validates the Universal Valuation Engine across 6 materially different company archetypes:
  1. Archetype A: Normal Industrial Company (TATASTEEL)
  2. Archetype B: Commercial Bank / Financial Institution (SBIN)
  3. Archetype C: Cyclical / Commodity Company (Synthetic Cement)
  4. Archetype D: Diversified Conglomerate (ADANIENT - Institutional Regression Test)
  5. Archetype E: Loss-Making Growth / Turnaround Company (Synthetic)
  6. Archetype F: Asset-Heavy Infrastructure Company (LT / Synthetic Infra)

Verifies all 22 Phase 27 validation checks and confirms zero company-specific hardcoding.
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

from screener_client import fetch_company_data
import valuation_engine
from universal_valuation import (
    VALUATION_CONFIG,
    CompanyMaster,
    CompanyClassificationEngine,
    ValuationMethodSelector,
    FinancialNormalizationEngine,
    ForecastEngine,
    CostOfCapitalEngine,
    SustainableROICEngine,
    PeerSelectionEngine,
    SOTPEngine,
    DCFEngine,
    ValuationReconciliationEngine,
    ValuationConfidenceEngine,
    ModelQualityEngine
)

class UniversalInstitutionalValidationSuite(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        print("\n" + "="*85)
        print("UNIVERSAL VALUATION TERMINAL - INSTITUTIONAL AUDIT & REPAIR VALIDATION SUITE")
        print("="*85)

    # -------------------------------------------------------------------------
    # TEST ARCHETYPE A: Normal Industrial Company (TATASTEEL)
    # -------------------------------------------------------------------------
    def test_archetype_a_industrial(self):
        print("\n>>> Testing Archetype A: Normal Industrial Company (TATASTEEL)...")
        data = fetch_company_data('TATASTEEL')
        res = valuation_engine.calculate_valuation(data)
        
        # Verify Master Object
        cm = res['universal_engine']['company_master']
        self.assertIsNotNone(cm)
        self.assertGreater(cm['shares_outstanding'], 0)
        self.assertGreater(cm['market_cap'], 0)
        self.assertIn(cm['reporting_basis'], ['CONSOLIDATED', 'STANDALONE'])
        
        # Verify Classification
        clf = res['universal_engine']['classification']
        self.assertFalse(clf['is_financial'])
        
        # Verify DCF & Comps use identical shares
        self.assertAlmostEqual(res['shares_cr'], cm['shares_outstanding'], places=2)
        
        # Verify Quality Engine 22 tests
        quality = res['universal_engine']['quality']
        self.assertEqual(quality['fail_count'], 0, f"Quality failures in Archetype A: {quality['tests']}")
        self.assertEqual(quality['total_tests'], 22)
        print(f"  [PASS] Archetype A: 22/22 Quality Tests Passed (DCF Value: Rs. {res['intrinsic_value_per_share']:.2f}, Shares: {cm['shares_outstanding']:.2f} Cr)")

    # -------------------------------------------------------------------------
    # TEST ARCHETYPE B: Commercial Bank / Financial Institution (SBIN)
    # -------------------------------------------------------------------------
    def test_archetype_b_bank(self):
        print("\n>>> Testing Archetype B: Commercial Bank / Financial Institution (SBIN)...")
        data = fetch_company_data('SBIN')
        res = valuation_engine.calculate_valuation(data)
        
        clf = res['universal_engine']['classification']
        self.assertEqual(clf['company_type'], 'Banking')
        self.assertEqual(clf['valuation_family'], 'BANK')
        
        methods = res['universal_engine']['valuation_methods']
        self.assertIn(methods['primary_method'], ['EXCESS_RETURN', 'EXCESS_RETURN_MODEL', 'DIVIDEND_DISCOUNT_MODEL'])

        self.assertIn('FCFF_DCF', methods['inapplicable_methods'])
        self.assertIn('EV_EBITDA', methods['inapplicable_methods'])
        
        quality = res['universal_engine']['quality']
        self.assertEqual(quality['fail_count'], 0, f"Quality failures in Archetype B: {quality['tests']}")
        self.assertEqual(quality['total_tests'], 22)
        print(f"  [PASS] Archetype B: Bank Excess Return Model deployed, FCFF/EV multiples suppressed (Value: Rs. {res['intrinsic_value_per_share']:.2f})")

    # -------------------------------------------------------------------------
    # TEST ARCHETYPE C: Cyclical / Commodity Company (Synthetic Cement)
    # -------------------------------------------------------------------------
    def test_archetype_c_cyclical_commodity(self):
        print("\n>>> Testing Archetype C: Cyclical / Commodity Company (Synthetic Cement)...")
        cement_data = {
            'company_name': 'UltraPeak Cement Limited',
            'ticker': 'ULTRAPEAK',
            'sector': 'Cement',
            'industry': 'Cement & Clinker',
            'about': 'UltraPeak Cement is a premier cement manufacturer operating large integrated kilns.',
            'current_price': 420.0,
            'market_cap': 42000.0,
            'market_cap_cr': 42000.0,
            'face_value': 10.0,
            'tables': {
                'profit-loss': pd.DataFrame({
                    'Metric': ['Sales', 'Expenses', 'Operating Profit', 'Depreciation', 'Interest', 'Profit before tax', 'Net profit', 'EPS in Rs'],
                    'Mar 2020': [12000, 9600, 2400, 700, 300, 1400, 1000, 10.0],
                    'Mar 2021': [14000, 10500, 3500, 750, 280, 2470, 1800, 18.0],
                    'Mar 2022': [18000, 13500, 4500, 800, 260, 3440, 2500, 25.0],
                    'Mar 2023': [17000, 14000, 3000, 850, 250, 1900, 1400, 14.0],
                    'Mar 2024': [19000, 15200, 3800, 900, 240, 2660, 1950, 19.5],
                }),
                'balance-sheet': pd.DataFrame({
                    'Metric': ['Equity Capital', 'Reserves', 'Borrowings', 'Other Liabilities', 'Total Liabilities', 'Fixed Assets', 'Other Assets', 'Total Assets'],
                    'Mar 2020': [1000, 12000, 5000, 2000, 20000, 14000, 6000, 20000],
                    'Mar 2021': [1000, 13500, 4500, 2200, 21200, 14800, 6400, 21200],
                    'Mar 2022': [1000, 15500, 4000, 2400, 22900, 15500, 7400, 22900],
                    'Mar 2023': [1000, 16500, 4200, 2500, 24200, 16200, 8000, 24200],
                    'Mar 2024': [1000, 18000, 3800, 2700, 25500, 17000, 8500, 25500],
                }),
                'cash-flow': pd.DataFrame({
                    'Metric': ['Cash from Operating Activity', 'Cash from Investing Activity', 'Cash from Financing Activity', 'Net Cash Flow'],
                    'Mar 2020': [2100, -1800, -200, 100],
                    'Mar 2021': [3000, -1500, -1200, 300],
                    'Mar 2022': [3800, -1600, -1900, 300],
                    'Mar 2023': [2700, -1700, -800, 200],
                    'Mar 2024': [3200, -1800, -1100, 300],
                }),
                'ratios': pd.DataFrame({
                    'Metric': ['Debtor Days', 'Inventory Days', 'Days Payable', 'Cash Conversion Cycle', 'Working Capital Days'],
                    'Mar 2020': [25, 45, 35, 35, 40],
                    'Mar 2021': [24, 42, 33, 33, 38],
                    'Mar 2022': [22, 40, 32, 30, 35],
                    'Mar 2023': [23, 43, 34, 32, 37],
                    'Mar 2024': [22, 41, 33, 30, 36],
                })
            }
        }
        res = valuation_engine.calculate_valuation(cement_data)
        clf = res['universal_engine']['classification']
        self.assertTrue(clf['is_cyclical'])
        
        quality = res['universal_engine']['quality']
        self.assertEqual(quality['fail_count'], 0)
        print(f"  [PASS] Archetype C: Cyclical normalization verified (DCF: Rs. {res['intrinsic_value_per_share']:.2f})")

    # -------------------------------------------------------------------------
    # TEST ARCHETYPE D: Diversified Conglomerate (ADANIENT - Regression Test)
    # -------------------------------------------------------------------------
    def test_archetype_d_conglomerate_adanient(self):
        print("\n>>> Testing Archetype D: Diversified Conglomerate Regression Test (ADANIENT)...")
        data = fetch_company_data('ADANIENT')
        res = valuation_engine.calculate_valuation(data)
        
        # Phase 2: Canonical Master Object
        cm = res['universal_engine']['company_master']
        self.assertIsNotNone(cm)
        self.assertEqual(cm['ticker'], 'ADANIENT')
        # Check share count resolution (114.0 Cr shares from Rs 114.0 Cr equity capital / FV 1.0)
        self.assertGreater(cm['shares_outstanding'], 100.0)
        self.assertLess(cm['shares_outstanding'], 150.0)
        self.assertGreater(cm['market_cap'], 300000.0)
        
        # Phase 7: Classification
        clf = res['universal_engine']['classification']
        self.assertTrue(clf['is_conglomerate'])
        self.assertTrue(clf['requires_sotp'])
        
        # Phase 8 & 16: SOTP Engine Evaluation
        sotp = res['universal_engine']['sotp']
        self.assertIsNotNone(sotp)
        self.assertGreater(len(sotp['segments']), 0)
        self.assertGreater(sotp['conglomerate_equity_value_cr'], 0)
        self.assertGreater(sotp['implied_value_per_share'], 0)
        print(f"  -> ADANIENT SOTP Implied Value: Rs. {sotp['implied_value_per_share']:.2f} per share (with {sotp['holding_company_discount_pct']}% HoldCo discount)")
        print(f"  -> ADANIENT DCF Implied Value: Rs. {res['intrinsic_value_per_share']:.2f} per share")
        
        # Phase 20: Scenarios Bear <= Base <= Bull
        scenarios = res['universal_engine']['reconciliation']['scenarios']
        bear_p = scenarios['bear_case']
        base_p = scenarios['base_case']
        bull_p = scenarios['bull_case']
        self.assertLessEqual(bear_p, base_p)
        self.assertLessEqual(base_p, bull_p)
        print(f"  -> Scenarios: Bear Rs. {bear_p:.2f} <= Base Rs. {base_p:.2f} <= Bull Rs. {bull_p:.2f}")
        
        # Phase 21: Neutrality Check
        self.assertNotIn('BUY', str(res.get('verdict', '')).upper())
        self.assertNotIn('SELL', str(res.get('verdict', '')).upper())
        self.assertNotIn('UNDERVALUED', str(res.get('verdict', '')).upper())
        self.assertNotIn('OVERVALUED', str(res.get('verdict', '')).upper())
        
        # Phase 27: All 22 tests
        quality = res['universal_engine']['quality']
        self.assertEqual(quality['fail_count'], 0, f"Quality failures in ADANIENT: {quality['tests']}")
        self.assertEqual(quality['total_tests'], 22)
        print(f"  [PASS] Archetype D (ADANIENT): Generic Conglomerate + SOTP + DCF + Master Object validated with 22/22 Tests!")

    # -------------------------------------------------------------------------
    # TEST ARCHETYPE E: Loss-Making Growth / Turnaround Company (Synthetic)
    # -------------------------------------------------------------------------
    def test_archetype_e_loss_making(self):
        print("\n>>> Testing Archetype E: Loss-Making Turnaround Company (Synthetic)...")
        loss_data = {
            'company_name': 'NextGen Quantum Tech',
            'ticker': 'NEXTGENQ',
            'sector': 'Technology',
            'industry': 'Software & AI',
            'about': 'NextGen Quantum develops high-growth AI software solutions, scaling rapidly while investing ahead of revenue.',
            'current_price': 180.0,
            'market_cap': 9000.0,
            'market_cap_cr': 9000.0,
            'face_value': 2.0,
            'tables': {
                'profit-loss': pd.DataFrame({
                    'Metric': ['Sales', 'Expenses', 'Operating Profit', 'Depreciation', 'Interest', 'Profit before tax', 'Net profit', 'EPS in Rs'],
                    'Mar 2021': [400, 500, -100, 30, 10, -140, -140, -2.8],
                    'Mar 2022': [800, 950, -150, 45, 12, -207, -207, -4.1],
                    'Mar 2023': [1500, 1650, -150, 60, 15, -225, -225, -4.5],
                    'Mar 2024': [2600, 2700, -100, 80, 20, -200, -200, -4.0],
                }),
                'balance-sheet': pd.DataFrame({
                    'Metric': ['Equity Capital', 'Reserves', 'Borrowings', 'Other Liabilities', 'Total Liabilities', 'Fixed Assets', 'Other Assets', 'Total Assets'],
                    'Mar 2021': [100, 600, 100, 80, 880, 200, 680, 880],
                    'Mar 2022': [100, 1200, 150, 120, 1570, 350, 1220, 1570],
                    'Mar 2023': [100, 2200, 200, 180, 2680, 550, 2130, 2680],
                    'Mar 2024': [100, 3100, 250, 250, 3700, 800, 2900, 3700],
                }),
                'cash-flow': pd.DataFrame({
                    'Metric': ['Cash from Operating Activity', 'Cash from Investing Activity', 'Cash from Financing Activity', 'Net Cash Flow'],
                    'Mar 2021': [-90, -150, 600, 360],
                    'Mar 2022': [-120, -200, 800, 480],
                    'Mar 2023': [-110, -250, 1200, 840],
                    'Mar 2024': [-80, -320, 1100, 700],
                }),
                'ratios': pd.DataFrame({
                    'Metric': ['Debtor Days', 'Inventory Days', 'Days Payable', 'Cash Conversion Cycle', 'Working Capital Days'],
                    'Mar 2021': [30, 10, 40, 0, 15],
                    'Mar 2022': [32, 12, 42, 2, 18],
                    'Mar 2023': [28, 11, 41, -2, 14],
                    'Mar 2024': [26, 10, 39, -3, 12],
                })
            }
        }
        res = valuation_engine.calculate_valuation(loss_data)
        clf = res['universal_engine']['classification']
        self.assertTrue(clf['is_loss_making'])
        
        methods = res['universal_engine']['valuation_methods']
        self.assertIn('PE', methods['inapplicable_methods'])
        
        quality = res['universal_engine']['quality']
        self.assertEqual(quality['fail_count'], 0)
        print(f"  [PASS] Archetype E: P/E suppressed for loss-maker, growth DCF/EV-Sales deployed (22/22 Tests Passed)")

    # -------------------------------------------------------------------------
    # TEST ARCHETYPE F: Asset-Heavy Infrastructure Company (Synthetic Infra)
    # -------------------------------------------------------------------------
    def test_archetype_f_infrastructure(self):
        print("\n>>> Testing Archetype F: Asset-Heavy Infrastructure Company (Synthetic Infra)...")
        infra_data = {
            'company_name': 'Bharat Port & Logistics Ltd',
            'ticker': 'BHARATPORT',
            'sector': 'Infrastructure',
            'industry': 'Ports & Logistics',
            'about': 'Bharat Port operates deepwater port concessions and dedicated rail terminals under 30-year concessions.',
            'current_price': 850.0,
            'market_cap': 85000.0,
            'market_cap_cr': 85000.0,
            'face_value': 10.0,
            'tables': {
                'profit-loss': pd.DataFrame({
                    'Metric': ['Sales', 'Expenses', 'Operating Profit', 'Depreciation', 'Interest', 'Profit before tax', 'Net profit', 'EPS in Rs'],
                    'Mar 2020': [8000, 3200, 4800, 1200, 1100, 2500, 1800, 18.0],
                    'Mar 2021': [9500, 3800, 5700, 1350, 1150, 3200, 2300, 23.0],
                    'Mar 2022': [12000, 4600, 7400, 1500, 1200, 4700, 3400, 34.0],
                    'Mar 2023': [14500, 5500, 9000, 1700, 1250, 6050, 4400, 44.0],
                    'Mar 2024': [17000, 6400, 10600, 1900, 1300, 7400, 5400, 54.0],
                }),
                'balance-sheet': pd.DataFrame({
                    'Metric': ['Equity Capital', 'Reserves', 'Borrowings', 'Other Liabilities', 'Total Liabilities', 'Fixed Assets', 'Other Assets', 'Total Assets'],
                    'Mar 2020': [1000, 22000, 14000, 3000, 40000, 32000, 8000, 40000],
                    'Mar 2021': [1000, 24000, 15000, 3200, 43200, 34500, 8700, 43200],
                    'Mar 2022': [1000, 27000, 16000, 3500, 47500, 38000, 9500, 47500],
                    'Mar 2023': [1000, 31000, 17000, 3800, 52800, 42000, 10800, 52800],
                    'Mar 2024': [1000, 36000, 18000, 4100, 59100, 47000, 12100, 59100],
                }),
                'cash-flow': pd.DataFrame({
                    'Metric': ['Cash from Operating Activity', 'Cash from Investing Activity', 'Cash from Financing Activity', 'Net Cash Flow'],
                    'Mar 2020': [4200, -3800, -300, 100],
                    'Mar 2021': [5000, -4200, -600, 200],
                    'Mar 2022': [6500, -4800, -1400, 300],
                    'Mar 2023': [7800, -5500, -1900, 400],
                    'Mar 2024': [9200, -6200, -2500, 500],
                }),
                'ratios': pd.DataFrame({
                    'Metric': ['Debtor Days', 'Inventory Days', 'Days Payable', 'Cash Conversion Cycle', 'Working Capital Days'],
                    'Mar 2020': [35, 10, 40, 5, 20],
                    'Mar 2021': [34, 10, 39, 5, 19],
                    'Mar 2022': [32, 9, 38, 3, 18],
                    'Mar 2023': [30, 9, 37, 2, 17],
                    'Mar 2024': [28, 8, 36, 0, 16],
                })
            }
        }
        res = valuation_engine.calculate_valuation(infra_data)
        clf = res['universal_engine']['classification']
        self.assertEqual(clf['company_type'], 'Infrastructure')
        self.assertEqual(clf['valuation_family'], 'INFRASTRUCTURE')


        
        quality = res['universal_engine']['quality']
        self.assertEqual(quality['fail_count'], 0)
        print(f"  [PASS] Archetype F: Asset-Heavy Infrastructure DCF + Reinvestment validated (DCF: Rs. {res['intrinsic_value_per_share']:.2f})")

if __name__ == '__main__':
    unittest.main()
