"""
scratch/test_adversarial_regression_audit_fixes.py
==================================================
Regression tests specifically guarding against the 5 architectural bugs discovered
during the Adversarial Regression Audit:
1. NBFC misclassification as industrial (BAJAJHFL / structural balance sheet test)
2. Column pollution causing false loss-making classification
3. Disconnected legacy DCF overwriting universal DCF engine in valuation_engine.py
4. Hardcoded company name 'adani' in sotp_engine.py
5. Missing root 'beta' key in CostOfCapitalEngine return dictionary
"""

import os
import sys
import unittest
import pandas as pd
import re

ws_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ws_root not in sys.path:
    sys.path.insert(0, ws_root)

from screener_client import fetch_company_data
from universal_valuation import (
    CompanyClassificationEngine,
    CostOfCapitalEngine,
    FinancialNormalizationEngine,
    DCFEngine
)
import valuation_engine

class AdversarialRegressionFixesSuite(unittest.TestCase):

    def test_fix_01_nbfc_structural_classification(self):
        """Bug 1: NBFC must classify as NBFC / EXCESS_RETURN, never ASSET_HEAVY_INDUSTRIAL."""
        data = fetch_company_data('BAJAJHFL')
        c = CompanyClassificationEngine.classify(data)
        self.assertEqual(c['valuation_family'], 'NBFC')
        self.assertTrue(c['is_financial'])
        self.assertFalse(c['is_cyclical'])
        
        # Run valuation engine
        res = valuation_engine.calculate_valuation(data)
        self.assertEqual(res['company_classification']['valuation_family'], 'NBFC')
        dcf = res['universal_engine']['dcf']
        self.assertEqual(dcf['valuation_model'], 'EXCESS_RETURN_MODEL')
        self.assertGreater(res['intrinsic_value_per_share'], 0.0)

    def test_fix_02_column_pollution_loss_making_immunity(self):
        """Bug 2: Non-period columns (Trailing, Best Case, Worst Case) must not trigger is_loss_making."""
        polluted_pl = pd.DataFrame({
            'Metric': ['Sales', 'Expenses', 'Operating Profit', 'Depreciation', 'Interest', 'Profit before tax', 'Net profit'],
            'Mar 2021': [1000, 700, 300, 30, 10, 260, 195],
            'Mar 2022': [1200, 820, 380, 35, 12, 333, 250],
            'Mar 2023': [1500, 1000, 500, 40, 15, 445, 334],
            'Mar 2024': [1800, 1180, 620, 45, 18, 557, 418],
            'Trailing': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
            'Best Case': [None, None, None, None, None, None, None],
            'Worst Case': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
        })
        test_data = {
            'company_name': 'Profitable Corp Ltd',
            'ticker': 'PROFITCORP',
            'sector': 'Technology',
            'tables': {
                'profit-loss': polluted_pl
            }
        }
        c = CompanyClassificationEngine.classify(test_data)
        self.assertFalse(c['is_loss_making'], "Profitable company was falsely marked loss-making due to column pollution!")

    def test_fix_03_root_dcf_synchronization_with_universal_engine(self):
        """Bug 3: valuation_engine.py root intrinsic value must equal univ_dcf intrinsic value."""
        data = fetch_company_data('LICI')
        res = valuation_engine.calculate_valuation(data)
        univ_dcf = res['universal_engine']['dcf']
        
        self.assertEqual(univ_dcf['valuation_model'], 'EXCESS_RETURN_MODEL')
        self.assertGreater(univ_dcf['intrinsic_value_per_share'], 0.0)
        # Verify single source of truth: root must match univ_dcf, not legacy parallel loop
        self.assertEqual(res['intrinsic_value_per_share'], univ_dcf['intrinsic_value_per_share'])
        self.assertGreater(res['intrinsic_value_per_share'], 0.0)

    def test_fix_04_zero_company_hardcoding_in_universal_valuation(self):
        """Bug 4: Assert zero occurrences of 'adani' in universal_valuation python code."""
        univ_dir = os.path.join(ws_root, 'universal_valuation')
        for root, _, files in os.walk(univ_dir):
            for f in files:
                if f.endswith('.py'):
                    f_path = os.path.join(root, f)
                    with open(f_path, 'r', encoding='utf-8') as fh:
                        content = fh.read()
                    matches = re.findall(r'\badani\b', content, re.I)
                    self.assertEqual(len(matches), 0, f"Found hardcoded company name 'adani' in {f_path}")

    def test_fix_05_cost_of_capital_beta_contract(self):
        """Bug 5: CostOfCapitalEngine return dict must contain valid root 'beta' key."""
        data = fetch_company_data('INFY')
        c = CompanyClassificationEngine.classify(data)
        fn = FinancialNormalizationEngine.normalize(data, c)
        coc = CostOfCapitalEngine.calculate(data, c, fn)
        
        self.assertIn('beta', coc)
        self.assertIsNotNone(coc['beta'])
        self.assertGreater(coc['beta'], 0.0)
        self.assertEqual(coc['beta'], coc['selected_beta'])

if __name__ == '__main__':
    unittest.main()
