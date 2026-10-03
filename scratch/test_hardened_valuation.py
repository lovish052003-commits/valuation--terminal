"""
test_hardened_valuation.py
Comprehensive regression and multi-company test suite for the hardened financial valuation engine.

Covers:
1. TEST 1: AI Summary Column Mapping Regression Test (B = LABEL, C = VALUE)
2. TEST 2: Peer Exclusion Regression Test (Strictly exclude target from peer group)
3. TEST 3: Multiple Formulas Dynamic Row Generation Regression Test (IFERROR & Row Reference)
4. TEST 4: End-to-End Multi-Company Test:
   - Wipro (IT)
   - Ambuja Cements (Cement / Materials)
   - Tata Steel (Industrial / Materials)
   - Nestle India (FMCG)
   - Suzlon Energy (Renewable / Capital Goods)
   - HDFC Bank (Banking / Financial Institution)
"""

import os
import sys
import unittest
import openpyxl

# Add parent directory to path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from excel_exporter import (
    AI_SUMMARY_LABEL_COL,
    AI_SUMMARY_VALUE_COL,
    AI_SUMMARY_UNIT_COL,
    RELATIVE_VALUATION_COLUMNS,
    WorkbookValidationError,
    validate_ai_summary_entry,
    normalize_ticker,
    normalize_company_name,
    is_same_company,
    filter_target_from_peers,
    get_effective_peers,
    update_comp_valuation_sheet_openpyxl,
    populate_ai_summary_sheet_openpyxl,
    validate_generated_workbook,
    export_valuation_model
)
from screener_client import fetch_company_data
from valuation_engine import calculate_valuation


class TestHardenedValuationEngine(unittest.TestCase):

    # =========================================================================
    # REGRESSION TEST 1: AI SUMMARY COLUMN MAPPING
    # =========================================================================
    def test_01_ai_summary_column_mapping(self):
        """
        REGRESSION TEST 1:
        Verify Column B = LABEL, Column C = VALUE.
        Must fail if Column B is numeric/formula and Column C is label.
        """
        print("\n=== RUNNING REGRESSION TEST 1: AI SUMMARY COLUMN MAPPING ===")
        self.assertEqual(AI_SUMMARY_LABEL_COL, "B")
        self.assertEqual(AI_SUMMARY_VALUE_COL, "C")
        self.assertEqual(AI_SUMMARY_UNIT_COL, "D")

        # 1. Correct Input: Label is string, Value is float or formula
        self.assertTrue(validate_ai_summary_entry("Discount Rate", 0.1204))
        self.assertTrue(validate_ai_summary_entry("Discount Rate", "=WACC!K46"))
        self.assertTrue(validate_ai_summary_entry("Terminal Growth Rate", 0.04))
        self.assertTrue(validate_ai_summary_entry("Cost of Equity", 0.105))
        self.assertTrue(validate_ai_summary_entry("Cost of Debt", 0.075))

        # 2. Reversed Input: Label is numeric or formula, Value is text label -> MUST RAISE ValueError
        with self.assertRaises(ValueError) as ctx1:
            validate_ai_summary_entry(0.1204, "Discount Rate")
        self.assertIn("Label", str(ctx1.exception))

        with self.assertRaises(ValueError) as ctx2:
            validate_ai_summary_entry("0.1204", "Discount Rate")
        self.assertIn("numeric", str(ctx2.exception))

        with self.assertRaises(ValueError) as ctx3:
            validate_ai_summary_entry("=WACC!K46", "Discount Rate")
        self.assertIn("formula", str(ctx3.exception))

        # 3. Test writing to workbook and validating cell positions
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "AI Valuation Summary"

        screener_data = {'company_name': 'Test Corp', 'ticker': 'TEST'}
        val_result = {'wacc': 0.12}
        populate_ai_summary_sheet_openpyxl(wb, screener_data, val_result)

        # Check Row 8: Discount Rate (WACC)
        row8_b = ws['B8'].value
        row8_c = str(ws['C8'].value)
        print(f"AI Summary Row 8: B='{row8_b}', C='{row8_c}'")

        self.assertIsInstance(row8_b, str)
        self.assertIn("Discount Rate", row8_b)
        self.assertTrue(row8_c.startswith("="), f"Expected formula in C8, got: {row8_c}")

        # Check Bridge Row 37 & 38 (Intrinsic Value & Current Price)
        row37_b = ws['B37'].value
        row37_c = str(ws['C37'].value)
        row38_b = ws['B38'].value
        row38_c = str(ws['C38'].value)
        row39_c = str(ws['C39'].value)

        self.assertEqual(row37_b, "Intrinsic Value per Share")
        self.assertEqual(row37_c, "=DCF!D42")
        self.assertEqual(row38_b, "Current Market Price")
        self.assertEqual(row38_c, "=DCF!D44")
        self.assertEqual(row39_c, "=(C37-C38)/C38", "Margin of Safety formula must reference Col C!")

        print("REGRESSION TEST 1: PASSED (AI Summary Col B=LABEL, Col C=VALUE)\n")

    # =========================================================================
    # REGRESSION TEST 2: TARGET EXCLUSION FROM PEER GROUP
    # =========================================================================
    def test_02_peer_exclusion(self):
        """
        REGRESSION TEST 2:
        Verify target company is strictly and permanently excluded from peer pool.
        Tests normalization across .NS, .BO, -EQ, case, whitespace, and company name variants.
        """
        print("\n=== RUNNING REGRESSION TEST 2: TARGET EXCLUSION FROM PEER GROUP ===")

        # 1. Normalization tests
        self.assertEqual(normalize_ticker("AMBUJACEM.NS"), "AMBUJACEM")
        self.assertEqual(normalize_ticker(" ambujacem.bo "), "AMBUJACEM")
        self.assertEqual(normalize_ticker("WIPRO-EQ"), "WIPRO")
        self.assertEqual(normalize_ticker("TATASTEEL.BSE"), "TATASTEEL")

        # 2. Company matching tests
        self.assertTrue(is_same_company("AMBUJACEM.NS", "Ambuja Cements Ltd", "AMBUJACEM", "Ambuja Cements"))
        self.assertTrue(is_same_company("WIPRO.BO", "Wipro Limited", "WIPRO.NS", "Wipro"))
        self.assertTrue(is_same_company("500470", "Tata Steel Limited", "TATASTEEL", "Tata Steel"))
        self.assertFalse(is_same_company("ULTRACEMCO", "UltraTech Cement", "AMBUJACEM", "Ambuja Cements"))

        # 3. Target exclusion filter
        target_ticker = "AMBUJACEM"
        target_name = "Ambuja Cements Ltd"
        raw_peers = [
            {'name': 'Ambuja Cements Ltd', 'ticker': 'AMBUJACEM.NS', 'cmp': 640.0, 'mcap': 150000.0},
            {'name': 'UltraTech Cement', 'ticker': 'ULTRACEMCO.NS', 'cmp': 11200.0, 'mcap': 320000.0},
            {'name': 'Shree Cement', 'ticker': 'SHREECEM.BO', 'cmp': 24500.0, 'mcap': 90000.0},
            {'name': 'ACC Limited', 'ticker': 'ACC.NS', 'cmp': 2400.0, 'mcap': 50000.0},
            {'name': 'Ambuja Cements', 'ticker': 'AMBUJACEM', 'cmp': 640.0, 'mcap': 150000.0},  # duplicate
        ]

        filtered = filter_target_from_peers(raw_peers, target_ticker, target_name)
        filtered_names = [p['name'] for p in filtered]
        filtered_tickers = [normalize_ticker(p['ticker']) for p in filtered]

        print(f"Filtered peers for {target_ticker}: {filtered_names}")
        self.assertNotIn("Ambuja Cements Ltd", filtered_names)
        self.assertNotIn("Ambuja Cements", filtered_names)
        self.assertNotIn("AMBUJACEM", filtered_tickers)
        self.assertIn("UltraTech Cement", filtered_names)
        self.assertIn("Shree Cement", filtered_names)
        self.assertIn("ACC Limited", filtered_names)
        self.assertEqual(len(filtered), 3)

        print("REGRESSION TEST 2: PASSED (Target strictly excluded, duplicates removed)\n")

    # =========================================================================
    # REGRESSION TEST 3: DYNAMIC ROW FORMULAS FOR RELATIVE VALUATION
    # =========================================================================
    def test_03_relative_valuation_formulas(self):
        """
        REGRESSION TEST 3:
        Verify EV/Revenue, EV/EBITDA, and P/E formulas dynamically reference the correct row
        and use IFERROR safe wrappers.
        """
        print("\n=== RUNNING REGRESSION TEST 3: DYNAMIC FORMULAS FOR RELATIVE VALUATION ===")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Comp_Valuation"

        screener_data = {'company_name': 'Ambuja Cements', 'ticker': 'AMBUJACEM'}
        peers = get_effective_peers(screener_data)
        update_comp_valuation_sheet_openpyxl(wb, screener_data, peer_count=len(peers), peers=peers)

        # Check Active Peer Rows
        for idx, p in enumerate(peers):
            r = 12 + idx
            ev_rev = ws[f'O{r}'].value
            ev_ebitda = ws[f'P{r}'].value
            pe = ws[f'Q{r}'].value

            # Formula existence
            self.assertTrue(str(ev_rev).startswith('='), f"Row {r} EV/Rev must be formula, got: {ev_rev}")
            self.assertTrue(str(ev_ebitda).startswith('='), f"Row {r} EV/EBITDA must be formula, got: {ev_ebitda}")
            self.assertTrue(str(pe).startswith('='), f"Row {r} P/E must be formula, got: {pe}")

            # Correct row reference
            self.assertIn(f'H{r}', str(ev_rev).replace('$', ''))
            self.assertIn(f'K{r}', str(ev_rev).replace('$', ''))
            self.assertIn(f'H{r}', str(ev_ebitda).replace('$', ''))
            self.assertIn(f'L{r}', str(ev_ebitda).replace('$', ''))
            self.assertIn(f'F{r}', str(pe).replace('$', ''))
            self.assertIn(f'M{r}', str(pe).replace('$', ''))

            # Safe IFERROR wrapper
            self.assertTrue(str(ev_rev).startswith('=IFERROR('))
            self.assertTrue(str(ev_ebitda).startswith('=IFERROR('))
            self.assertTrue(str(pe).startswith('=IFERROR('))

        # Check Peer Statistics (Rows 23 to 28)
        end_r = 12 + len(peers) - 1
        self.assertEqual(ws['O23'].value, f'=MAX(O12:O{end_r})')
        self.assertEqual(ws['O25'].value, f'=MEDIAN(O12:O{end_r})')
        self.assertEqual(ws['O26'].value, f'=AVERAGE(O12:O{end_r})')

        # Check Column N is an empty spacer
        self.assertIsNone(ws['N10'].value, "Column N10 must be None/empty spacer!")
        self.assertIsNone(ws['N12'].value, "Column N12 must be None/empty spacer!")

        # Check Target Valuation Bridge in Row 34: P/E must NOT subtract Net Debt!
        self.assertEqual(ws['Q34'].value, '=Q32', "P/E Implied Market Value (Q34) must be '=Q32' without subtracting Net Debt!")

        # Check Banking / Financial institution handling
        wb_bank = openpyxl.Workbook()
        ws_b = wb_bank.active
        ws_b.title = "Comp_Valuation"
        bank_data = {'company_name': 'HDFC Bank', 'ticker': 'HDFCBANK'}
        update_comp_valuation_sheet_openpyxl(wb_bank, bank_data, peer_count=10)

        # For banks, EV/EBITDA should be 'N/A' for active peer rows and None for cleared leftover rows
        bank_peers = get_effective_peers(bank_data)
        bank_end_r = 12 + len(bank_peers) - 1
        for r in range(12, bank_end_r + 1):
            self.assertEqual(ws_b[f'P{r}'].value, "N/A", f"Banking row {r} EV/EBITDA must be 'N/A'")
        for r in range(bank_end_r + 1, 22):
            self.assertIsNone(ws_b[f'P{r}'].value, f"Leftover row {r} must be None")
        self.assertEqual(ws_b['P32'].value, "N/A", "Banking Implied EV for EV/EBITDA must be 'N/A'")

        print("REGRESSION TEST 3: PASSED (Dynamic row formulas, Column N spacer, Net Debt bridge, and Bank 'N/A' verified)\n")

    # =========================================================================
    # END-TO-END MULTI-COMPANY TEST (7 COMPANIES INCL. ADANI POWER)
    # =========================================================================
    def test_04_multi_company_exports(self):
        """
        END-TO-END TEST:
        Runs the full valuation export and validates generated workbooks across 7 diverse companies:
        1. Adani Power (Power / Utilities)
        2. Wipro (IT)
        3. Ambuja Cements (Cement / Materials)
        4. Tata Steel (Industrial / Materials)
        5. Nestle India (FMCG)
        6. Suzlon Energy (Renewable / Energy)
        7. HDFC Bank (Banking / Financial Institution)
        """
        print("\n=======================================================")
        print("RUNNING END-TO-END TEST ACROSS 7 DIVERSE COMPANIES")
        print("=======================================================")

        test_companies = [
            ("ADANIPOWER", "Adani Power Limited", False),
            ("WIPRO", "Wipro Limited", False),
            ("AMBUJACEM", "Ambuja Cements Limited", False),
            ("TATASTEEL", "Tata Steel Limited", False),
            ("NESTLEIND", "Nestle India Limited", False),
            ("SUZLON", "Suzlon Energy Limited", False),
            ("HDFCBANK", "HDFC Bank Limited", True)
        ]

        results = {}

        for ticker, name, is_fin in test_companies:
            print(f"\n-------------------------------------------------------")
            print(f"Testing Company: {ticker} ({name}) | Financial: {is_fin}")
            print(f"-------------------------------------------------------")
            try:
                # 1. Fetch live data
                screener_data = fetch_company_data(ticker)
                if not screener_data or not screener_data.get('company_name'):
                    screener_data['company_name'] = name
                    screener_data['ticker'] = ticker

                # 2. Run valuation
                val_result = calculate_valuation(screener_data)

                # 3. Export workbook
                export_path = export_valuation_model(screener_data, val_result)
                print(f"[EXPORT] Successfully generated workbook: {export_path}")

                # 4. Post-generation validation
                validate_generated_workbook(export_path, ticker, name, is_financial=is_fin)
                results[ticker] = "PASS"
                print(f"[RESULT] {ticker}: PASS")
            except Exception as e:
                results[ticker] = f"FAIL: {e}"
                print(f"[RESULT] {ticker}: FAIL -> {e}")

        print("\n=======================================================")
        print("MULTI-COMPANY TEST SUMMARY:")
        print("=======================================================")
        for t, res in results.items():
            print(f"  {t}: {res}")
            self.assertEqual(res, "PASS", f"Company {t} failed validation: {res}")


if __name__ == '__main__':
    unittest.main()
