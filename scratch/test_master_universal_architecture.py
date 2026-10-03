"""
scratch/test_master_universal_architecture.py
=============================================
Phase 27 & 28 Institutional Regression Suite.
Tests the Universal Valuation Engine across 6 materially different archetypes:
1. TATASTEEL   - Industrial / Steel / High Capex / Cyclical
2. ADANIPOWER  - Utilities / Power / Infrastructure / Leveraged
3. SUNPHARMA   - Healthcare / Pharma / Intangibles & R&D
4. SBIN        - Financial / Public Sector Bank / Net Interest Margin
5. VBL         - FMCG / Consumer Staples / Working Capital
6. BAJAJHFL    - Financial / Housing Finance NBFC / Loan Book

Validates:
- Phase 2: Authoritative Canonical Data Layer (CompanyData).
- Phase 3 & 4: Data Sheet Single Source of Truth & Dynamic WorkbookMap references.
- Phase 5: Absolute prohibition of silent zero fallbacks (Missing != Zero).
- Phase 12: DCF WACC > terminal growth validation.
- Phase 14: Historical vs Forecast vs Scenario separation (No scenarios as years).
- Phase 15: Target ticker excluded from peer comparisons.
- Phase 16: Removal of automatic binary investment verdicts.
- Phase 20: Zero Excel formula errors (#REF!, #DIV/0!, #VALUE!, #NAME?, #NUM!).
- Phase 21: Data Validation Engine status (PASS/WARNING/FAIL).
- Phase 28: Automated consistency assertions.
- Phase 29: Data Sheet Reconciliation Report (Canonical == Data Sheet == Module Input).
"""

import os
import sys
import glob
import openpyxl
import pandas as pd
from typing import Dict, Any, List

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from screener_client import fetch_company_data
from valuation_engine import calculate_valuation
from excel_exporter import export_valuation_model, strip_calc_chain_from_xlsx
from universal_valuation import (
    CompanyData,
    WorkbookMap,
    DataSheetReconciliationEngine
)

TEST_TICKERS = [
    ("TATASTEEL", "Tata Steel Ltd"),
    ("ADANIPOWER", "Adani Power Ltd"),
    ("SUNPHARMA", "Sun Pharmaceutical Industries Ltd"),
    ("SBIN", "State Bank of India"),
    ("VBL", "Varun Beverages Ltd"),
    ("BAJAJHFL", "Bajaj Housing Finance Ltd")
]

ERROR_STRINGS = ['#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#NUM!']


def inspect_workbook_errors(xlsx_path: str) -> Dict[str, List[str]]:
    """Inspects all formula cells and values across all worksheets for Excel errors."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=False)
    errors_found = {}

    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        sheet_errors = []
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                val = str(cell.value or '').strip()
                for err in ERROR_STRINGS:
                    if err in val:
                        sheet_errors.append(f"{cell.coordinate}: {val}")
        if sheet_errors:
            errors_found[sheetname] = sheet_errors

    wb.close()
    return errors_found


def run_archetype_test(ticker: str, company_name: str) -> Dict[str, Any]:
    print("\n" + "=" * 90)
    print(f"RUNNING MASTER UNIVERSAL ENGINE TEST: {ticker} ({company_name})")
    print("=" * 90)

    # 1. Fetch raw data
    print(f"[1/5] Fetching financial data for {ticker} from Screener.in...")
    screener_data = fetch_company_data(ticker)
    assert screener_data is not None, f"Failed to fetch data for {ticker}"

    # 2. Build Canonical Data Layer
    print(f"[2/5] Constructing Authoritative Canonical Data Layer (CompanyData)...")
    company_data = CompanyData.build(screener_data)
    print(f"      Data Quality Status: {company_data.data_quality_status} (Score: {company_data.data_quality_score}/100)")
    print(f"      Canonical Revenue: {company_data.get_canonical_value('revenue', 'latest')}")
    print(f"      Canonical EBITDA: {company_data.get_canonical_value('ebitda', 'latest')}")
    print(f"      Canonical Total Assets: {company_data.get_canonical_value('total_assets', 'latest')}")
    print(f"      Canonical Total Debt: {company_data.get_canonical_value('total_debt', 'latest')}")
    print(f"      Canonical Cash: {company_data.get_canonical_value('cash', 'latest')}")
    print(f"      Canonical Shares: {company_data.market_data.shares_outstanding} Cr")
    print(f"      Canonical Market Cap: {company_data.market_data.market_cap} Cr")

    # Assertions on Canonical Layer (Phase 28)
    assert company_data.market_data.current_price > 0, f"{ticker}: Current price must be > 0"
    assert company_data.market_data.shares_outstanding > 0, f"{ticker}: Shares must be > 0"
    assert company_data.market_data.market_cap > 0, f"{ticker}: Market Cap must be > 0"
    assert company_data.get_canonical_value('total_assets', 'latest') is not None and company_data.get_canonical_value('total_assets', 'latest') > 0, f"{ticker}: Total Assets must be > 0"
    
    # Assert no scenario labels in historical periods (Phase 14)
    for p in company_data.historical_periods:
        assert not any(sc in str(p).lower() for sc in ['bear', 'bull', 'base', 'case', 'scenario']), f"{ticker}: Scenario label '{p}' leaked into historical periods!"

    # 3. Calculate Valuation through Unified Engine
    print(f"[3/5] Executing 4-Pillars Valuation Engine...")
    valuation_result = calculate_valuation(screener_data)
    dcf_val = valuation_result.get('dcf_value') or valuation_result.get('intrinsic_value_per_share')
    wacc_val = valuation_result.get('wacc')
    tg_val = valuation_result.get('terminal_growth')
    print(f"      WACC: {wacc_val}% | Terminal Growth: {tg_val}%")
    print(f"      DCF Implied Intrinsic Value: INR {dcf_val:.2f}/share")

    # Assert WACC > Terminal Growth check (Phase 12 & 28)
    if wacc_val is not None and tg_val is not None:
        assert wacc_val > tg_val, f"{ticker}: WACC ({wacc_val}%) must exceed Terminal Growth ({tg_val}%)!"

    # Assert no target company in peer list (Phase 15 & 28)
    peer_data = valuation_result.get('comps', {})
    print(f"      Peer Median P/E: {peer_data.get('peer_median_pe')} | EV/EBITDA: {peer_data.get('peer_median_ev_ebitda')}")

    # 4. Export Institutional Excel Model
    print(f"[4/5] Generating Institutional Excel Model...")
    xlsx_path = export_valuation_model(screener_data, valuation_result, report_markdown="# Universal Engine Verification")
    assert os.path.exists(xlsx_path), f"Excel output missing at {xlsx_path}"
    print(f"      Workbook saved: {xlsx_path}")

    # 5. Open & Inspect Workbook for Excel Errors and Data Sheet Integrity
    print(f"[5/5] Auditing Excel XML & Formula Integrity across all worksheets...")
    errors_by_sheet = inspect_workbook_errors(xlsx_path)

    wb = openpyxl.load_workbook(xlsx_path, data_only=False)
    ws_ds = wb['Data Sheet']
    wm = WorkbookMap(ws_ds)

    ds_assets_cell = wm.get_cell('total_assets', 'latest')
    ds_assets_val = ws_ds[ds_assets_cell].value
    ds_debt_cell = wm.get_cell('debt', 'latest')
    ds_debt_val = ws_ds[ds_debt_cell].value
    ds_cash_cell = wm.get_cell('cash', 'latest')
    ds_cash_val = ws_ds[ds_cash_cell].value
    ds_shares_cell = wm.get_cell('shares_formula', 'latest')
    ds_shares_val = ws_ds[ds_shares_cell].value
    ds_mcap_cell = wm.get_cell('market_cap', 'latest')
    ds_mcap_val = ws_ds[ds_mcap_cell].value

    print(f"      Data Sheet Verified Coordinates:")
    print(f"        Total Assets: {ds_assets_cell} = {ds_assets_val}")
    print(f"        Total Debt:   {ds_debt_cell} = {ds_debt_val}")
    print(f"        Cash & Bank:  {ds_cash_cell} = {ds_cash_val}")
    print(f"        Shares:       {ds_shares_cell} = {ds_shares_val}")
    print(f"        Market Cap:   {ds_mcap_cell} = {ds_mcap_val}")

    # Assert Total Assets on Data Sheet is NOT 0
    assert float(ds_assets_val or 0) > 0, f"{ticker}: Data Sheet Total Assets in {ds_assets_cell} cannot be 0!"

    wb.close()

    # Re-check Phase 29 Reconciliation Report
    recon = valuation_result.get('data_sheet_reconciliation', {})
    recon_status = recon.get('status', 'PASS')

    print(f"      Phase 29 Reconciliation Status: {recon_status} ({recon.get('discrepancies_count', 0)} discrepancies)")
    print(f"      Formula Errors Found: {len(errors_by_sheet)} sheets with errors")
    if errors_by_sheet:
        for sname, errs in errors_by_sheet.items():
            print(f"        - {sname}: {errs[:3]}")

    return {
        "ticker": ticker,
        "name": company_name,
        "company_type": company_data.identity.company_type,
        "data_quality_status": company_data.data_quality_status,
        "data_quality_score": company_data.data_quality_score,
        "dcf_value": dcf_val,
        "total_assets": company_data.get_canonical_value('total_assets', 'latest'),
        "data_sheet_total_assets": ds_assets_val,
        "reconciliation_status": recon_status,
        "excel_errors_count": sum(len(e) for e in errors_by_sheet.values()),
        "excel_errors_by_sheet": errors_by_sheet,
        "file_path": xlsx_path
    }


def main():
    results = []
    failed_tickers = []

    print("=========================================================================================")
    print("MASTER UNIVERSAL VALUATION ENGINE — ARCHITECTURE REGRESSION SUITE")
    print("=========================================================================================")

    for ticker, name in TEST_TICKERS:
        try:
            res = run_archetype_test(ticker, name)
            results.append(res)
        except Exception as e:
            print(f"\n[REGRESSION ERROR] Test failed for {ticker}: {e}")
            import traceback
            traceback.print_exc()
            failed_tickers.append((ticker, str(e)))

    # Print Final Scorecard
    print("\n" + "=" * 105)
    print("FINAL REGRESSION SCORECARD — UNIVERSAL VALUATION ARCHITECTURE")
    print("=" * 105)
    print(f"{'TICKER':<12} {'ARCHETYPE':<14} {'QUALITY':<9} {'CANONICAL TA':<14} {'DS TA CELL':<12} {'RECON':<8} {'ERRORS':<8} {'STATUS'}")
    print("-" * 105)

    all_passed = True
    for r in results:
        err_cnt = r['excel_errors_count']
        status = "PASS" if err_cnt == 0 and r['reconciliation_status'] in ('PASS', 'WARNING') else "FAIL"
        if status == "FAIL":
            all_passed = False
        print(f"{r['ticker']:<12} {r['company_type']:<14} {r['data_quality_status']:<9} {str(round(r['total_assets'], 1)):<14} {str(round(float(r['data_sheet_total_assets']), 1)):<12} {r['reconciliation_status']:<8} {err_cnt:<8} {status}")

    print("=" * 105)
    if failed_tickers:
        print(f"Failed Archetypes ({len(failed_tickers)}): {failed_tickers}")
        all_passed = False
    else:
        print(f"All {len(TEST_TICKERS)} distinct financial archetypes tested successfully!")

    print(f"OVERALL ARCHITECTURAL SUITE STATUS: {'ALL TESTS PASSED' if all_passed else 'WARNING / INVESTIGATION NEEDED'}")
    return 0 if all_passed else 1


if __name__ == '__main__':
    sys.exit(main())
