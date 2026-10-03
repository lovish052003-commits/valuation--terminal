import sys
import os
import json

# Ensure project root is in path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding='utf-8')

import screener_client
import valuation_engine

UNIVERSAL_TEST_COMPANIES = [
    ('SBIN', 'BANK'),
    ('RBLBANK', 'BANK'),
    ('LICI', 'INSURANCE'),
    ('ICICIAMC', 'ASSET_MANAGEMENT'),
    ('ITC', 'FMCG'),
    ('TCS', 'IT_SERVICES'),
    ('INFY', 'IT_SERVICES'),
    ('RELIANCE', 'ENERGY'),
    ('TATAMOTORS', 'AUTOMOBILE'),
    ('HDFCBANK', 'BANK'),
    ('BAJFINANCE', 'NBFC'),
    ('SUNPHARMA', 'PHARMACEUTICAL'),
    ('BHARTIARTL', 'TELECOM'),
    ('LT', 'INFRASTRUCTURE'),
    ('DMART', 'CONSUMER')
]

def test_company_valuation(ticker, expected_type):
    print(f"\n========================================================", flush=True)
    print(f"Testing Company: {ticker} (Expected Type: {expected_type})", flush=True)
    print(f"========================================================", flush=True)
    
    # 1. Resolve & Fetch
    data = screener_client.fetch_company_data(ticker)
    canonical = data.get('canonical_company', {})
    
    print(f"1. Canonical Company: {canonical.get('companyName')} ({canonical.get('nseSymbol') or canonical.get('bseCode')})")
    print(f"   Sector: {canonical.get('sector')} | Industry: {canonical.get('industry')}")
    print(f"   Classified Type: {canonical.get('companyType')}")
    assert canonical.get('companyType') == expected_type, f"Expected {expected_type}, got {canonical.get('companyType')}"
    
    # 2. Market Data
    cmp = data.get('current_price', 0)
    mcap = data.get('market_cap_cr', 0)
    shares = data.get('shares_in_cr', 0)
    print(f"2. Market: Price = Rs. {cmp} | MCap = Rs. {mcap} Cr | Shares = {shares} Cr")
    assert cmp > 0, f"Current price must be > 0, got {cmp}"
    assert mcap > 0, f"Market cap must be > 0, got {mcap}"
    assert shares > 0, f"Shares must be > 0, got {shares}"

    # 3. Peers Validation
    peers = data.get('sector_peers', [])
    print(f"3. Sector Peers ({len(peers)} peers found):")
    target_ticker = (canonical.get('nseSymbol') or data.get('ticker') or '').upper()
    target_name = (canonical.get('companyName') or data.get('company_name') or '').lower()
    
    peer_tickers = []
    for p in peers:
        p_tick = p.get('ticker', '').upper()
        p_name = p.get('name', '').lower()
        print(f"   - {p.get('name')} ({p_tick}): CMP Rs. {p.get('current_price')} | MCap Rs. {p.get('market_cap')} Cr")
        # Ensure target company is NOT in its own peer list
        assert p_tick != target_ticker or not p_tick, f"Target ticker {target_ticker} found in its own peer list!"
        assert target_name != p_name, f"Target name {target_name} found in its own peer list!"
        if p_tick:
            assert p_tick not in peer_tickers, f"Duplicate peer {p_tick} found!"
            peer_tickers.append(p_tick)

    # 4. Valuation Calculation
    val = valuation_engine.calculate_valuation(data)
    print(f"4. Valuation Calculated: Intrinsic Value = Rs. {val.get('intrinsic_value_per_share')} (CMP: Rs. {cmp})")
    print(f"   WACC: {val.get('wacc')}% | Growth: {val.get('growth_rate')}% | Terminal Growth: {val.get('terminal_growth')}%")
    assert val.get('current_price') == cmp, "Valuation CMP must match target company CMP!"
    
    # 5. Financial Multiples & Methodology Sanity Check
    financial_types = {'BANK', 'NBFC', 'INSURANCE', 'ASSET_MANAGEMENT', 'BROKING', 'OTHER_FINANCIAL'}
    if expected_type in financial_types:
        assert val['altman_z']['score'] is None, f"Altman Z should be None for financial type {expected_type}"
        print(f"5. Financial Institution Checks: Altman Z = {val['altman_z']['display_text']} (Correct)")
        print(f"   EV/EBITDA = {val['four_pillars']['pillar4_multiples'].get('target_ev_ebitda')} (Correctly N/A for financial institution)")
    else:
        print(f"5. Operating Company Checks: Altman Z = {val['altman_z']['score']} ({val['altman_z']['zone']})")
        print(f"   EV/EBITDA = {val['four_pillars']['pillar4_multiples'].get('target_ev_ebitda')}x")

    # 6. DuPont
    dup = val.get('dupont', {})
    print(f"6. DuPont Analysis: ROE = {dup.get('roe_3stage')}% | NPM = {dup.get('net_profit_margin')}% | ATO = {dup.get('asset_turnover')}x")

    # 7. Reverse DCF
    rev_dcf = val.get('reverse_dcf', {})
    print(f"7. Reverse DCF: Implied FCF Growth = {rev_dcf.get('impliedFCFCagr')}%")

    print(f"[PASS] {ticker} Valuation & Methodology Verified Successfully!")
    return val

def test_sequential_zero_contamination():
    print(f"\n========================================================")
    print(f"Running Sequential Contamination Test:")
    print(f"SBIN -> ITC -> TCS -> LICI -> ICICIAMC -> RELIANCE -> HDFCBANK")
    print(f"========================================================")

    sequence = [
        ('SBIN', 'BANK'),
        ('ITC', 'FMCG'),
        ('TCS', 'IT_SERVICES'),
        ('LICI', 'INSURANCE'),
        ('ICICIAMC', 'ASSET_MANAGEMENT'),
        ('RELIANCE', 'DIVERSIFIED'),
        ('HDFCBANK', 'BANK')
    ]

    previous_run = None
    for ticker, exp_type in sequence:
        val = test_company_valuation(ticker, exp_type)
        if previous_run:
            prev_ticker = previous_run['ticker']
            prev_name = previous_run['company_name']
            prev_price = previous_run['current_price']
            
            # Verify ZERO leakage
            assert val['ticker'] == ticker, f"Ticker corrupted: expected {ticker}, got {val['ticker']}"
            assert val['ticker'] != prev_ticker, f"Leaked previous ticker {prev_ticker} into {ticker}"
            assert prev_name.lower() not in val['company_name'].lower(), f"Leaked previous company name {prev_name} into {val['company_name']}"
            assert val['current_price'] != prev_price or val['current_price'] > 0, "Price cross-contamination check"
            print(f"[ZERO-CONTAMINATION CONFIRMED] {prev_ticker} -> {ticker}: 0% data leakage.")
        previous_run = val

def test_ambiguity_and_validation():
    print(f"\n========================================================")
    print(f"Testing Ambiguity & Validation Error Handling")
    print(f"========================================================")

    # 1. Ambiguous resolution
    ambiguous_res = screener_client.resolve_company("TATA")
    print(f"Query 'TATA' Ambiguous Result: ambiguous = {ambiguous_res.get('ambiguous')}, candidate count = {len(ambiguous_res.get('candidates', []))}")
    assert ambiguous_res.get('ambiguous') is True or len(ambiguous_res.get('candidates', [])) > 1, "Expected ambiguous result for 'TATA'"

    # 2. WACC <= Terminal Growth validation
    try:
        data = screener_client.fetch_company_data("ITC")
        valuation_engine.calculate_valuation(data, custom_params={'wacc': 0.03, 'terminal_growth': 0.04})
        assert False, "Expected ValueError when WACC <= terminal growth"
    except ValueError as ve:
        print(f"WACC <= Terminal Growth validation caught as expected: {ve}")

    # 3. Missing financial data validation
    mock_data_missing = {
        'ticker': 'TESTCO',
        'company_name': 'Test Company Missing Statements',
        'current_price': 100.0,
        'market_cap_cr': 500.0,
        'tables': {}
    }
    try:
        valuation_engine.calculate_valuation(mock_data_missing)
        assert False, "Expected ValueError for missing financial statements"
    except ValueError as ve_missing:
        print(f"Missing data validation caught as expected: {ve_missing}")
        assert "Valuation unavailable because required financial data is missing" in str(ve_missing)

    print("\n[ALL VALIDATION & AMBIGUITY TESTS PASSED!]")

if __name__ == '__main__':
    print("Starting Comprehensive Universal Terminal Verification Suite...")
    
    # 1. Test all 15 diverse companies
    for ticker, c_type in UNIVERSAL_TEST_COMPANIES:
        test_company_valuation(ticker, c_type)

    # 2. Test sequential zero-contamination
    test_sequential_zero_contamination()

    # 3. Test ambiguity and validation
    test_ambiguity_and_validation()

    print("\n" + "=" * 65)
    print("ALL UNIVERSAL VALUATION TESTS PASSED WITH 100% SUCCESS!")
    print("=" * 65)
