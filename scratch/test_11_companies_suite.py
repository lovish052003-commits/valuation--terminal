import os
import sys
import time
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from screener_client import fetch_company_data, is_same_company
from valuation_engine import calculate_valuation
from excel_exporter import get_effective_peers

TEST_COMPANIES = [
    "ONGC",
    "RELIANCE",
    "TCS",
    "ITC",
    "TATAMOTORS",
    "SUNPHARMA",
    "HDFCBANK",
    "SBIN",
    "BAJFINANCE",
    "LICI",
    "ICICIAMC"
]

def run_company_test(ticker_or_name):
    print(f"\n========================================================")
    print(f"Testing Company: {ticker_or_name}")
    print(f"========================================================")
    t0 = time.time()
    sd = fetch_company_data(ticker_or_name)
    c_name = sd.get('company_name', '')
    ticker = sd.get('ticker', '')
    classification = sd.get('company_classification', {})
    mcap = sd.get('market_cap_cr', 0)
    cmp = sd.get('current_price', 0)
    print(f"Resolved: {c_name} ({ticker}) | Sector: {sd.get('sector')} | Industry: {sd.get('industry')}")
    print(f"Classification: {classification.get('sector_type')} | Valuation Method: {classification.get('valuation_methodology')}")
    print(f"CMP: Rs. {cmp} | MCap: Rs. {mcap} Cr")

    val = calculate_valuation(sd)
    method = val.get('valuation_method')
    intrinsic = val.get('intrinsic_value_per_share')
    wacc = val.get('wacc')
    peers = get_effective_peers(sd, val)
    comps_summary = val.get('peers', {})

    print(f"Valuation Output: Method={method} | Intrinsic Value={intrinsic} | WACC={wacc}")
    print(f"Peers Count: {len(peers)}")

    # 1. Verification: Target must NEVER be in peers
    for p in peers:
        p_tick = p.get('ticker', '')
        p_name = p.get('name', '')
        is_match = is_same_company(p_tick, p_name, ticker, c_name)
        if is_match or p_tick.upper() == ticker.upper():
            raise AssertionError(f"CRITICAL FAULT: Target {c_name} ({ticker}) found in peer set: {p_name} ({p_tick})")
    print("PASS: Target company is 100% excluded from peers.")

    # 2. Print top peers
    print("Top Comps:")
    for idx, p in enumerate(peers[:8], 1):
        print(f"  {idx}. {p.get('name')} ({p.get('ticker')}): MCap={p.get('mcap')} Cr, PE={p.get('pe')}, EV/EBITDA={p.get('ev_ebitda')}, Score={p.get('peer_score')}")

    # 3. Check Comps Summary
    if comps_summary:
        print(f"Comps Summary: PE Median={comps_summary.get('pe_median')}, EV/EBITDA Median={comps_summary.get('ev_ebitda_median')}")
        if 'divergence_warning' in comps_summary:
            print(f"  Divergence Warning: {comps_summary['divergence_warning']}")

    t_elapsed = round(time.time() - t0, 2)
    print(f"Completed in {t_elapsed}s")
    return {
        'company': c_name,
        'ticker': ticker,
        'peers': [p.get('name') for p in peers[:8]],
        'peer_count': len(peers),
        'method': method,
        'intrinsic': intrinsic,
        'wacc': wacc
    }

if __name__ == '__main__':
    results = {}
    print("STARTING 11-COMPANY TEST SUITE")
    for comp in TEST_COMPANIES:
        try:
            results[comp] = run_company_test(comp)
        except Exception as e:
            print(f"ERROR testing {comp}: {e}")
            import traceback
            traceback.print_exc()

    print("\n\n========================================================")
    print("STARTING SEQUENTIAL ZERO-CONTAMINATION TEST")
    print("Sequence: ONGC -> TCS -> ITC -> SBIN -> LICI -> ICICIAMC -> ONGC")
    print("========================================================")
    seq = ["ONGC", "TCS", "ITC", "SBIN", "LICI", "ICICIAMC", "ONGC"]
    seq_results = []
    for s_comp in seq:
        res = run_company_test(s_comp)
        seq_results.append((s_comp, res['peers'], res['wacc'], res['intrinsic']))

    ongc_1 = seq_results[0]
    ongc_7 = seq_results[6]
    print("\n--- ZERO-CONTAMINATION COMPARISON ---")
    print(f"ONGC Run 1 Peers: {ongc_1[1]}")
    print(f"ONGC Run 7 Peers: {ongc_7[1]}")
    print(f"ONGC Run 1 WACC: {ongc_1[2]} | Run 7 WACC: {ongc_7[2]}")
    print(f"ONGC Run 1 Intrinsic: {ongc_1[3]} | Run 7 Intrinsic: {ongc_7[3]}")

    assert ongc_1[1] == ongc_7[1], f"Peer lists differ! Run 1: {ongc_1[1]} != Run 7: {ongc_7[1]}"
    assert ongc_1[2] == ongc_7[2], f"WACC differs! Run 1: {ongc_1[2]} != Run 7: {ongc_7[2]}"
    assert ongc_1[3] == ongc_7[3], f"Intrinsic value differs! Run 1: {ongc_1[3]} != Run 7: {ongc_7[3]}"
    print("\nZERO-CONTAMINATION TEST 100% PASSED! State is completely clean across sequential runs.")
