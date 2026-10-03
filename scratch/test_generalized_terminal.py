import json
import requests
import sys
import time
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

BASE_URL = "http://127.0.0.1:5000"

def wait_for_server():
    print("Waiting for server to be ready...")
    for _ in range(10):
        try:
            r = requests.get(f"{BASE_URL}/", timeout=2)
            if r.status_code == 200:
                print("Server is ready!")
                return
        except Exception:
            time.sleep(1)
    print("Server wait timeout, proceeding anyway...")

def test_peer_selection_isolation():
    print("\n=======================================================")
    print("TEST 1: Excel Exporter Peer Selection & Sector Isolation")
    print("=======================================================")
    import excel_exporter
    # Test banking vs NBFC sector keys
    bank_key = excel_exporter.get_sector_key({"sector": "Financial Services", "industry": "Private Sector Bank", "company_name": "RBL Bank Ltd", "ticker": "RBLBANK"})
    hdfc_key = excel_exporter.get_sector_key({"sector": "Financial Services", "industry": "Private Sector Bank", "company_name": "HDFC Bank Ltd", "ticker": "HDFCBANK"})
    nbfc_key = excel_exporter.get_sector_key({"sector": "Financial Services", "industry": "Non Banking Financial Company (NBFC)", "company_name": "Bajaj Finance Ltd", "ticker": "BAJFINANCE"})
    itc_key = excel_exporter.get_sector_key({"sector": "Consumer Goods", "industry": "Cigarettes", "company_name": "ITC Ltd", "ticker": "ITC"})
    tcs_key = excel_exporter.get_sector_key({"sector": "Information Technology", "industry": "Computers - Software", "company_name": "Tata Consultancy Services Ltd", "ticker": "TCS"})

    print(f"RBL Bank Sector Key: {bank_key}")
    print(f"HDFC Bank Sector Key: {hdfc_key}")
    print(f"Bajaj Finance Sector Key: {nbfc_key}")
    print(f"ITC Sector Key: {itc_key}")
    print(f"TCS Sector Key: {tcs_key}")

    assert bank_key == "banking", f"Expected banking for RBL, got {bank_key}"
    assert hdfc_key == "banking", f"Expected banking for HDFC, got {hdfc_key}"
    assert nbfc_key == "nbfc_finance", f"Expected nbfc_finance for Bajaj Finance, got {nbfc_key}"
    assert itc_key == "fmcg", f"Expected fmcg for ITC, got {itc_key}"
    assert tcs_key == "it", f"Expected it for TCS, got {tcs_key}"

    # Verify RBL peers
    bank_peers = excel_exporter.SECTOR_LEADERS['banking']
    peer_tickers = [p['ticker'] for p in bank_peers]
    print(f"Banking Sector Peer Tickers: {peer_tickers}")
    assert "HDFCBANK" in peer_tickers
    assert "ICICIBANK" in peer_tickers
    assert "SBIN" in peer_tickers
    assert "BAJFINANCE" not in peer_tickers, "BAJFINANCE must not be in banking peer list"
    assert "CHOLAFIN" not in peer_tickers, "CHOLAFIN must not be in banking peer list"
    print(">>> Peer Selection Isolation test PASSED!")

def test_api_analyze():
    wait_for_server()
    print("\n=======================================================")
    print("TEST 2: POST /api/analyze with RBL Bank (skipAi=True)")
    print("=======================================================")
    payload = {"company": "RBL Bank", "skipAi": True}
    res = requests.post(f"{BASE_URL}/api/analyze", json=payload, timeout=60)
    print("Status:", res.status_code)
    assert res.status_code == 200, f"Expected 200, got {res.status_code}"
    data = res.json()
    assert data.get("success") is True, "Expected success=True"
    
    company = data.get("company", {})
    ticker = company.get("ticker")
    company_type = company.get("company_type")
    run_id = data.get("run_id")
    val = data.get("valuation", {})
    altman_z = val.get("altman_z", {})
    four_pillars = val.get("four_pillars", {})
    multiples = four_pillars.get("pillar4_multiples", {})
    intrinsic = val.get("intrinsic_value_per_share")
    verdict = val.get("verdict")

    print(f"Ticker: {ticker}, Type: {company_type}")
    print(f"Run ID: {run_id}, Data Source: {data.get('data_source')}, As Of: {data.get('data_as_of')}")
    print(f"Altman Z: {altman_z.get('display_text')}")
    print(f"Multiples: P/E={multiples.get('target_pe')}, P/B={multiples.get('target_pb')}, EV/EBITDA={multiples.get('target_ev_ebitda')}")
    print(f"DCF Intrinsic Value: Rs {intrinsic}, Verdict: {verdict}")

    assert ticker == "RBLBANK", f"Expected RBLBANK, got {ticker}"
    assert company_type == "BANK", f"Expected BANK, got {company_type}"
    assert altman_z.get("score") is None, f"Expected Altman Z score to be None for BANK, got {altman_z.get('score')}"
    assert altman_z.get("is_applicable") is False, "Expected is_applicable=False"
    assert "Not applicable" in altman_z.get("display_text", "")
    assert multiples.get("target_ev_ebitda") is None, "EV/EBITDA should be None for Bank"
    assert multiples.get("target_pe") is not None and multiples.get("target_pe") > 0, "P/E should exist for Bank"
    assert multiples.get("target_pb") is not None and multiples.get("target_pb") > 0, "P/B should exist for Bank"
    assert intrinsic is not None and intrinsic > 0, "DCF intrinsic value must exist"

    print(">>> RBL Bank analyze test PASSED!")
    return data

def test_api_valuation():
    print("\n=======================================================")
    print("TEST 3: POST /api/valuation with ITC (General Endpoint)")
    print("=======================================================")
    payload = {"company": "ITC", "skipAi": True}
    res = requests.post(f"{BASE_URL}/api/valuation", json=payload, timeout=60)
    print("Status:", res.status_code)
    assert res.status_code == 200, f"Expected 200, got {res.status_code}: {res.text}"
    data = res.json()
    assert data.get("success") is True, "Expected success=True"
    
    company = data.get("company", {})
    ticker = company.get("ticker")
    company_type = company.get("company_type")
    run_id = data.get("run_id")
    val = data.get("valuation", {})
    altman_z = val.get("altman_z", {})
    four_pillars = val.get("four_pillars", {})
    multiples = four_pillars.get("pillar4_multiples", {})
    intrinsic = val.get("intrinsic_value_per_share")
    verdict = val.get("verdict")

    print(f"Ticker: {ticker}, Type: {company_type}")
    print(f"Run ID: {run_id}, As Of: {data.get('data_as_of')}")
    print(f"Altman Z: {altman_z.get('display_text')}")
    print(f"Multiples: P/E={multiples.get('target_pe')}, EV/EBITDA={multiples.get('target_ev_ebitda')}")
    print(f"DCF Intrinsic Value: Rs {intrinsic}, Verdict: {verdict}")

    assert ticker == "ITC", f"Expected ITC, got {ticker}"
    assert company_type == "NON_FINANCIAL", f"Expected NON_FINANCIAL, got {company_type}"
    assert altman_z.get("score") is not None, "Expected Altman Z score for ITC"
    assert altman_z.get("is_applicable") is True, "Expected is_applicable=True for ITC"
    assert multiples.get("target_ev_ebitda") is not None, "EV/EBITDA must exist for Non-Financial"
    assert intrinsic is not None and intrinsic > 0, "DCF intrinsic value must exist"
    print(">>> ITC valuation test PASSED!")
    return data

def test_company_pipeline(company_name, expected_ticker, expected_type):
    print(f"\n=======================================================")
    print(f"TEST: POST /api/valuation with {company_name}")
    print("=======================================================")
    payload = {"company": company_name, "skipAi": True}
    res = requests.post(f"{BASE_URL}/api/valuation", json=payload, timeout=60)
    print("Status:", res.status_code)
    assert res.status_code == 200, f"Failed for {company_name}: {res.status_code}, {res.text}"
    data = res.json()
    assert data.get("success") is True, f"Failed for {company_name}: {data}"
    
    company = data.get("company", {})
    ticker = company.get("ticker")
    company_type = company.get("company_type")
    run_id = data.get("run_id")
    val = data.get("valuation", {})
    intrinsic = val.get("intrinsic_value_per_share")
    mos = val.get("margin_of_safety_pct")
    altman_z = val.get("altman_z", {})
    four_pillars = val.get("four_pillars", {})
    multiples = four_pillars.get("pillar4_multiples", {})

    print(f"Ticker: {ticker} (expected: {expected_ticker}), Type: {company_type} (expected: {expected_type})")
    print(f"Run ID: {run_id}, As Of: {data.get('data_as_of')}")
    print(f"DCF Intrinsic Value: Rs {intrinsic}, Margin of Safety: {mos}%")
    print(f"Altman Z: {altman_z.get('display_text')}")

    if isinstance(expected_ticker, (list, tuple)):
        assert ticker in expected_ticker, f"Expected one of {expected_ticker}, got {ticker}"
    else:
        assert ticker == expected_ticker, f"Expected {expected_ticker}, got {ticker}"
    assert company_type == expected_type, f"Expected {expected_type}, got {company_type}"

    if expected_type in ("BANK", "NBFC_FINANCIAL"):
        assert altman_z.get("score") is None, f"Expected Altman Z score None for {ticker}"
        assert altman_z.get("is_applicable") is False
        assert multiples.get("target_ev_ebitda") is None
    else:
        assert altman_z.get("score") is not None, f"Expected Altman Z score for {ticker}"
        assert altman_z.get("is_applicable") is True
        assert multiples.get("target_ev_ebitda") is not None

    print(f">>> {company_name} test PASSED!")
    return data

def test_sequence_zero_leakage():
    print("\n=======================================================")
    print("TEST: Sequence for Zero State Leakage: RBL -> ITC -> TCS -> RELIANCE")
    print("=======================================================")
    sequence = [
        ("RBL Bank", "RBLBANK", "BANK"),
        ("ITC", "ITC", "NON_FINANCIAL"),
        ("TCS", "TCS", "NON_FINANCIAL"),
        ("Reliance Industries", "RELIANCE", "NON_FINANCIAL")
    ]
    prev_run_id = None
    prev_ticker = None
    for name, exp_ticker, exp_type in sequence:
        res = requests.post(f"{BASE_URL}/api/valuation", json={"company": name, "skipAi": True}, timeout=60).json()
        curr_ticker = res["company"]["ticker"]
        curr_name = res["company"]["name"]
        curr_run_id = res["run_id"]
        print(f"Ran {name:20}: ticker={curr_ticker:10}, company_name='{curr_name:30}', run_id={curr_run_id}")
        assert curr_ticker == exp_ticker
        assert curr_run_id != prev_run_id, f"Run ID should be unique across requests! Got {curr_run_id}"
        assert curr_ticker != prev_ticker, "Ticker should change"
        # Ensure RBL data doesn't leak into ITC/TCS/RELIANCE
        if exp_ticker != "RBLBANK":
            assert "RBL" not in curr_name.upper(), f"Leaked RBL into {name}"
            assert curr_ticker != "RBLBANK"
        prev_run_id = curr_run_id
        prev_ticker = curr_ticker
    print(">>> Sequence Zero State Leakage test PASSED!")

if __name__ == "__main__":
    try:
        test_peer_selection_isolation()
        test_api_analyze()
        test_api_valuation()
        test_company_pipeline("TCS", "TCS", "NON_FINANCIAL")
        test_company_pipeline("Infosys", "INFY", "NON_FINANCIAL")
        test_company_pipeline("Tata Motors", ["TATAMOTORS", "TMCV"], "NON_FINANCIAL")
        test_company_pipeline("HDFC Bank", "HDFCBANK", "BANK")
        test_sequence_zero_leakage()
        print("\n=======================================================")
        print("ALL GENERALIZED TERMINAL TESTS COMPLETED SUCCESSFULLY!")
        print("=======================================================")
    except Exception as e:
        print(f"\nTEST FAILED: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
