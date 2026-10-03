import sys
from screener_client import fetch_company_data
from valuation_engine import calculate_valuation

tickers = ["SBIN", "ADANIPOWER", "OIL", "TCS", "TATASTEEL", "SUNPHARMA", "ADANIENT"]

print(f"{'Ticker':<12} | {'Sector':<18} | {'Type':<10} | {'Methodology':<25} | {'IV (Rs)':<10} | {'Status'}")
print("-" * 88)

for t in tickers:
    try:
        data = fetch_company_data(t)
        res = calculate_valuation(data)
        c_type = res.get('company_type', 'N/A')
        method = "Excess Return (Bank)" if c_type == 'BANK' else "DCF (FCFF)"
        iv = res.get('intrinsic_value_per_share', 0.0)
        status = "PASS" if iv > 0 else "FAIL"
        print(f"{t:<12} | {str(data.get('sector'))[:18]:<18} | {c_type:<10} | {method:<25} | {iv:>10.2f} | {status}")
    except Exception as e:
        print(f"{t:<12} | ERROR: {e}")

print("-" * 88)
