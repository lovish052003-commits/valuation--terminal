import sys, os
sys.path.insert(0, os.path.abspath('.'))

import pandas as pd
from screener_client import fetch_company_data
from excel_exporter import get_sector_key, get_effective_peers, is_mismatched_peer

print("=== TEST 1: ADITYA BIRLA CAPITAL (ABCAPITAL) ===")
d_ab = fetch_company_data('ABCAPITAL')
sec_ab = get_sector_key(d_ab)
print(f"Target: {d_ab['company_name']} ({d_ab['ticker']})")
print(f"Sector Key: {sec_ab}")
assert sec_ab == 'nbfc_finance', f"Expected 'nbfc_finance', got '{sec_ab}'"

peers_ab = get_effective_peers(d_ab)
print(f"\nTotal peers returned: {len(peers_ab)}")
for idx, p in enumerate(peers_ab):
    t_flag = " [TARGET COMPANY]" if p.get('is_target') else ""
    print(f"  Row {12+idx} (Index {idx}): {p['name']} | CMP: {p['cmp']} | Mcap: {p['mcap']}{t_flag}")

# Assertions for ABCAPITAL
assert peers_ab[1]['is_target'] == True, "Row 13 (Index 1) must be the target company!"
assert any(k in peers_ab[1]['name'].lower() for k in ['aditya birla', 'abcapital']), f"Row 13 must be Aditya Birla, got {peers_ab[1]['name']}"

# Verify NO commercial banks or insurance
for idx, p in enumerate(peers_ab):
    pn = p['name'].lower()
    for bad in ['hdfc bank', 'icici bank', 'state bank of india', 'sbi', 'axis bank', 'kotak mahindra bank', 'life insurance', 'sbi life']:
        assert bad not in pn or p.get('is_target'), f"Mismatched peer found: {p['name']} at index {idx}"

print("\nSUCCESS: All ABCAPITAL sector and peer assertions passed!")

print("\n=== TEST 2: SUZLON ENERGY ===")
d_suz = fetch_company_data('SUZLON')
sec_suz = get_sector_key(d_suz)
print(f"Target: {d_suz['company_name']} ({d_suz['ticker']})")
print(f"Sector Key: {sec_suz}")
assert sec_suz == 'renewable', f"Expected 'renewable', got '{sec_suz}'"

peers_suz = get_effective_peers(d_suz)
print(f"\nTotal peers returned: {len(peers_suz)}")
for idx, p in enumerate(peers_suz):
    t_flag = " [TARGET COMPANY]" if p.get('is_target') else ""
    print(f"  Row {12+idx} (Index {idx}): {p['name']} | CMP: {p['cmp']} | Mcap: {p['mcap']}{t_flag}")

assert peers_suz[1]['is_target'] == True, "Row 13 (Index 1) must be the target company!"
assert 'suzlon' in peers_suz[1]['name'].lower(), f"Row 13 must be Suzlon, got {peers_suz[1]['name']}"

# Verify NO defense or auto
for idx, p in enumerate(peers_suz):
    pn = p['name'].lower()
    for bad in ['aeronautics', 'hal', 'bel', 'tata motors', 'maruti', 'hdfc bank']:
        assert bad not in pn, f"Mismatched peer found in Suzlon: {p['name']}"

print("\nSUCCESS: All SUZLON sector and peer assertions passed!")
