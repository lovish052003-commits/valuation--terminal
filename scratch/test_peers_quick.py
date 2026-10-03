import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding='utf-8')
import screener_client

for ticker in ['SBIN', 'ITC', 'TCS', 'LICI', 'ICICIAMC']:
    try:
        data = screener_client.fetch_company_data(ticker)
        print(f"=== {ticker}: {data.get('company_name')} ({data.get('company_type')}) ===")
        print(f"Price: {data.get('current_price')}, MCap: {data.get('market_cap_cr')}, Sector: {data.get('sector')}, Industry: {data.get('industry')}")
        peers = data.get('sector_peers', [])
        print(f"Sector Peers count: {len(peers)}")
        for p in peers[:5]:
            print(f"  - {p.get('name')} ({p.get('ticker')}): CMP {p.get('current_price')}, MCap {p.get('market_cap')}")
    except Exception as e:
        print(f"=== {ticker} FAILED: {e} ===")
