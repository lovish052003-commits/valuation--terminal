import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client

tickers = ['RELIANCE', 'INFY', 'HDFCBANK', 'LT', 'SUNPHARMA', 'COALINDIA']
for t in tickers:
    try:
        d = screener_client.fetch_company_data(t)
        name = d.get('company_name')
        sec = d.get('sector')
        ind = d.get('industry')
        cmp_p = d.get('current_price')
        peers_count = len(d.get('peers_df', []))
        tables_keys = list(d.get('tables', {}).keys())
        print(f"{t}: Name={name} | Sector={sec} | Industry={ind} | CMP={cmp_p} | Peers={peers_count} | Tables={tables_keys}")
    except Exception as e:
        print(f"{t}: ERROR {e}")
