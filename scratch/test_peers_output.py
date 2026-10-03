import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client

data = screener_client.fetch_company_data('HINDUNILVR')
sp = data.get('sector_peers', [])
print(f"Found sector_peers: {len(sp)}")
for p in sp:
    tick = p.get('ticker') or ''
    nm = p.get('name') or ''
    mcap = p.get('market_cap', 0)
    debt = p.get('debt', 0)
    cash = p.get('cash', 0)
    ev = p.get('ev', 0)
    print(f"{tick:10s} | {nm:20s} | Mcap: {mcap:10.1f} | Debt: {debt:8.1f} | Cash: {cash:8.1f} | EV: {ev:10.1f}")
