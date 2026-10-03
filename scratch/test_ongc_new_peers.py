import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client

d = screener_client.fetch_company_data('ONGC')
print('Target:', d.get('company_name'), d.get('ticker'))
peers = d.get('sector_peers', [])
print(f"Total peers returned: {len(peers)}")
for i, p in enumerate(peers, 1):
    print(f"{i}. {p.get('name')} ({p.get('ticker')}): MCap = Rs. {p.get('market_cap'):,.1f} Cr, CMP = Rs. {p.get('current_price')}, Score = {p.get('peer_score')}")
