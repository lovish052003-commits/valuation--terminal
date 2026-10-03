import sys, os
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__) + '/..'))
import requests, lxml.html
from io import StringIO
import pandas as pd
from screener_client import get_screener_session

session = get_screener_session()
r = session.get('https://www.screener.in/company/TATAMOTORS/consolidated/')
tree = lxml.html.fromstring(r.content)
market_links = []
for a in tree.xpath('//*[@id="peers"]//a[contains(@href, "/market/")]'):
    t = a.text_content().strip()
    h = a.get('href', '').strip()
    if t and h:
        market_links.append((t, h))
print('Tata Motors market links:', market_links)

w_meta = tree.xpath('//*[@data-warehouse-id]')
cid = w_meta[0].get('data-warehouse-id') if w_meta else None
print('Warehouse ID:', cid)
if cid:
    pr = session.get(f"https://www.screener.in/api/company/{cid}/peers/")
    p_dfs = pd.read_html(StringIO(pr.text), flavor='lxml')
    if p_dfs:
        print("\n--- Direct Peers API for Tata Motors ---")
        for comp in p_dfs[0]['Company'].dropna().tolist()[:10]:
            print(" ", comp)
