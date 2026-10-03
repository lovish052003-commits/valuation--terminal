import sys, os
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__) + '/..'))
import requests, lxml.html
from io import StringIO
import pandas as pd
from screener_client import get_screener_session

session = get_screener_session()
r = session.get('https://www.screener.in/company/SUZLON/consolidated/')
tree = lxml.html.fromstring(r.content)

# Check all market links
for name, href in [('Industrials', '/market/IN07/'), ('Capital Goods', '/market/IN07/IN0702/'), ('Electrical Equipment', '/market/IN07/IN0702/IN070203/'), ('Heavy Electrical Equipment', '/market/IN07/IN0702/IN070203/IN070203001/')]:
    resp = session.get(f"https://www.screener.in{href}")
    dfs = pd.read_html(StringIO(resp.text), flavor='lxml')
    if dfs:
        print(f"\n--- {name} ({href}) ---")
        for comp in dfs[0]['Company'].dropna().tolist()[:6]:
            print(" ", comp)

# Check API peers
cid = None
w_meta = tree.xpath('//*[@data-warehouse-id]')
if w_meta:
    cid = w_meta[0].get('data-warehouse-id')
print(f"\nWarehouse ID: {cid}")
if cid:
    pr = session.get(f"https://www.screener.in/api/company/{cid}/peers/")
    p_dfs = pd.read_html(StringIO(pr.text), flavor='lxml')
    if p_dfs:
        print("\n--- Direct Peers API ---")
        for comp in p_dfs[0]['Company'].dropna().tolist()[:10]:
            print(" ", comp)
