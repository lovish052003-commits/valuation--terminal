import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from screener_client import fetch_company_data, get_screener_session, extract_peer_table_with_tickers
from lxml import html

session = get_screener_session()
r = session.get('https://www.screener.in/company/ONGC/consolidated/')
tree = html.fromstring(r.content)
info_el = tree.xpath('//*[@id="company-info"]')
print('company-info:', info_el[0].attrib if info_el else 'None')

market_links = []
for a in tree.xpath('//*[@id="peers"]//a[contains(@href, "/market/")]'):
    market_links.append((a.text_content().strip(), a.get('href', '').strip()))
print('market_links:', market_links)

w_id = info_el[0].attrib.get('data-warehouse-id') if info_el else None
c_id = info_el[0].attrib.get('data-company-id') if info_el else None
lookup_id = w_id or c_id
print(f'lookup_id: {lookup_id}')

p_url = f"https://www.screener.in/api/company/{lookup_id}/peers/"
pr = session.get(p_url)
print(f'peers API status: {pr.status_code}')
if pr.status_code == 200:
    df = extract_peer_table_with_tickers(pr.text)
    print('Direct Peers API df length:', len(df))
    if not df.empty:
        print('Direct Peers API df Companies:', df['Company'].tolist() if 'Company' in df.columns else df.columns)

for name, href in market_links:
    mr = session.get(f"https://www.screener.in{href}")
    mdf = extract_peer_table_with_tickers(mr.text)
    print(f'Market link {name} ({href}) length: {len(mdf)}')
    if not mdf.empty and 'Company' in mdf.columns:
        print(f'  Companies: {mdf["Company"].tolist()[:5]}')
