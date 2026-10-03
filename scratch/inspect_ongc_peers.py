import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client
import lxml.html

import requests
session = requests.Session()
session.headers.update(screener_client.HEADERS)
r = session.get('https://www.screener.in/company/ONGC/consolidated/', headers=screener_client.HEADERS)
tree = lxml.html.fromstring(r.content)
links = [(a.text_content().strip(), a.get('href')) for a in tree.xpath('//*[@id="peers"]//a[contains(@href, "/market/")]')]
print('Market links:', links)

for name, href in links:
    m_resp = session.get(f"https://www.screener.in{href}")
    m_tree = lxml.html.fromstring(m_resp.content)
    df = screener_client.extract_peer_table_with_tickers(m_resp.text)
    print(f"\n--- {name} ({href}) ---")
    if not df.empty:
        print("Cols:", df.columns.tolist())
        comp_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
        for _, row in df.head(10).iterrows():
            print(f"  {row.get(comp_col)} ({row.get('Ticker')}): MCap={row.get('Mar Cap  Rs.Cr.') or row.get('Mar Cap Rs.Cr.')}")
