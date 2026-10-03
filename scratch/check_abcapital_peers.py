import sys, os
sys.path.insert(0, os.path.abspath('.'))
from screener_client import fetch_company_data, get_screener_session
from lxml import html
import pandas as pd
from io import StringIO

session = get_screener_session()
r = session.get('https://www.screener.in/company/ABCAPITAL/consolidated/')
t = html.fromstring(r.content)

print("Sector links:")
for a in t.xpath('//*[@id="peers"]//a'):
    print(a.text_content().strip(), "-->", a.get('href'))

# Also check market links
market_links = []
for a in t.xpath('//*[@id="peers"]//a[contains(@href, "/market/")]'):
    market_links.append((a.text_content().strip(), a.get('href')))
print("market_links:", market_links)

for name, href in market_links:
    m_resp = session.get(f"https://www.screener.in{href}")
    dfs = pd.read_html(StringIO(m_resp.text), flavor='lxml')
    if dfs:
        print(f"--- Market Link: {name} ({href}) ---")
        print(dfs[0]['Company'].head(12).tolist())
