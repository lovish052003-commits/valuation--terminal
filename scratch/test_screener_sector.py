import requests
from lxml import html

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0'})
for sym in ['ADANIENT', 'ITC', 'INFY', 'TATAMOTORS', 'SUNPHARMA']:
    r = session.get(f'https://www.screener.in/company/{sym}/consolidated/', timeout=10)
    tree = html.fromstring(r.content)
    m_links = tree.xpath('//*[@id="peers"]//a[contains(@href, "/market/")]/text()')
    m_clean = [x.strip() for x in m_links if x.strip()]
    print(f'{sym:12s} -> Market hierarchy: {m_clean}')
