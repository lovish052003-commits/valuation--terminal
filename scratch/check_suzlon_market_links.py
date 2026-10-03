import sys, os
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__) + '/..'))
import requests, lxml.html
from screener_client import get_screener_session

session = get_screener_session()
r = session.get('https://www.screener.in/company/SUZLON/consolidated/')
tree = lxml.html.fromstring(r.content)
market_links = []
for a in tree.xpath('//*[@id="peers"]//a[contains(@href, "/market/")]'):
    t = a.text_content().strip()
    h = a.get('href', '').strip()
    if t and h:
        market_links.append((t, h))
print('Market links for Suzlon:', market_links)
