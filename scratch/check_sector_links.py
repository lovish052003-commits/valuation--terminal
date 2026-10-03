import requests
import lxml.html

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})

for sym in ['ITC', 'TCS', 'HINDUNILVR', 'TATASTEEL', 'RELIANCE', 'ADANIENT']:
    r = session.get(f'https://www.screener.in/company/{sym}/consolidated/')
    tree = lxml.html.fromstring(r.text)
    links = [(a.text_content().strip(), a.get('href')) for a in tree.xpath('//*[@id="peers"]//a[contains(@href, "/market/")]')]
    print(f"\n{sym}:")
    for name, href in links:
        print(f"  {name} -> {href}")
