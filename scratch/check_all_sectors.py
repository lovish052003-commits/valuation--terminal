import requests, lxml.html

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})

tickers = ['ADANIENT', 'ITC', 'TCS', 'MARUTI', 'SUNPHARMA', 'HDFCBANK', 'LT', 'TATASTEEL']
for t in tickers:
    r = session.get(f'https://www.screener.in/company/{t}/consolidated/')
    tree = lxml.html.fromstring(r.text)
    links = [(a.text_content().strip(), a.get('href')) for a in tree.xpath('//*[@id="peers"]//a[contains(@href, "/market/")]')]
    print(f"\n{t}:")
    for idx, (name, href) in enumerate(links):
        print(f"  [{idx}] {name} -> {href}")
