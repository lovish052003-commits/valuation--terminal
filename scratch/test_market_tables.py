import requests, lxml.html, io, pandas as pd

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})

tickers = ['ADANIENT', 'ITC', 'TCS', 'HINDUNILVR', 'TATASTEEL', 'RELIANCE', 'MARUTI', 'SUNPHARMA', 'HDFCBANK', 'LT']

for t in tickers:
    r = session.get(f'https://www.screener.in/company/{t}/consolidated/')
    tree = lxml.html.fromstring(r.text)
    m_links = [(a.text_content().strip(), a.get('href')) for a in tree.xpath('//*[@id="peers"]//a[contains(@href, "/market/")]')]
    print(f"\n=================== {t} ===================")
    # Test link index 1, and index 2 if available
    for idx in [1, 2]:
        if idx < len(m_links):
            name, href = m_links[idx]
            resp = session.get('https://www.screener.in' + href)
            dfs = pd.read_html(io.StringIO(resp.text))
            if dfs:
                df = dfs[0]
                comps = df['Company'].dropna().tolist()[:6]
                print(f"[{idx}] {name} ({href}): {len(df)} companies -> {comps}")
