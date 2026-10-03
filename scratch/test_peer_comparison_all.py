import requests, lxml.html, io, pandas as pd

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})

tickers = ['ADANIENT', 'ITC', 'TCS', 'HINDUNILVR', 'TATASTEEL', 'RELIANCE', 'INFY', 'BHARTIARTL', 'MARUTI', 'TITAN']

for t in tickers:
    r = session.get(f'https://www.screener.in/company/{t}/consolidated/')
    tree = lxml.html.fromstring(r.text)
    m_links = [(a.text_content().strip(), a.get('href')) for a in tree.xpath('//*[@id="peers"]//a[contains(@href, "/market/")]')]
    print(f"\n=== {t} ===")
    for i, (name, href) in enumerate(m_links):
        print(f"  {i}: {name} -> {href}")
        
    # Check default api peers
    api_peers = []
    comp_id = tree.xpath('//div[@id="company-info"]/@data-company-id')
    cid = comp_id[0] if comp_id else None
    # Let's check warehouse id
    wh = tree.xpath('//div[@id="peers-table-placeholder"]/@data-warehouse-id')
    wid = wh[0] if wh else cid
    if wid:
        try:
            pr = session.get(f'https://www.screener.in/api/company/{wid}/peers/')
            if pr.status_code == 200:
                p_dfs = pd.read_html(io.StringIO(pr.text))
                if p_dfs:
                    api_peers = p_dfs[0]['Company'].dropna().tolist()
                    api_peers = [p for p in api_peers if 'median' not in p.lower()]
        except Exception as e:
            pass
    print(f"  Default API peers ({len(api_peers)}): {api_peers[:4]}")
