import requests
import lxml.html
import io
import pandas as pd
import re

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
})

def fetch_sector_peers_test(symbol):
    url = f"https://www.screener.in/company/{symbol}/consolidated/"
    r = session.get(url, timeout=10)
    if r.status_code != 200:
        url = f"https://www.screener.in/company/{symbol}/"
        r = session.get(url, timeout=10)
        
    tree = lxml.html.fromstring(r.text)
    market_links = []
    for a in tree.xpath('//*[@id="peers"]//a[contains(@href, "/market/")]'):
        t = a.text_content().strip()
        h = a.get('href', '').strip()
        if t and h:
            market_links.append((t, h))
            
    print(f"\n================ Symbol: {symbol} ================")
    print("Market Links:", market_links)
    
    # Target industry link: index 1 is industry group; fallback to index 0
    target_link = None
    if len(market_links) > 1:
        target_link = market_links[1]
    elif market_links:
        target_link = market_links[0]
        
    if not target_link:
        print("No market link found")
        return None
        
    name, href = target_link
    print(f"Selected Sector / Industry: {name} -> https://www.screener.in{href}")
    m_resp = session.get(f"https://www.screener.in{href}", timeout=10)
    dfs = pd.read_html(io.StringIO(m_resp.text), flavor='lxml')
    if dfs:
        df = dfs[0]
        print(f"Fetched {len(df)} peers from {name}:")
        for i, row in df.head(10).iterrows():
            c_name = row.get('Company')
            mcap = row.get('Mar Cap  Rs.Cr.') or row.get('Mar Cap Rs.Cr.')
            pe = row.get('P/E')
            print(f"  {i+1}. {c_name} | Mcap: {mcap} Cr | P/E: {pe}")
        return df
    return None

for sym in ['ADANIENT', 'TATASTEEL', 'ITC', 'TCS', 'NESTLEIND']:
    fetch_sector_peers_test(sym)
