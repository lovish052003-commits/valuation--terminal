import re
import time
import requests
from lxml import html
from urllib.parse import quote

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept-Language': 'en-US,en;q=0.9',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
})

def test_fetch_slug(symbol_or_query):
    clean_query = symbol_or_query.strip()
    candidate_paths = [f"/company/{clean_query}/consolidated/"]
    
    # Try bse_code or search
    resp = None
    full_url = None
    is_consolidated = True
    
    # Live Screener search fallback
    clean_search_terms = [clean_query]
    no_ltd = re.sub(r'\b(ltd|limited|pvt|private|inc|corp)\b\.?', '', clean_query, flags=re.I).strip()
    if no_ltd and no_ltd != clean_query:
        clean_search_terms.append(no_ltd)
        
    print(f'Trying search terms: {clean_search_terms}')
    live_matches = []
    for st in clean_search_terms:
        try:
            s_url = f"https://www.screener.in/api/company/search/?q={quote(st)}"
            s_resp = session.get(s_url, timeout=(2.0, 5.0))
            if s_resp.status_code == 200:
                for item in s_resp.json():
                    if item not in live_matches:
                        live_matches.append(item)
        except Exception as e:
            print(f'Search API error: {e}')
        if live_matches:
            break

    print(f'Live matches found: {live_matches}')
    for lm in live_matches:
        u_base = lm.get('url', '')
        if not u_base:
            continue
        lm_urls = [f"https://www.screener.in{u_base}"]
        if '/consolidated/' not in u_base:
            lm_urls.insert(0, f"https://www.screener.in{u_base.rstrip('/')}/consolidated/")
        else:
            lm_urls.append(f"https://www.screener.in{u_base.replace('/consolidated/', '/')}")

        for u_try in lm_urls:
            try:
                r_try = session.get(u_try, timeout=(3.0, 8.0))
                if r_try.status_code == 200:
                    resp = r_try
                    full_url = u_try
                    is_consolidated = '/consolidated/' in u_try
                    print(f'Found working URL: {full_url} (status: {r_try.status_code})')
                    return full_url, resp
            except Exception as e:
                print(f'Error fetching {u_try}: {e}')
                
    return None, None

url, resp = test_fetch_slug('YUVRAAJ HYGIENE PRODUCTS LTD.')
if resp:
    tree = html.fromstring(resp.content)
    h1 = tree.xpath('//h1/text()')
    print('Page H1:', h1)
