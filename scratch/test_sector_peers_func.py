import requests
import lxml.html
import io
import pandas as pd
import re

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
})

def clean_num(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace('\u20b9', '').replace(',', '').replace('%', '').strip()
    m = re.search(r'[-+]?\d+(?:\.\d+)?', s)
    return float(m.group(0)) if m else 0.0

def is_company_match(cell_text, target_name, ticker):
    c = str(cell_text or '').lower().strip()
    t_name = str(target_name or '').lower().strip()
    t_sym = str(ticker or '').lower().strip()
    if not c:
        return False
    if t_sym and (c == t_sym or t_sym in c.split()):
        return True
    clean_c = re.sub(r'[^a-z0-9]', '', c)
    clean_target = re.sub(r'[^a-z0-9]', '', t_name)
    if clean_target and (clean_target in clean_c or clean_c in clean_target):
        return True
    first_w = t_name.split()[0] if t_name else ''
    if len(first_w) >= 4 and first_w in c:
        return True
    return False

def fetch_sector_peers_table(tree, company_name, ticker, session, warehouse_id=None, company_id=None):
    market_links = []
    for a in tree.xpath('//*[@id="peers"]//a[contains(@href, "/market/")]'):
        t = a.text_content().strip()
        h = a.get('href', '').strip()
        if t and h:
            market_links.append((t, h))
            
    chosen_link = None
    if len(market_links) > 1:
        chosen_link = market_links[1]
    elif market_links:
        chosen_link = market_links[0]
        
    df_result = pd.DataFrame()
    if chosen_link:
        name, href = chosen_link
        try:
            m_resp = session.get(f"https://www.screener.in{href}", timeout=10)
            if m_resp.status_code == 200:
                dfs = pd.read_html(io.StringIO(m_resp.text), flavor='lxml')
                if dfs:
                    df_raw = dfs[0]
                    # Drop median rows or blank companies
                    df_clean = df_raw[~df_raw['Company'].astype(str).str.contains('Median', case=False, na=False)].copy()
                    df_clean = df_clean.dropna(subset=['Company'])
                    
                    # Find target company in the dataframe
                    target_idx = None
                    for idx, (_, row) in enumerate(df_clean.iterrows()):
                        if is_company_match(row.get('Company'), company_name, ticker):
                            target_idx = idx
                            break
                            
                    if target_idx is not None:
                        if target_idx < 10:
                            df_result = df_clean.head(10).copy()
                        else:
                            # Include top 9 + target company
                            top_9 = df_clean.head(9)
                            target_row = df_clean.iloc[[target_idx]]
                            df_result = pd.concat([top_9, target_row], ignore_index=True)
                    else:
                        df_result = df_clean.head(10).copy()
        except Exception as e:
            print(f"Notice: Could not fetch sector peers from {chosen_link}: {e}")
            
    # Fallback to default API if needed
    if df_result.empty and (warehouse_id or company_id):
        try:
            lookup_id = warehouse_id or company_id
            p_url = f"https://www.screener.in/api/company/{lookup_id}/peers/"
            pr = session.get(p_url, timeout=10)
            if pr.status_code == 200:
                p_dfs = pd.read_html(io.StringIO(pr.text), flavor='lxml')
                if p_dfs:
                    df_result = p_dfs[0][~p_dfs[0]['Company'].astype(str).str.contains('Median', case=False, na=False)].head(10).copy()
        except Exception as e:
            print(f"Fallback peers API notice: {e}")
            
    return df_result

# Test for Adani Enterprises
resp = session.get('https://www.screener.in/company/ADANIENT/consolidated/')
tree = lxml.html.fromstring(resp.content)
df_adani = fetch_sector_peers_table(tree, "Adani Enterprises Ltd", "ADANIENT", session)
print("Adani Peers (Count:", len(df_adani), "):")
for i, r in df_adani.iterrows():
    print(f"  {i+1}. {r.get('Company')} | Mcap: {r.get('Mar Cap  Rs.Cr.')} | P/E: {r.get('P/E')}")
