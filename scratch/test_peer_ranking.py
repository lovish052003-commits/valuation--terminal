import sys, os, math
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import requests
import lxml.html
import screener_client
from excel_exporter import is_same_company

session = requests.Session()
session.headers.update(screener_client.HEADERS)

def test_peer_pool_and_ranking(ticker):
    print(f"\n=======================================================")
    print(f"Testing Peer Pool & Ranking for: {ticker}")
    print(f"=======================================================")
    
    url = f"https://www.screener.in/company/{ticker}/consolidated/"
    r = session.get(url)
    if r.status_code != 200:
        url = f"https://www.screener.in/company/{ticker}/"
        r = session.get(url)
    tree = lxml.html.fromstring(r.content)
    
    # Target info
    target_name = tree.xpath('//h1/text()')[0].strip() if tree.xpath('//h1/text()') else ticker
    target_mcap_el = tree.xpath('//li[contains(., "Market Cap")]//span[@class="number"]')
    target_mcap = screener_client.clean_num(target_mcap_el[0].text_content()) if target_mcap_el else 100000.0
    target_type = screener_client.classify_company('', '', target_name)
    
    print(f"Target: {target_name} ({ticker}) | MCap: Rs. {target_mcap:,.1f} Cr | Type: {target_type}")

    # Gather candidate peers from all market links and direct peers API
    market_links = [(a.text_content().strip(), a.get('href')) for a in tree.xpath('//*[@id="peers"]//a[contains(@href, "/market/")]')]
    print("Market links in hierarchy:", market_links)
    
    candidates = {} # ticker/name -> dict

    # 1. Direct peers API
    wh_ids = tree.xpath('//div[@id="company-info"]/@data-warehouse-id') or tree.xpath('//@data-warehouse-id')
    if wh_ids:
        p_url = f"https://www.screener.in/api/company/{wh_ids[0]}/peers/"
        pr = session.get(p_url)
        if pr.status_code == 200:
            df = screener_client.extract_peer_table_with_tickers(pr.text)
            if not df.empty:
                for _, row in df.iterrows():
                    name = row.get('Company')
                    tkr = row.get('Ticker')
                    mcap = screener_client.clean_num(row.get('Mar Cap  Rs.Cr.') or row.get('Mar Cap Rs.Cr.'))
                    cmp_v = screener_client.clean_num(row.get('CMP  Rs.') or row.get('CMP Rs.'))
                    if name and not is_same_company(tkr, name, ticker, target_name):
                        candidates[tkr or name] = {
                            'name': name,
                            'ticker': tkr,
                            'mcap': mcap,
                            'cmp': cmp_v,
                            'source_level': 'direct',
                            'level_depth': 4
                        }

    # 2. Market links (from sub-industry to sector)
    for depth, (lvl_name, href) in enumerate(reversed(market_links), 1):
        # depth 1 = sub-industry, depth 2 = industry, depth 3 = industry group, depth 4 = sector
        m_resp = session.get(f"https://www.screener.in{href}")
        if m_resp.status_code == 200:
            df = screener_client.extract_peer_table_with_tickers(m_resp.text)
            if not df.empty:
                for _, row in df.iterrows():
                    name = row.get('Company')
                    tkr = row.get('Ticker')
                    mcap = screener_client.clean_num(row.get('Mar Cap  Rs.Cr.') or row.get('Mar Cap Rs.Cr.'))
                    cmp_v = screener_client.clean_num(row.get('CMP  Rs.') or row.get('CMP Rs.'))
                    key = tkr or name
                    if name and not is_same_company(tkr, name, ticker, target_name):
                        if key not in candidates:
                            candidates[key] = {
                                'name': name,
                                'ticker': tkr,
                                'mcap': mcap,
                                'cmp': cmp_v,
                                'source_level': lvl_name,
                                'level_depth': depth
                            }

    print(f"Total raw candidates collected: {len(candidates)}")

    # 3. Score and rank candidates
    # peerScore = industryMatch * 40 + businessModelMatch * 25 + sectorMatch * 15 + sizeSimilarity * 10 + financialSimilarity * 10
    scored_peers = []
    for key, c in candidates.items():
        # Quality filter
        if c['cmp'] <= 0 or c['mcap'] <= 0:
            continue
        
        # Level depth: 1 = exact sub-industry (40 pts), 2 = industry (30 pts), 3 = group (20 pts), 4 = sector (10 pts)
        depth = c['level_depth']
        if depth == 1:
            industry_match = 1.0
        elif depth == 2:
            industry_match = 0.75
        elif depth == 3:
            industry_match = 0.50
        else:
            industry_match = 0.25

        c_type = screener_client.classify_company('', '', c['name'])
        biz_match = 1.0 if c_type == target_type else (0.5 if 'FINANCIAL' not in c_type and 'FINANCIAL' not in target_type else 0.0)
        sector_match = 1.0 if depth <= 3 else 0.7

        # Size similarity (log-scale)
        if target_mcap > 0 and c['mcap'] > 0:
            log_diff = abs(math.log10(c['mcap'] / target_mcap))
            size_sim = max(0.0, 1.0 - (log_diff / 2.0))
        else:
            size_sim = 0.5

        # Market-cap filter: prefer 0.1x to 10x target_mcap
        size_penalty = 0.0
        ratio = c['mcap'] / target_mcap if target_mcap > 0 else 1.0
        if ratio < 0.05:
            size_penalty = 25.0 # heavily penalize microcaps for large targets
        elif ratio < 0.1:
            size_penalty = 10.0
        elif ratio > 15.0:
            size_penalty = 5.0

        score = (
            industry_match * 40.0
            + biz_match * 25.0
            + sector_match * 15.0
            + size_sim * 10.0
            + 10.0 # baseline financial data presence
            - size_penalty
        )

        c['score'] = round(score, 1)
        c['ratio'] = round(ratio, 2)
        scored_peers.append(c)

    scored_peers.sort(key=lambda x: x['score'], reverse=True)
    
    print("\nTop 10 Ranked Peers:")
    for i, p in enumerate(scored_peers[:10], 1):
        print(f" {i:2d}. {p['name']:<30} ({p['ticker']:<10}): MCap=Rs. {p['mcap']:>10,.1f} Cr ({p['ratio']:>5.2f}x) | Score={p['score']:>5.1f} | Level={p['source_level']}")

if __name__ == '__main__':
    test_peer_pool_and_ranking('ONGC')
    test_peer_pool_and_ranking('TCS')
