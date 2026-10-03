import requests, re
from lxml import html

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
})

for sym in ['RELIANCE', 'TATAMOTORS', 'SUZLON', 'SUNPHARMA']:
    for sub in ['', 'consolidated/']:
        url = f"https://www.screener.in/company/{sym}/{sub}"
        r = session.get(url)
        tree = html.fromstring(r.content)
        wids = tree.xpath('//*[@data-warehouse-id]/@data-warehouse-id')
        cids = tree.xpath('//*[@data-company-id]/@data-company-id')
        m_w = re.findall(r'data-warehouse-id=[\'"](\d+)[\'"]', r.text)
        m_c = re.findall(r'data-company-id=[\'"](\d+)[\'"]', r.text)
        m_chart = re.findall(r'/api/company/(\d+)/chart/', r.text)
        print(f"{sym} ({sub}): wids={wids[:3]}, cids={cids[:3]}, m_w={m_w[:3]}, m_c={m_c[:3]}, m_chart={m_chart[:3]}")
        
        # Test which ID works on chart API
        candidate_ids = list(set(wids + cids + m_w + m_c + m_chart))
        for cid in candidate_ids:
            c_url = f"https://www.screener.in/api/company/{cid}/chart/?q=Price-DMA50-DMA200-Volume&days=365"
            cr = session.get(c_url)
            if cr.status_code == 200:
                print(f"  -> SUCCESS! id {cid} works for {sym} chart: {len(cr.json().get('datasets', [{}])[0].get('values', []))} points")
                break
