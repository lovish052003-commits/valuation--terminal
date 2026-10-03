import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import re
import pandas as pd
from io import StringIO
import lxml.html
from screener_client import get_screener_session

def test_extract():
    session = get_screener_session()
    r = session.get('https://www.screener.in/company/ADANIPOWER/consolidated/', timeout=10)
    tree = lxml.html.fromstring(r.text)
    info_el = tree.xpath('//*[@id="company-info"]')
    wid = info_el[0].attrib.get('data-warehouse-id') if info_el else None
    print('wid:', wid)
    if wid:
        p_url = f'https://www.screener.in/api/company/{wid}/peers/'
        pr = session.get(p_url, timeout=10)
        p_dfs = pd.read_html(StringIO(pr.text), flavor='lxml')
        df = p_dfs[0]
        df_clean = df[~df['Company'].astype(str).str.contains('Median', case=False, na=False)].copy()
        
        ticker_map = {}
        tree = lxml.html.fromstring(pr.text)
        for tr in tree.xpath('//tr'):
            links = tr.xpath('.//a[contains(@href, "/company/")]')
            if links:
                a = links[0]
                c_name = a.text_content().strip()
                href = a.get('href', '')
                m = re.search(r'/company/([^/]+)/', href)
                if m and c_name:
                    ticker_map[c_name] = m.group(1).strip().upper()
        
        df_clean['Ticker'] = df_clean['Company'].map(lambda x: ticker_map.get(str(x).strip(), ''))
        print(df_clean[['Company', 'Ticker', 'CMP  Rs.', 'Mar Cap  Rs.Cr.']])

if __name__ == '__main__':
    test_extract()
