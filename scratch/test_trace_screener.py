import requests, re
from lxml import html

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
})

url = 'https://www.screener.in/company/RELIANCE/consolidated/'
resp = session.get(url)
print('Company page status:', resp.status_code)

tree = html.fromstring(resp.content)
company_id = None
info_el = tree.xpath('//*[@id="company-info"]')
if info_el:
    company_id = info_el[0].attrib.get('data-warehouse-id')
if not company_id:
    m_id = re.search(r'data-company-id=[\'"](\d+)[\'"]', resp.text)
    if m_id:
        company_id = m_id.group(1)

print('Extracted company_id:', company_id)

chart_url = f'https://www.screener.in/api/company/{company_id}/chart/?q=Price-DMA50-DMA200-Volume&days=3650'
c_resp = session.get(chart_url, timeout=10)
print('Chart 10y status:', c_resp.status_code, 'len:', len(c_resp.text))

chart_1y_url = f'https://www.screener.in/api/company/{company_id}/chart/?q=Price-DMA50-DMA200-Volume&days=365'
c1_resp = session.get(chart_1y_url, timeout=10)
print('Chart 1y status:', c1_resp.status_code, 'len:', len(c1_resp.text))
if c1_resp.status_code == 200:
    c1_data = c1_resp.json()
    for ds in c1_data.get('datasets', []):
        print('dataset metric:', ds.get('metric'), 'values len:', len(ds.get('values', [])))
