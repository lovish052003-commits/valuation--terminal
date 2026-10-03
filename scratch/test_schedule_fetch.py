import requests
from lxml import html

r = requests.get('https://www.screener.in/company/ITC/consolidated/', headers={'User-Agent': 'Mozilla/5.0'})
tree = html.fromstring(r.content)
el = tree.xpath('//*[@id="company-info"]')
if el:
    cid = el[0].attrib.get('data-company-id')
    print('company-info attrs:', el[0].attrib)
    print('Company ID:', cid)
    for parent in ['Fixed Assets', 'Other Assets', 'Borrowings', 'Other Liabilities']:
        s_url = f'https://www.screener.in/api/company/{cid}/schedules/?parent={requests.utils.quote(parent)}&section=balance-sheet&consolidated='
        s_resp = requests.get(s_url, headers={'User-Agent': 'Mozilla/5.0'})
        print(f"Schedule '{parent}': status={s_resp.status_code}, len={len(s_resp.text)}")
        if s_resp.status_code == 200:
            print(s_resp.text[:300])
