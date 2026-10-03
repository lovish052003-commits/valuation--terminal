import requests, json
from lxml import html

HEADERS = {'User-Agent': 'Mozilla/5.0'}
r = requests.get('https://www.screener.in/company/TMCV/consolidated/', headers=HEADERS)
tree = html.fromstring(r.content)
info_el = tree.xpath('//*[@id="company-info"]')
if info_el:
    cid = info_el[0].attrib.get('data-company-id')
    is_cons = 'data-consolidated' in info_el[0].attrib
    print('Tata Motors cid:', cid, 'is_cons:', is_cons)
    for parent in ['Fixed Assets', 'Other Assets', 'Borrowings', 'Cash from Operating Activity']:
        sec = 'cash-flow' if 'Cash' in parent else 'balance-sheet'
        cons_p = '&consolidated=' if is_cons else ''
        url = f'https://www.screener.in/api/company/{cid}/schedules/?parent={requests.utils.quote(parent)}&section={sec}{cons_p}'
        res = requests.get(url, headers=HEADERS)
        print(f"  {parent}: status={res.status_code}, keys={list(res.json().keys()) if res.status_code==200 and res.text.startswith('{') else res.status_code}")
