import requests
from lxml import html

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0'})

for u in ['https://www.screener.in/company/531663/consolidated/', 'https://www.screener.in/company/531663/']:
    r = session.get(u)
    print(f'URL: {u} (status {r.status_code})')
    tree = html.fromstring(r.content)
    meta = {}
    for li in tree.xpath('//ul[@id="top-ratios"]//li'):
        name = li.xpath('.//span[@class="name"]/text()')
        val = li.xpath('.//span[@class="nowrap value"]//text()')
        if name and val:
            k = name[0].strip()
            v = "".join(val).strip().replace('\u20b9', 'Rs ')
            meta[k] = v
    print('  Meta:', meta)
