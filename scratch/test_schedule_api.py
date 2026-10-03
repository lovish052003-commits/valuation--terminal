import requests, re
from lxml import html

HEADERS = {'User-Agent': 'Mozilla/5.0'}
r = requests.get('https://www.screener.in/company/ITC/consolidated/', headers=HEADERS)

tree = html.fromstring(r.content)
scripts = tree.xpath('//script/@src')
print('Script src:', scripts)

inline = "".join(tree.xpath('//script[not(@src)]/text()'))
for line in inline.split('\n'):
    if 'showSchedule' in line or 'schedule' in line.lower():
        print('Inline match:', line.strip())

# Also check endpoint: /api/company/{company_id}/schedules/
# ITC warehouse id is 1285768 or check data-warehouse-id
m = re.search(r'data-warehouse-id="(\d+)"', r.text)
if m:
    wid = m.group(1)
    print('Warehouse ID:', wid)
    # Test schedule API
    for sec in ['Fixed Assets', 'Other Assets', 'Borrowings', 'Other Liabilities']:
        test_url = f'https://www.screener.in/api/company/{wid}/schedules/?parent={requests.utils.quote(sec)}&section=balance-sheet'
        s_resp = requests.get(test_url, headers=HEADERS)
        print(f'Schedule URL {sec}: status={s_resp.status_code}, len={len(s_resp.text)}')
        if s_resp.status_code == 200:
            print(s_resp.text[:200])

