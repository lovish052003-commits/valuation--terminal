import requests
import re
import json

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0'})
r = session.get('https://www.screener.in/company/NESTLEIND/consolidated/')
m = re.search(r'data-warehouse-id=[\"\'](\d+)[\"\']', r.text)
if not m:
    r = session.get('https://www.screener.in/company/NESTLEIND/')
    m = re.search(r'data-warehouse-id=[\"\'](\d+)[\"\']', r.text)

if m:
    wid = m.group(1)
    print("Nestle Warehouse ID:", wid)
    url = f'https://www.screener.in/api/company/{wid}/schedules/?parent=Other+Assets&section=balance-sheet'
    r2 = session.get(url)
    print("Status:", r2.status_code)
    try:
        data = r2.json()
        print("Schedule JSON for Other Assets:")
        print(json.dumps(data, indent=2))
    except Exception as e:
        print("JSON parse error:", e, r2.text[:200])
