import requests
import re
import json

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0'})
r = session.get('https://www.screener.in/company/NESTLEIND/')
m = re.search(r'data-warehouse-id=[\"\'](\d+)[\"\']', r.text)
if m:
    wid = m.group(1)
    url = f'https://www.screener.in/api/company/{wid}/schedules/?parent=Other+Assets&section=balance-sheet'
    r2 = session.get(url)
    data = r2.json()
    print("Schedules JSON keys:", data.keys() if isinstance(data, dict) else "not dict")
    print("JSON sample:", json.dumps(data, indent=2)[:1000])
