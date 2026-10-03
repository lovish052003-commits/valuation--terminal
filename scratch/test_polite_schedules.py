import sys, os, time, re, requests
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

session = requests.Session()
retries = Retry(total=5, backoff_factor=1.0, status_forcelist=[429, 500, 502, 503, 504])
session.mount('https://', HTTPAdapter(max_retries=retries))
session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})

for sym in ['WIPRO', 'BRITANNIA']:
    resp = session.get(f'https://www.screener.in/company/{sym}/consolidated/')
    if resp.status_code != 200:
        resp = session.get(f'https://www.screener.in/company/{sym}/')
    m_id = re.search(r'data-company-id=[\'"](\d+)[\'"]', resp.text)
    cid = m_id.group(1) if m_id else None
    print(f"{sym} -> Company ID: {cid}")
    
    sch_names = [
        ('Fixed Assets', 'balance-sheet'),
        ('Other Assets', 'balance-sheet'),
        ('Borrowings', 'balance-sheet'),
        ('Other Liabilities', 'balance-sheet'),
        ('Cash from Operating Activity', 'cash-flow'),
        ('Cash from Investing Activity', 'cash-flow'),
        ('Cash from Financing Activity', 'cash-flow'),
        ('Expenses', 'profit-loss'),
        ('Material Cost %', 'profit-loss')
    ]
    for parent, sec in sch_names:
        time.sleep(0.3)
        url = f'https://www.screener.in/api/company/{cid}/schedules/?parent={requests.utils.quote(parent)}&section={sec}&consolidated='
        r = session.get(url, timeout=10)
        is_j = r.text.startswith('{')
        print(f"  {parent}: status={r.status_code}, is_json={is_j}, len={len(r.text)}")
