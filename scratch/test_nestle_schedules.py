import requests
import json

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
})

import sys
import os
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__) + '/..'))
from screener_client import fetch_single_schedule

company_id = "4019" # Let's find company_id for NESTLEIND
r = session.get('https://www.screener.in/company/NESTLEIND/')
import re
m_id = re.search(r'data-company-id=[\'"](\d+)[\'"]', r.text)
if m_id:
    company_id = m_id.group(1)
print(f"Company ID: {company_id}")

print("\nTesting schedule with consolidated=True:")
sch_cons = fetch_single_schedule(company_id, 'Other Assets', 'balance-sheet', is_consolidated=True, session=session)
print("Consolidated Other Assets keys:", sch_cons.keys() if sch_cons else None)
if sch_cons and 'Trade receivables' in sch_cons:
    print("Trade receivables:", sch_cons['Trade receivables'])

print("\nTesting schedule with consolidated=False:")
sch_std = fetch_single_schedule(company_id, 'Other Assets', 'balance-sheet', is_consolidated=False, session=session)
print("Standalone Other Assets keys:", sch_std.keys() if sch_std else None)
if sch_std and 'Trade receivables' in sch_std:
    print("Trade receivables:", sch_std['Trade receivables'])
