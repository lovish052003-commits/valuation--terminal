import requests
import os, sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from screener_client import HEADERS

# Test candidates for NESTLEIND (company_id: 2026, warehouse_id: 2026 or whatever warehouse id is)
urls = [
    'https://www.screener.in/excel/NESTLEIND/',
    'https://www.screener.in/company/NESTLEIND/excel/',
    'https://www.screener.in/company/NESTLEIND/export/',
    'https://www.screener.in/api/company/2026/export/',
    'https://www.screener.in/api/company/2026/excel/',
    'https://www.screener.in/user/export/NESTLEIND/',
    'https://www.screener.in/excel/export/NESTLEIND/',
    'https://www.screener.in/excel/company/2026/',
    'https://www.screener.in/company/NESTLEIND/export-to-excel/',
]

for u in urls:
    try:
        r = requests.get(u, headers=HEADERS, allow_redirects=False, timeout=5)
        print(f"URL: {u} -> Status: {r.status_code}, Location: {r.headers.get('Location')}, Content-Type: {r.headers.get('Content-Type')}")
    except Exception as e:
        print(f"URL: {u} -> Error: {e}")
