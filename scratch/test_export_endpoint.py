import requests
import os, sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from screener_client import HEADERS

url = 'https://www.screener.in/user/company/export/331/'
r = requests.get(url, headers=HEADERS, allow_redirects=False)
print("Without cookies: Status:", r.status_code, "Location:", r.headers.get('Location'))
