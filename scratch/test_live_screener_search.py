import requests
import re
from urllib.parse import quote

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})

query = 'YUVRAAJ HYGIENE PRODUCTS LTD.'
print('Testing resolution for:', query)

# 1. Screener live search
clean_q = re.sub(r'\b(ltd|limited|pvt|private|inc|corp)\b\.?', '', query, flags=re.I).strip()
search_urls = [
    f'https://www.screener.in/api/company/search/?q={quote(query)}',
    f'https://www.screener.in/api/company/search/?q={quote(clean_q)}'
]

matches = []
for su in search_urls:
    try:
        r = session.get(su, timeout=5)
        if r.status_code == 200:
            for item in r.json():
                if item not in matches:
                    matches.append(item)
    except Exception as e:
        print('Search error:', e)

print('Matches found:', matches)
for m in matches:
    u = m['url']
    if not u.endswith('/consolidated/'):
        test_urls = [f'https://www.screener.in{u}consolidated/', f'https://www.screener.in{u}']
    else:
        test_urls = [f'https://www.screener.in{u}', f'https://www.screener.in{u.replace("/consolidated/", "/")}']
    for tu in test_urls:
        resp = session.get(tu, timeout=5)
        print('  Testing:', tu, '-> status:', resp.status_code)
        if resp.status_code == 200:
            print('  SUCCESS! Found company page at:', tu)
            break
