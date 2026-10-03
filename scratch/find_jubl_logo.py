import requests, re

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
try:
    r = requests.get('https://www.jubilantfoodworks.com/', headers=headers, timeout=5)
    print('Status:', r.status_code)
    for m in re.finditer(r'src=["\']([^"\']+)["\']', r.text):
        src = m.group(1)
        if any(k in src.lower() for k in ['logo', 'jubilant', 'brand']):
            print('Found img on site:', src)
    for m in re.finditer(r'<link[^>]+href=["\']([^"\']+)["\']', r.text):
        href = m.group(1)
        if any(k in href.lower() for k in ['icon', 'logo', 'favicon']):
            print('Found icon link:', href)
except Exception as e:
    print('Error:', e)
