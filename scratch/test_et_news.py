import requests
import xml.etree.ElementTree as ET
import re

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
})

def fetch_recent_updates(company_name, ticker=None):
    query = f"{company_name} Economic Times"
    url = f"https://news.google.com/rss/search?q={requests.utils.quote(query)}&hl=en-IN&gl=IN&ceid=IN:en"
    
    r = session.get(url, timeout=10)
    root = ET.fromstring(r.text)
    items = []
    for item in root.findall('.//item'):
        title_el = item.find('title')
        desc_el = item.find('description')
        title = title_el.text if title_el is not None else ""
        desc = desc_el.text if desc_el is not None else ""
        # Clean html from desc
        clean_desc = re.sub(r'<[^>]+>', '', desc).strip()
        # Remove source suffix like " - The Economic Times"
        clean_title = re.sub(r'\s*-\s*(The\s*)?Economic Times.*$', '', title, flags=re.IGNORECASE).strip()
        if clean_title:
            items.append((clean_title, clean_desc))
            if len(items) >= 8:
                break
    return items

print("Adani Enterprises news:")
for i, (t, d) in enumerate(fetch_recent_updates("Adani Enterprises"), 1):
    print(f"\n{i}. Title: {t}\n   Desc: {d}")
