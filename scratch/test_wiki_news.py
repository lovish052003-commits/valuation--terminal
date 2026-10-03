import requests
import json
import xml.etree.ElementTree as ET

def get_wikipedia_about(company_name):
    session = requests.Session()
    session.headers.update({'User-Agent': 'ValuationPlatform/1.0 (contact@example.com)'})
    
    # 1. Try opensearch to get exact page title
    search_url = f"https://en.wikipedia.org/w/api.php?action=opensearch&search={requests.utils.quote(company_name)}&limit=3&namespace=0&format=json"
    res = session.get(search_url, timeout=5).json()
    if len(res) > 1 and res[1]:
        title = res[1][0]
        summary_url = f"https://en.wikipedia.org/api/rest_v1/page/summary/{requests.utils.quote(title)}"
        s_res = session.get(summary_url, timeout=5)
        if s_res.status_code == 200:
            extract = s_res.json().get('extract', '')
            if extract:
                return extract
    return None

def get_recent_news(company_name):
    session = requests.Session()
    session.headers.update({'User-Agent': 'Mozilla/5.0'})
    # Google News RSS for Economic Times
    rss_url = f"https://news.google.com/rss/search?q={requests.utils.quote(company_name + ' site:economictimes.indiatimes.com')}&hl=en-IN&gl=IN&ceid=IN:en"
    r = session.get(rss_url, timeout=5)
    root = ET.fromstring(r.text)
    items = []
    for item in root.findall('.//item')[:5]:
        title = item.find('title').text if item.find('title') is not None else ''
        desc = item.find('description').text if item.find('description') is not None else ''
        items.append(title)
    return items

print("Testing ADANIENT:")
print("Wiki:", get_wikipedia_about("Adani Enterprises"))
print("\nET News:", get_recent_news("Adani Enterprises"))
