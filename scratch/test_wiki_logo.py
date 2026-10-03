import requests
from lxml import html
from PIL import Image
from io import BytesIO

def get_wiki_infobox_logo(title):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    url = f"https://en.wikipedia.org/wiki/{title.replace(' ', '_')}"
    try:
        r = requests.get(url, headers=headers, timeout=5)
        if r.status_code == 200:
            tree = html.fromstring(r.content)
            logos = tree.xpath('//td[contains(@class, "infobox-image")]//img/@src') or \
                    tree.xpath('//table[contains(@class, "infobox")]//tr[1]//img/@src') or \
                    tree.xpath('//table[contains(@class, "infobox")]//img[contains(@alt, "Logo") or contains(@alt, "logo")]/@src')
            if logos:
                src = logos[0]
                if src.startswith('//'):
                    src = 'https:' + src
                # Convert thumbnail to higher resolution if needed
                return src
    except Exception as e:
        print(f"Error for {title}: {e}")
    return None

test_companies = ['Tata Steel', 'Reliance Industries', 'State Bank of India', 'ITC Limited', 'JSW Steel', 'Hindustan Unilever']
for c in test_companies:
    src = get_wiki_infobox_logo(c)
    print(f"{c:25} -> {src}")
