import requests
from lxml import html

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
})

r = session.get('https://www.screener.in/company/NESTLEIND/')
tree = html.fromstring(r.content)
for a in tree.xpath('//a'):
    href = a.get('href', '')
    text = a.text_content().strip()
    if 'excel' in href.lower() or 'export' in href.lower() or 'excel' in text.lower() or 'export' in text.lower():
        print(f"Text: {text} | Href: {href}")

# Also check buttons or forms
for f in tree.xpath('//form'):
    print("Form action:", f.get('action'))
