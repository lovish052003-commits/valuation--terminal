import requests, re
from lxml import html
import os, sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from screener_client import HEADERS

r = requests.get('https://www.screener.in/company/NESTLEIND/consolidated/', headers=HEADERS)
print('Status:', r.status_code)
tree = html.fromstring(r.content)

# Check all links with 'excel' or 'export'
links = tree.xpath('//a[contains(@href, "excel") or contains(@href, "export") or contains(text(), "Export") or contains(text(), "Excel")]')
for a in links:
    print('Found link:', a.text_content().strip(), 'href:', a.attrib.get('href'))

# Also search full html text
all_matches = re.findall(r'href=["\']([^"\']*(?:excel|export)[^"\']*)["\']', r.text, re.I)
print('Regex matches:', all_matches)

# Check buttons or forms
forms = tree.xpath('//form[contains(@action, "excel") or contains(@action, "export")]')
for f in forms:
    print('Found form:', f.attrib.get('action'))
