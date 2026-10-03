import requests, re
from lxml import html

HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
r = requests.get('https://www.screener.in/company/ITC/consolidated/', headers=HEADERS)

tree = html.fromstring(r.content)

# Look at balance sheet section rows
bs_sec = tree.xpath('//section[@id="balance-sheet"]')
if bs_sec:
    rows = bs_sec[0].xpath('.//tr')
    for tr in rows:
        text = " ".join([t.strip() for t in tr.xpath('.//text()') if t.strip()])
        classes = tr.attrib.get('class', '')
        btn = tr.xpath('.//button')
        btn_info = [b.attrib for b in btn]
        print(f"TR (class='{classes}'): {text[:50]} | buttons: {btn_info}")

# Look at all buttons in the page with data-
print("\n--- Buttons with data attributes ---")
for b in tree.xpath('//button[@data-row-disp] | //button[@data-custom-params] | //button[contains(@onclick, "schedule")]'):
    print(b.attrib)
