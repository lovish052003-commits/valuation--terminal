import requests, re
from lxml import html

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
})

r = session.get('https://www.screener.in/company/SUNPHARMA/consolidated/')
tree = html.fromstring(r.content)

# Look at company id and warehouse id
company_info = tree.xpath('//*[@id="company-info"]')
attrs = company_info[0].attrib if company_info else {}
print("company_info attrs:", attrs)

# Check all buttons in the profit-loss section
pl_sec = tree.xpath('//section[@id="profit-loss"]')
if pl_sec:
    buttons = pl_sec[0].xpath('.//button')
    print(f"Profit-loss buttons ({len(buttons)}):")
    for b in buttons:
        print("  button text:", b.text_content().strip(), "attrs:", b.attrib)

# Check all buttons in balance-sheet section
bs_sec = tree.xpath('//section[@id="balance-sheet"]')
if bs_sec:
    buttons = bs_sec[0].xpath('.//button')
    print(f"Balance-sheet buttons ({len(buttons)}):")
    for b in buttons:
        print("  button text:", b.text_content().strip(), "attrs:", b.attrib)
