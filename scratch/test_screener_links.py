import requests
import lxml.html

r = requests.get('https://www.screener.in/company/TATASTEEL/', headers={'User-Agent': 'Mozilla/5.0'})
doc = lxml.html.fromstring(r.content)

external = [a for a in doc.xpath('//a/@href') if a.startswith('http') and 'screener.in' not in a]
print('External links on Screener for TATASTEEL:')
for link in external[:10]:
    print('  ', link)

# Also test ITC
r2 = requests.get('https://www.screener.in/company/ITC/', headers={'User-Agent': 'Mozilla/5.0'})
doc2 = lxml.html.fromstring(r2.content)
external2 = [a for a in doc2.xpath('//a/@href') if a.startswith('http') and 'screener.in' not in a]
print('External links on Screener for ITC:')
for link in external2[:10]:
    print('  ', link)
