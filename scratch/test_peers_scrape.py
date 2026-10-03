import requests
import lxml.html

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
r = session.get('https://www.screener.in/company/ADANIENT/consolidated/')
tree = lxml.html.fromstring(r.text)
peers_sec = tree.xpath('//*[@id="peers"]')
if peers_sec:
    for a in peers_sec[0].xpath('.//a'):
        print(repr(a.text_content().strip()), '-->', repr(a.get('href')))
