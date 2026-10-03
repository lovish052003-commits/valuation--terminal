import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client, lxml.html

session = screener_client.get_screener_session()
r = session.get('https://www.screener.in/company/ADANIENT/consolidated/')
doc = lxml.html.fromstring(r.content)
table = doc.xpath('//section[@id="peers"]//table')
if table:
    rows = table[0].xpath('.//tr')
    for r in rows:
        tds = [td.text_content().strip() for td in r.xpath('.//td | .//th')]
        print(tds[:6])
