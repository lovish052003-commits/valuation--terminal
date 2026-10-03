import requests
import lxml.html

def get_company_website_from_screener(ticker):
    r = requests.get(f'https://www.screener.in/company/{ticker}/', headers={'User-Agent': 'Mozilla/5.0'})
    doc = lxml.html.fromstring(r.content)
    ignore = ['screener.in', 'bseindia.com', 'nseindia.com', 'icra.in', 'careratings.com', 'crisil.com']
    for a in doc.xpath('//a/@href'):
        if a.startswith('http') and not any(x in a.lower() for x in ignore):
            return a
    return None

for t in ['TATASTEEL', 'ITC', 'JSWSTEEL', 'RELIANCE', 'HINDUNILVR']:
    print(t, '->', get_company_website_from_screener(t))
