import sys, os
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__) + '/..'))
import requests
import lxml.html
from screener_client import clean_num

peers = ['JSW Steel', 'Tata Steel', 'Jindal Steel', 'S A I L', 'Jindal Stain.']
for p in peers:
    try:
        r = requests.get(f'https://www.screener.in/api/company/search/?q={p}', headers={'User-Agent': 'Mozilla/5.0'}).json()
        if r:
            url = f"https://www.screener.in{r[0]['url']}"
            page = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'})
            doc = lxml.html.fromstring(page.content)
            rows = doc.xpath('//section[@id="balance-sheet"]//table//tr')
            borrowing = 0
            for row in rows:
                txt = ''.join(row.xpath('.//text()'))
                if 'Borrowings' in txt:
                    vals = [clean_num(td.text_content()) for td in row.xpath('.//td') if clean_num(td.text_content()) != 0]
                    if vals:
                        borrowing = vals[-1]
                    break
            print(p, 'Borrowing:', borrowing)
    except Exception as e:
        print(p, 'Error:', e)
