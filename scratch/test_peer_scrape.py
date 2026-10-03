import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client
import lxml.html

session = screener_client.get_screener_session()
for ticker in ['ITC', 'NESTLEIND', 'BRITANNIA', 'VBL', 'MARICO', 'GODREJCP']:
    try:
        r = session.get(f"https://www.screener.in/company/{ticker}/consolidated/", timeout=3.0)
        doc = lxml.html.fromstring(r.content)
        debt = 0.0
        cash = 0.0
        inv = 0.0
        rows = doc.xpath('//section[@id="balance-sheet"]//table//tr')
        for row in rows:
            txt = ''.join(row.xpath('.//text()'))
            if 'Borrowings' in txt:
                tds = [td.text_content().strip().replace(',', '') for td in row.xpath('.//td')]
                vals = [float(v) for v in tds if v and v.replace('-', '').replace('.', '').isdigit()]
                if vals:
                    debt = vals[-1]
            if 'Investments' in txt:
                tds = [td.text_content().strip().replace(',', '') for td in row.xpath('.//td')]
                vals = [float(v) for v in tds if v and v.replace('-', '').replace('.', '').isdigit()]
                if vals:
                    inv = vals[-1]
        print(f"{ticker}: Debt={debt}, Investments={inv}")
    except Exception as e:
        print(f"{ticker} error: {e}")
