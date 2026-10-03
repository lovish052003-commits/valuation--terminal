import requests
import pandas as pd
from io import StringIO
from lxml import html

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
})

r_s = session.get('https://www.screener.in/company/NESTLEIND/')
t_s = html.fromstring(r_s.content)

print("=== STANDALONE SECTIONS ===")
for sec_id in ['profit-loss', 'balance-sheet', 'cash-flow', 'quarters']:
    t = t_s.xpath(f'//section[@id="{sec_id}"]//table')
    if t:
        df = pd.read_html(StringIO(html.tostring(t[0]).decode('utf-8')))[0]
        cols = [c for c in df.columns if c not in ('Metric', 'TTM') and not str(c).startswith('Unnamed')]
        print(f'{sec_id} cols count: {len(cols)}')
        print(f'{sec_id} cols:', cols)

r_c = session.get('https://www.screener.in/company/NESTLEIND/consolidated/')
t_c = html.fromstring(r_c.content)

print("\n=== CONSOLIDATED SECTIONS ===")
for sec_id in ['profit-loss', 'balance-sheet', 'cash-flow', 'quarters']:
    t = t_c.xpath(f'//section[@id="{sec_id}"]//table')
    if t:
        df = pd.read_html(StringIO(html.tostring(t[0]).decode('utf-8')))[0]
        cols = [c for c in df.columns if c not in ('Metric', 'TTM') and not str(c).startswith('Unnamed')]
        print(f'{sec_id} cols count: {len(cols)}')
        print(f'{sec_id} cols:', cols)
