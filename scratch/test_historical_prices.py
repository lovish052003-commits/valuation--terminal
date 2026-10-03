import requests
from datetime import datetime

def get_historical_prices(company_id, target_dates):
    url = f'https://www.screener.in/api/company/{company_id}/chart/?q=Price-DMA50-DMA200-Volume&days=3650'
    session = requests.Session()
    session.headers.update({'User-Agent': 'Mozilla/5.0'})
    r = session.get(url)
    if r.status_code != 200:
        return {}
    
    data = r.json()
    price_series = []
    for ds in data.get('datasets', []):
        if ds.get('metric') == 'Price':
            for pt in ds.get('values', []):
                # pt is [date_str, price_str]
                try:
                    dt = datetime.strptime(pt[0], '%Y-%m-%d')
                    pr = float(pt[1])
                    price_series.append((dt, pr))
                except Exception:
                    pass
            break
    
    if not price_series:
        return {}
    
    price_series.sort(key=lambda x: x[0])
    
    res = {}
    for td in target_dates:
        # find closest price on or before td, or nearest
        closest_p = None
        min_diff = float('inf')
        for dt, pr in price_series:
            diff = abs((dt - td).days)
            if diff < min_diff:
                min_diff = diff
                closest_p = pr
        res[td] = closest_p
    return res

# Test with Nestle target dates
dates = [
    datetime(2017, 12, 31),
    datetime(2018, 12, 31),
    datetime(2019, 12, 31),
    datetime(2020, 12, 31),
    datetime(2021, 12, 31),
    datetime(2022, 12, 31),
    datetime(2023, 12, 31),
    datetime(2024, 3, 31),
    datetime(2025, 3, 31),
    datetime(2026, 3, 31)
]

# We need company_id for NESTLEIND
r = requests.get('https://www.screener.in/company/NESTLEIND/', headers={'User-Agent': 'Mozilla/5.0'})
import re
m = re.search(r'data-company-id=[\"\'](\d+)[\"\']', r.text) or re.search(r'data-warehouse-id=[\"\'](\d+)[\"\']', r.text)
if m:
    cid = m.group(1)
    print("Found CID:", cid)
    prices = get_historical_prices(cid, dates)
    expected = [393.55, 554.24, 739.27, 919.51, 985.29, 980.3, 1329.02, 1311.18, 1125.38, 1174.8]
    for (d, p), exp in zip(prices.items(), expected):
        print(f"Date: {d.strftime('%Y-%m-%d')} | Fetched Price: {p:8.2f} | Expected: {exp:8.2f} | Diff: {abs(p-exp):6.2f}")
