import requests, datetime

def fetch_nse_daily_prices_1y(ticker):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    clean_t = ticker.upper().replace('.NS', '').replace('.BO', '').strip()
    
    # Common mappings
    ticker_candidates = [clean_t]
    if clean_t == 'TATAMOTORS':
        ticker_candidates.extend(['TATAMTRDVR', 'TTM'])
    
    for c_ticker in ticker_candidates:
        for suffix in ['.NS', '.BO']:
            sym = f"{c_ticker}{suffix}"
            url = f"https://query1.finance.yahoo.com/v8/finance/chart/{sym}?range=1y&interval=1d"
            try:
                r = requests.get(url, headers=headers, timeout=6)
                if r.status_code == 200:
                    data = r.json()
                    res = data['chart']['result'][0]
                    timestamps = res.get('timestamp', [])
                    quote = res['indicators']['quote'][0]
                    closes = quote.get('close', [])
                    
                    pts = []
                    for ts, c_val in zip(timestamps, closes):
                        if c_val is not None:
                            try:
                                dt = datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d')
                                pts.append((dt, round(float(c_val), 2)))
                            except Exception:
                                pass
                    if len(pts) >= 50:
                        pts.sort(key=lambda x: x[0], reverse=True)
                        return pts
            except Exception:
                pass
    return []

for t in ['RELIANCE', 'SUZLON', 'SUNPHARMA', 'ABCAPITAL', 'ITC']:
    res = fetch_nse_daily_prices_1y(t)
    print(f"{t}: {len(res)} daily points! Newest 3: {res[:3]}, Oldest 1: {res[-1]}")
