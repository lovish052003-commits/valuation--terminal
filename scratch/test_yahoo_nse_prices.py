import requests
import datetime

tickers = ['RELIANCE', 'SUZLON', 'SUNPHARMA', 'ABCAPITAL', 'ITC', 'TATAMOTORS']
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

for t in tickers:
    for suffix in ['.NS', '.BO']:
        sym = f"{t}{suffix}"
        url = f"https://query1.finance.yahoo.com/v8/finance/chart/{sym}?range=1y&interval=1d"
        try:
            r = requests.get(url, headers=headers, timeout=5)
            if r.status_code == 200:
                data = r.json()
                res = data['chart']['result'][0]
                timestamps = res['timestamp']
                quote = res['indicators']['quote'][0]
                closes = quote.get('close', [])
                adjclose = res['indicators'].get('adjclose', [{}])[0].get('adjclose', closes)
                
                # Pair timestamps with closes
                pts = []
                for ts, c_val in zip(timestamps, closes):
                    if c_val is not None:
                        dt = datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d')
                        pts.append((dt, round(float(c_val), 2)))
                        
                print(f"SUCCESS {sym}: {len(pts)} points! Latest 3: {pts[-3:]}")
                break
            else:
                print(f"Failed {sym}: status {r.status_code}")
        except Exception as e:
            print(f"Error {sym}: {e}")
