import requests, datetime, os, sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from screener_client import HEADERS, clean_num

def get_historical_prices(company_id, period_dates):
    """
    Fetches historical stock prices from Screener.in chart API
    and matches the closing price closest to each period end date.
    """
    if not company_id:
        return {}
    try:
        url = f"https://www.screener.in/api/company/{company_id}/chart/?q=Price&days=4000"
        r = requests.get(url, headers=HEADERS, timeout=8)
        if r.status_code != 200:
            return {}
        data = r.json()
        datasets = data.get('datasets', [])
        if not datasets:
            return {}
        raw_vals = datasets[0].get('values', [])
        price_by_date = {}
        for d_str, p_str in raw_vals:
            try:
                dt = datetime.datetime.strptime(d_str, '%Y-%m-%d').date()
                price_by_date[dt] = clean_num(p_str)
            except:
                pass
        
        sorted_dates = sorted(price_by_date.keys())
        if not sorted_dates:
            return {}
            
        matched_prices = {}
        for p_date_str in period_dates:
            target_dt = None
            for fmt in ['%Y-%m-%d', '%b %Y', '%b-%y', '%B %Y', '%Y-%m-%d %H:%M:%S']:
                try:
                    target_dt = datetime.datetime.strptime(p_date_str.strip()[:10], fmt).date()
                    break
                except:
                    pass
            if not target_dt:
                continue
            
            candidates = [d for d in sorted_dates if d <= target_dt]
            if candidates:
                closest_d = candidates[-1]
            else:
                closest_d = sorted_dates[0]
            matched_prices[p_date_str] = price_by_date[closest_d]
        return matched_prices
    except Exception as e:
        print("Error fetching chart prices:", e)
        return {}

# Test on NESTLEIND (company_id 2026)
test_dates = ['2017-12-31', '2018-12-31', '2019-12-31', '2020-12-31', '2021-12-31', '2022-12-31', '2023-12-31', '2024-03-31', '2025-03-31', '2026-03-31']
res = get_historical_prices(2026, test_dates)
print("Matched prices for Nestle:")
for d, p in res.items():
    print(f"  {d}: {p}")
