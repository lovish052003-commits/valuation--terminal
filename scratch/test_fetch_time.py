import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import time
import screener_client

t0 = time.time()
print("Starting fetch...")
data = screener_client.fetch_company_data('THE UNITED NILGIRI TEA ESTATES COMPANY LTD')
t1 = time.time()
print(f"Fetch completed in {t1 - t0:.2f}s")
print(f"Company: {data.get('company_name')}, Ticker: {data.get('ticker')}")
print(f"Current Price: {data.get('current_price')}")
print(f"Peers: {len(data.get('peers', []))}")
