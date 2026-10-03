import sys, os
sys.path.insert(0, os.path.abspath('.'))
from screener_client import fetch_company_data

data = fetch_company_data('NESTLEIND')
schedules = data.get('schedules', {})
print("Schedules fetched keys:", list(schedules.keys()))

for k, df in schedules.items():
    print(f"\n--- Schedule: {k} ---")
    if df is not None and not df.empty:
        print(df.head(4))
