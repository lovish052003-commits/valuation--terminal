import os, sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from screener_client import fetch_company_data

data = fetch_company_data('NESTLEIND')
sch = data.get('schedules', {})
oa = sch.get('Other Assets', {})

print("Other Assets schedule keys:", list(oa.keys()))
for k in ['Inventories', 'Trade receivables', 'Cash Equivalents']:
    print(f"\n{k}:")
    if k in oa:
        # print last 10 periods
        for p, v in list(oa[k].items())[-10:]:
            print(f"  {p}: {v}")
