import sys, os
sys.path.insert(0, os.path.abspath('.'))
from screener_client import fetch_company_data

data = fetch_company_data('NESTLEIND')
schedules = data.get('schedules', {})

for k, val in schedules.items():
    print(f"\n================ Schedule: {k} (type: {type(val)}) ================")
    if isinstance(val, dict):
        for sub_k, sub_v in val.items():
            print(f"  {sub_k}: {sub_v}")
    elif hasattr(val, 'head'):
        print(val.head(4))
