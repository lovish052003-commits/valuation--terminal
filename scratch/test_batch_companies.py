import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import screener_client

companies = ['YUVRAAJ HYGIENE PRODUCTS LTD.', 'TATA MOTORS', 'ITC LTD', 'INFOSYS', 'SALASAR TECHNO', 'KOPRAN']
for c in companies:
    try:
        d = screener_client.fetch_company_data(c)
        print(f"SUCCESS: {c} -> {d.get('company_name')} ({d.get('ticker')}) CMP: Rs {d.get('current_price')}")
    except Exception as e:
        print(f"FAILED: {c} -> {e}")
