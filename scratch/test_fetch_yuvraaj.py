import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import screener_client

print('Fetching YUVRAAJ HYGIENE PRODUCTS LTD....')
data = screener_client.fetch_company_data('YUVRAAJ HYGIENE PRODUCTS LTD.')
print('SUCCESS!')
print('Company Name:', data.get('company_name'))
print('Current Price:', data.get('current_price'))
print('Market Cap:', data.get('market_cap_cr'))
print('Tables:', list(data.get('tables', {}).keys()))
