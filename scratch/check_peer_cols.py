import os, sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from screener_client import fetch_company_data

data = fetch_company_data('ITC')
pdf = data.get('peers_df')
if pdf is not None and not pdf.empty:
    print("Peers columns:", pdf.columns.tolist())
    print(pdf.head(3))
else:
    print("No peers df found")
