import os, sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from screener_client import fetch_company_data

data = fetch_company_data('NESTLEIND')
tables = data.get('tables', {})

for sec, df in tables.items():
    print(f"\n=== Table '{sec}' ===")
    print("Metrics:", df['Metric'].tolist()[:15])
