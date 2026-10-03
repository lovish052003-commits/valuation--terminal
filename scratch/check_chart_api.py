import requests
import json

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0'})
r = session.get('https://www.screener.in/api/company/1247/chart/?q=Price-DMA50-DMA200-Volume&days=3650')
print("Status code:", r.status_code)
if r.status_code == 200:
    data = r.json()
    print("Chart JSON keys:", data.keys())
    datasets = data.get('datasets', [])
    for ds in datasets:
        print("Metric:", ds.get('metric'), "Values count:", len(ds.get('values', [])))
        if ds.get('metric') == 'Price' and ds.get('values'):
            print("First 3 price points:", ds.get('values')[:3])
            print("Last 3 price points:", ds.get('values')[-3:])
