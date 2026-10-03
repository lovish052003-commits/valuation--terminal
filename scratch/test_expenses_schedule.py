import requests, json

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
})

# Company ID 3245 is Sun Pharma
for parent in ['Expenses', 'Sales', 'Other Income', 'Net Profit']:
    for cons in ['', '&consolidated=']:
        url = f"https://www.screener.in/api/company/3245/schedules/?parent={requests.utils.quote(parent)}&section=profit-loss{cons}"
        try:
            r = session.get(url, timeout=10)
            print(f"URL: {url} -> Status: {r.status_code}")
            if r.status_code == 200:
                print("Response keys / content:")
                try:
                    js = r.json()
                    print(json.dumps(list(js.keys()) if isinstance(js, dict) else js, indent=2))
                    if isinstance(js, dict):
                        for k, v in list(js.items())[:3]:
                            print(f"  Item: {k} -> {v}")
                except Exception as ex:
                    print("Text:", r.text[:200])
        except Exception as e:
            print("Error:", e)
