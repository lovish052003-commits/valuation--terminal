import requests, json

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
})

url = "https://www.screener.in/api/company/3245/schedules/?parent=Material%20Cost%20%25&section=profit-loss&consolidated="
r = session.get(url, timeout=10)
print("Status:", r.status_code)
if r.status_code == 200:
    print("Content:")
    print(json.dumps(r.json(), indent=2))
