import requests

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
})

url = "https://en.wikipedia.org/w/api.php?action=query&list=search&srsearch=Adani%20Enterprises&format=json"
r = session.get(url, timeout=5)
print("Status:", r.status_code)
print("Text:", r.text[:200])
