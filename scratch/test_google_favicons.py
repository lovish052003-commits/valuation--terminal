import requests
from PIL import Image
from io import BytesIO

endpoints = [
    'https://t2.gstatic.com/faviconV2?client=SOCIAL&type=FAVICON&fallback_opts=TYPE,SIZE,URL&url=https://{domain}&size=256',
    'https://www.google.com/s2/favicons?domain={domain}&sz=128',
    'https://icon.horse/icon/{domain}',
    'https://logo.clearbit.com/{domain}'
]

test_domains = {
    'TATASTEEL': ['tatasteel.com', 'tata.com'],
    'RELIANCE': ['ril.com', 'relianceindustries.com'],
    'ITC': ['itcportal.com', 'itc.com'],
    'JSWSTEEL': ['jsw.in', 'jswsteel.in'],
    'HINDUNILVR': ['hul.co.in', 'unilever.com'],
    'SBIN': ['sbi.co.in', 'statebankofindia.com']
}

for ticker, doms in test_domains.items():
    found = False
    for d in doms:
        if found:
            break
        for ep in endpoints:
            url = ep.format(domain=d)
            try:
                r = requests.get(url, timeout=4)
                if r.status_code == 200 and len(r.content) > 500:
                    im = Image.open(BytesIO(r.content))
                    if im.size[0] >= 32 and im.size[1] >= 32:
                        print(f"[{ticker}] SUCCESS from {url}: size={im.size}, mode={im.mode}")
                        found = True
                        break
            except Exception:
                pass
    if not found:
        print(f"[{ticker}] Not found")
