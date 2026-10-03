import requests
from io import BytesIO
from PIL import Image

def test_fetch_domain_logo(domain):
    urls = [
        f'https://logo.clearbit.com/{domain}',
        f'https://t2.gstatic.com/faviconV2?client=SOCIAL&type=FAVICON&fallback_opts=TYPE,SIZE,URL&url=https://{domain}&size=256',
        f'https://www.google.com/s2/favicons?domain={domain}&sz=128',
        f'https://icon.horse/icon/{domain}'
    ]
    for u in urls:
        try:
            r = requests.get(u, headers={'User-Agent': 'Mozilla/5.0'}, timeout=5)
            if r.status_code == 200 and len(r.content) > 600:
                im = Image.open(BytesIO(r.content))
                print(f"Domain {domain} fetched from {u}: size={im.size}, mode={im.mode}")
                return im
        except Exception as e:
            pass
    print(f"Domain {domain}: No logo found")
    return None

for d in ['tatasteel.com', 'itcportal.com', 'jsw.in', 'ril.com', 'hul.co.in']:
    test_fetch_domain_logo(d)
