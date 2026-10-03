import requests
from io import BytesIO
from PIL import Image
import lxml.html
import re

def get_real_company_logo(company_name, ticker, domain=None):
    session = requests.Session()
    session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
    
    # 1. Official domain from Screener page if not provided
    if not domain or 'screener' in domain:
        try:
            r = session.get(f'https://www.screener.in/company/{ticker}/', timeout=5)
            if r.status_code == 200:
                doc = lxml.html.fromstring(r.content)
                ignore = ['screener.in', 'bseindia.com', 'nseindia.com', 'icra.in', 'careratings.com', 'crisil.com']
                for a in doc.xpath('//a/@href'):
                    if a.startswith('http') and not any(x in a.lower() for x in ignore):
                        m = re.search(r'https?://(?:www\.)?([^/]+)', a)
                        if m:
                            domain = m.group(1).lower()
                            break
        except Exception:
            pass

    print(f"[{ticker}] Domain resolved to: {domain}")

    # 2. Wikipedia Infobox Logo (very high quality for major Indian companies)
    for q in [f"{company_name} (company)", company_name, ticker]:
        try:
            wiki_url = f"https://en.wikipedia.org/wiki/{q.replace(' ', '_')}"
            r = session.get(wiki_url, timeout=5)
            if r.status_code == 200:
                doc = lxml.html.fromstring(r.content)
                logo_imgs = doc.xpath('//td[contains(@class, "infobox-image")]//img/@src') or \
                            doc.xpath('//table[contains(@class, "infobox")]//img[contains(@alt, "logo") or contains(@alt, "Logo")]/@src')
                if logo_imgs:
                    img_url = "https:" + logo_imgs[0] if logo_imgs[0].startswith("//") else logo_imgs[0]
                    # Get higher res version by modifying wikipedia thumb URL if thumb
                    # e.g. /thumb/.../220px-... -> 400px
                    img_url = re.sub(r'/\d+px-', '/400px-', img_url)
                    resp = session.get(img_url, timeout=5)
                    if resp.status_code == 200:
                        im = Image.open(BytesIO(resp.content)).convert('RGBA')
                        print(f"[{ticker}] Sourced from Wikipedia: size={im.size}")
                        return im
        except Exception:
            pass

    # 3. Google Favicon V2 & Clearbit on official domain
    if domain and 'screener' not in domain:
        for u in [
            f'https://logo.clearbit.com/{domain}',
            f'https://t2.gstatic.com/faviconV2?client=SOCIAL&type=FAVICON&fallback_opts=TYPE,SIZE,URL&url=https://{domain}&size=256',
            f'https://icon.horse/icon/{domain}'
        ]:
            try:
                r = session.get(u, timeout=5)
                if r.status_code == 200 and len(r.content) > 600:
                    im = Image.open(BytesIO(r.content)).convert('RGBA')
                    print(f"[{ticker}] Sourced from {u}: size={im.size}")
                    return im
            except Exception:
                pass

    return None

for c, t in [('Tata Steel', 'TATASTEEL'), ('ITC Limited', 'ITC'), ('JSW Steel', 'JSWSTEEL'), ('Reliance Industries', 'RELIANCE')]:
    im = get_real_company_logo(c, t)
    if im:
        im.save(f'scratch/{t}_logo.png')
        print(f"Saved scratch/{t}_logo.png successfully")
