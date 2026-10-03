import requests
import re
from PIL import Image
from io import BytesIO

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

def search_company_logo_online(company_name, ticker=''):
    s = requests.Session()
    s.headers.update(headers)
    
    # Strategy 1: Google Favicon API (V1 & V2) with domains
    domains = []
    if ticker:
        t = ticker.lower().replace(' ', '')
        domains.extend([f"{t}.com", f"{t}.in", f"{t}.co.in"])
    name_clean = re.sub(r'[^\w\s]', '', company_name.lower())
    words = [w for w in name_clean.split() if w not in ['ltd', 'limited', 'inds', 'industries', 'co', 'corp', 'india']]
    if words:
        domains.extend([
            f"{words[0]}.com", f"{words[0]}.in", f"{words[0]}.co.in",
            f"{''.join(words[:2])}.com", f"{''.join(words[:2])}.in"
        ])
    
    # Conglomerate domain mapping
    domain_map = {
        'TATASTEEL': 'tatasteel.com',
        'TATAMOTORS': 'tatamotors.com',
        'TCS': 'tcs.com',
        'RELIANCE': 'ril.com',
        'HINDUNILVR': 'hul.co.in',
        'ITC': 'itcportal.com',
        'JSWSTEEL': 'jsw.in',
        'INFY': 'infosys.com',
        'SBIN': 'sbi.co.in',
        'HDFCBANK': 'hdfcbank.com',
        'ICICIBANK': 'icicibank.com',
        'BHARTIARTL': 'airtel.in',
        'LT': 'larsentoubro.com',
        'NESTLEIND': 'nestle.in',
        'VBL': 'varunpepsi.com',
        'BAJFINANCE': 'bajajfinserv.in',
        'VEDL': 'vedantalimited.com',
        'COALINDIA': 'coalindia.in',
        'IOC': 'iocl.com',
        'NTPC': 'ntpc.co.in',
        'ONGC': 'ongcindia.com',
        'POWERGRID': 'powergrid.in',
        'SUNPHARMA': 'sunpharma.com',
        'CIPLA': 'cipla.com',
        'DRREDDY': 'drreddys.com',
        'ADANIENT': 'adanienterprises.com',
        'ADANIPORTS': 'adaniports.com',
        'ASIANPAINT': 'asianpaints.com',
        'TITAN': 'titancompany.in',
        'MARUTI': 'marutisuzuki.com',
        'WIPRO': 'wipro.com',
        'HCLTECH': 'hcltech.com'
    }
    if ticker.upper() in domain_map:
        domains.insert(0, domain_map[ticker.upper()])

    for dom in domains:
        # Try Google Favicon V2 (256px resolution)
        try:
            g_url = f"https://t2.gstatic.com/faviconV2?client=SOCIAL&type=FAVICON&fallback_opts=TYPE,SIZE,URL&url=https://www.{dom}&size=256"
            r = s.get(g_url, timeout=4)
            if r.status_code == 200 and len(r.content) > 1500:
                im = Image.open(BytesIO(r.content))
                if im.size[0] >= 32:
                    print(f"[{company_name}] Found Google favicon logo via {dom} (size: {im.size})")
                    return im
        except Exception:
            pass

        # Try Clearbit
        try:
            cb_url = f"https://logo.clearbit.com/{dom}"
            r = s.get(cb_url, timeout=4)
            if r.status_code == 200 and len(r.content) > 1000:
                im = Image.open(BytesIO(r.content))
                print(f"[{company_name}] Found Clearbit logo via {dom} (size: {im.size})")
                return im
        except Exception:
            pass

    # Strategy 2: DuckDuckGo search
    try:
        q = f"{company_name} logo png"
        r1 = s.get(f"https://duckduckgo.com/?q={requests.utils.quote(q)}", timeout=5)
        m = re.search(r'vqd=([\'"]?)([\d-]+)\1', r1.text)
        if m:
            vqd = m.group(2)
            i_url = f"https://duckduckgo.com/i.js?q={requests.utils.quote(q)}&vqd={vqd}"
            r2 = s.get(i_url, timeout=5)
            if r2.status_code == 200:
                res = r2.json().get('results', [])
                for item in res[:5]:
                    img_url = item.get('image')
                    if img_url and any(ext in img_url.lower() for ext in ['.png', '.jpg', '.webp']):
                        ir = s.get(img_url, timeout=5)
                        if ir.status_code == 200 and len(ir.content) > 2000:
                            im = Image.open(BytesIO(ir.content))
                            print(f"[{company_name}] Found DDG image logo: {img_url} (size: {im.size})")
                            return im
    except Exception as e:
        print(f"DDG search error: {e}")

    return None

test_companies = [
    ('Tata Steel Ltd', 'TATASTEEL'),
    ('Reliance Industries Ltd', 'RELIANCE'),
    ('ITC Ltd', 'ITC'),
    ('JSW Steel Ltd', 'JSWSTEEL'),
    ('Hindustan Unilever Ltd', 'HINDUNILVR'),
    ('State Bank of India', 'SBIN')
]

for name, tick in test_companies:
    img = search_company_logo_online(name, tick)
    print(f"Result for {name}: {img}\n")
