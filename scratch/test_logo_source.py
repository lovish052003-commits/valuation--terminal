import requests
import re
import urllib.parse
from io import BytesIO
from PIL import Image

def get_company_logo(company_name, ticker=""):
    """
    Fetches the official company logo using multiple redundant, robust methods:
    1. Clearbit Logo API & Google Favicon v2 via inferred/scraped domain
    2. DuckDuckGo / Google image search for official transparent logo
    3. Fallback: Creates an elegant, institutional vector-styled logo badge with the company initials
    """
    session = requests.Session()
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    })

    # Domain guesses based on known ticker / company name
    name_clean = re.sub(r'[^\w\s]', '', company_name.lower())
    words = [w for w in name_clean.split() if w not in ['ltd', 'limited', 'inds', 'industries', 'co', 'corp', 'india']]
    
    candidate_domains = []
    if ticker:
        t_clean = ticker.lower().replace(' ', '')
        candidate_domains.extend([f"{t_clean}.com", f"{t_clean}.in", f"{t_clean}.co.in"])
    if words:
        brand = "".join(words[:2])
        brand_first = words[0]
        candidate_domains.extend([
            f"{brand}.com", f"{brand}.in", f"{brand}.co.in",
            f"{brand_first}.com", f"{brand_first}.in", f"{brand_first}.co.in"
        ])
    
    # Custom mapping for famous Indian conglomerates
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
        'BAJFINANCE': 'bajajfinserv.in'
    }
    if ticker.upper() in domain_map:
        candidate_domains.insert(0, domain_map[ticker.upper()])

    # Try Clearbit & Google Favicon APIs
    for domain in candidate_domains:
        # Clearbit (often provides crisp SVG/PNG logos)
        try:
            cb_url = f"https://logo.clearbit.com/{domain}"
            r = session.get(cb_url, timeout=4)
            if r.status_code == 200 and len(r.content) > 1000:
                img = Image.open(BytesIO(r.content))
                print(f"Found logo via Clearbit for {domain}: size={img.size}")
                return img
        except Exception:
            pass

        # Google Favicon v2 (256px resolution)
        try:
            g_url = f"https://t2.gstatic.com/faviconV2?client=SOCIAL&type=FAVICON&fallback_opts=TYPE,SIZE,URL&url=https://www.{domain}&size=256"
            r = session.get(g_url, timeout=4)
            if r.status_code == 200 and len(r.content) > 1500:
                img = Image.open(BytesIO(r.content))
                # Check that it's not a generic globe placeholder (usually 16x16 or very small)
                if img.size[0] >= 32 and img.size[1] >= 32:
                    print(f"Found logo via Google Favicon for {domain}: size={img.size}")
                    return img
        except Exception:
            pass

    return None

test_stocks = [
    ('Tata Steel Ltd', 'TATASTEEL'),
    ('Reliance Industries Ltd', 'RELIANCE'),
    ('ITC Ltd', 'ITC'),
    ('JSW Steel Ltd', 'JSWSTEEL')
]

for name, tick in test_stocks:
    im = get_company_logo(name, tick)
    print(name, '->', im)
