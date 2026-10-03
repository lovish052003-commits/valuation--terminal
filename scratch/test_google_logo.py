import requests
import re
from urllib.parse import quote_plus
from PIL import Image
from io import BytesIO

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.5'
}

def search_google_logo(company_name):
    """Searches Google Images for the company logo."""
    query = f"{company_name} logo png"
    url = f"https://www.google.com/search?q={quote_plus(query)}&tbm=isch&asearch=ichunk"
    try:
        r = requests.get(url, headers=headers, timeout=6)
        if r.status_code == 200:
            # Look for image URLs in the response
            # Google images embed direct image URLs in data attributes or json
            matches = re.findall(r'https://[^"\'>\s]+?\.(?:png|jpg|jpeg)', r.text)
            clean_matches = [m for m in matches if 'google' not in m and 'gstatic' not in m]
            print(f"Found {len(clean_matches)} external image URLs for {company_name}")
            if clean_matches:
                print("First 3:", clean_matches[:3])
                for img_url in clean_matches[:5]:
                    try:
                        ir = requests.get(img_url, headers=headers, timeout=5)
                        if ir.status_code == 200 and len(ir.content) > 1000:
                            im = Image.open(BytesIO(ir.content))
                            print(f"Successfully downloaded from Google search: {img_url} ({im.size})")
                            return im
                    except Exception:
                        continue
    except Exception as e:
        print(f"Error searching Google: {e}")
    return None

for c in ['Tata Steel', 'Reliance Industries', 'JSW Steel']:
    im = search_google_logo(c)
    print(c, '->', im)
