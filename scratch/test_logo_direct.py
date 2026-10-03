import requests, re
from PIL import Image
from io import BytesIO

def fetch_direct_from_website(website):
    if not website:
        return None
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    try:
        r = requests.get(website, headers=headers, timeout=5)
        if r.status_code == 200:
            for m in re.finditer(r'src=["\']([^"\']+)["\']', r.text, re.I):
                src = m.group(1)
                if 'logo' in src.lower() and not any(x in src.lower() for x in ['footer', 'icon', 'svg', 'flag', 'arrow', 'dummy']):
                    if src.startswith('//'):
                        img_url = 'https:' + src
                    elif src.startswith('/'):
                        base = re.match(r'https?://[^/]+', website).group(0)
                        img_url = base + src
                    elif src.startswith('http'):
                        img_url = src
                    else:
                        base = website.rstrip('/') + '/'
                        img_url = base + src
                    print('Attempting logo candidate:', img_url)
                    try:
                        r_img = requests.get(img_url, headers=headers, timeout=4)
                        if r_img.status_code == 200 and len(r_img.content) > 1000:
                            im = Image.open(BytesIO(r_img.content))
                            print('Successfully downloaded direct site logo!', im.size, im.mode)
                            return im
                    except Exception as e_im:
                        print('Failed to download image:', e_im)
    except Exception as e:
        print('Error:', e)
    return None

if __name__ == '__main__':
    im = fetch_direct_from_website('http://www.jubilantfoodworks.com')
    if im:
        im.save('scratch/jubl_logo.png')
        print('Saved to scratch/jubl_logo.png')
