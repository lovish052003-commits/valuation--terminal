"""
company_logo_manager.py
Manages dynamic company logo sourcing from Google / official web sources,
formats the logos with alpha transparency to match the financial model theme,
and embeds them into the exported Excel model (Dupont Analysis, Altman's Z Score,
DCF, Comp_Valuation) and terminal web application.
"""

import os
import re
import shutil
import zipfile
import requests
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont

STATIC_LOGO_DIR = os.path.join(os.path.dirname(__file__), 'static', 'img', 'logos')
os.makedirs(STATIC_LOGO_DIR, exist_ok=True)

# High-fidelity domain directory for Indian listed companies
DOMAIN_MAP = {
    'TATASTEEL': 'tatasteel.com',
    'TATAMOTORS': 'tatamotors.com',
    'TCS': 'tcs.com',
    'RELIANCE': 'relianceindustries.com',
    'HINDUNILVR': 'hul.co.in',
    'ITC': 'itcportal.com',
    'JSWSTEEL': 'jsw.in',
    'INFY': 'infosys.com',
    'SBIN': 'statebankofindia.com',
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
    'HCLTECH': 'hcltech.com',
    'JUBLFOOD': 'jubilantfoodworks.com',
    'BRITANNIA': 'britannia.co.in',
    'DABUR': 'dabur.com',
    'MARICO': 'marico.com',
    'GODREJCP': 'godrejcp.com',
    'ZOMATO': 'zomato.com',
    'SWIGGY': 'swiggy.com',
    'COLPAL': 'colgatepalmolive.co.in',
    'PIDILITIND': 'pidilite.com',
    'BERGEPAINT': 'bergerpaints.com',
    'HAVELLS': 'havells.com',
    'VOLTAS': 'voltas.com',
    'PAGEIND': 'pageindustries.com',
    'DMART': 'dmartindia.com'
}

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}


def generate_fallback_badge(company_name, ticker=""):
    """
    Generates a sleek, institutional monochrome vector-styled logo badge
    matching the corporate theme if no online logo is accessible.
    """
    canvas = Image.new('RGBA', (330, 344), (0, 0, 0, 0))
    draw = ImageDraw.Draw(canvas)
    
    # Draw rounded emblem
    # Theme color: Corporate dark blue / slate (#062F62)
    emblem_color = (6, 47, 98, 255)
    draw.rounded_rectangle([(30, 37), (300, 307)], radius=45, fill=emblem_color)
    
    # Initials
    initials = "".join([w[0] for w in company_name.split() if w.lower() not in ['ltd', 'limited', 'inds', 'co', 'the']][:3]).upper()
    if not initials and ticker:
        initials = ticker[:3].upper()
    if not initials:
        initials = "VAL"

    try:
        font = ImageFont.truetype("arialbd.ttf", 95)
    except Exception:
        font = ImageFont.load_default()

    # Draw centered initials
    bbox = draw.textbbox((0, 0), initials, font=font)
    w = bbox[2] - bbox[0]
    h = bbox[3] - bbox[1]
    draw.text(((330 - w) / 2, (344 - h) / 2 - 15), initials, fill=(255, 255, 255, 255), font=font)
    
    return canvas


def fetch_logo_from_website_html(website_url, session):
    """
    Directly scrapes official website HTML to locate the primary brand logo.
    """
    if not website_url:
        return None
    try:
        r = session.get(website_url, timeout=4)
        if r.status_code == 200:
            for m in re.finditer(r'src=["\']([^"\']+)["\']', r.text, re.I):
                src = m.group(1)
                if 'logo' in src.lower() and not any(x in src.lower() for x in ['footer', 'icon', 'svg', 'flag', 'arrow', 'dummy', 'banner']):
                    if src.startswith('//'):
                        img_url = 'https:' + src
                    elif src.startswith('/'):
                        base = re.match(r'https?://[^/]+', website_url).group(0)
                        img_url = base + src
                    elif src.startswith('http'):
                        img_url = src
                    else:
                        base = website_url.rstrip('/') + '/'
                        img_url = base + src
                    try:
                        r_img = session.get(img_url, timeout=4)
                        if r_img.status_code == 200 and len(r_img.content) > 1000:
                            im = Image.open(BytesIO(r_img.content)).convert('RGBA')
                            if im.size[0] >= 30 and im.size[1] >= 30:
                                return im
                    except Exception:
                        pass
    except Exception:
        pass
    return None


def fetch_company_logo(company_name, ticker="", website=None):
    """
    Sources the official company logo from Google / web APIs:
    1. Checks local cache
    2. Directly extracts from official corporate website
    3. Google Favicon V2 & S2 APIs with inferred candidate domains
    4. Icon Horse & Clearbit fallback
    5. Returns a PIL Image object (RGBA)
    """
    ticker_clean = (ticker or 'COMPANY').upper().strip()
    cache_path = os.path.join(STATIC_LOGO_DIR, f"{ticker_clean}.png")
    
    # Check cache
    if os.path.exists(cache_path):
        try:
            im = Image.open(cache_path)
            return im.convert('RGBA')
        except Exception:
            pass

    session = requests.Session()
    session.headers.update(HEADERS)

    # 1. Direct Official Corporate Website Scraping
    if website:
        dir_im = fetch_logo_from_website_html(website, session)
        if dir_im:
            dir_im.save(cache_path, format='PNG')
            print(f"[Logo Manager] Successfully sourced official website logo for '{company_name}' ({website})")
            return dir_im

    if ticker_clean in DOMAIN_MAP:
        dom_url = f"https://www.{DOMAIN_MAP[ticker_clean]}"
        dir_im = fetch_logo_from_website_html(dom_url, session)
        if dir_im:
            dir_im.save(cache_path, format='PNG')
            print(f"[Logo Manager] Successfully sourced direct website logo via DOMAIN_MAP for '{company_name}' ({dom_url})")
            return dir_im

    # 2. Try DuckDuckGo Instant Answer API for official high-resolution corporate logo
    for q in [f"{company_name} Limited", company_name, f"{ticker_clean} Limited", ticker_clean]:
        try:
            ddg_url = f"https://api.duckduckgo.com/?q={requests.utils.quote(q)}&format=json"
            r = session.get(ddg_url, timeout=4)
            if r.status_code == 200:
                data = r.json()
                img_path = data.get('Image')
                if img_path and not any(x in img_path.lower() for x in ['screener', 'wikimedia', 'flag']):
                    img_url = f"https://duckduckgo.com{img_path}" if img_path.startswith('/') else img_path
                    r_img = session.get(img_url, timeout=4)
                    if r_img.status_code == 200 and len(r_img.content) > 1000:
                        im = Image.open(BytesIO(r_img.content)).convert('RGBA')
                        if im.size[0] >= 30 and im.size[1] >= 30:
                            im.save(cache_path, format='PNG')
                            print(f"[Logo Manager] Successfully sourced official logo for '{company_name}' via DuckDuckGo ({img_url})")
                            return im
        except Exception:
            pass

    # 2. Build candidate official domains (EXCLUDING screener.in and exchanges)
    candidate_domains = []
    if website:
        m = re.search(r'https?://(?:www\.)?([^/]+)', website)
        if m:
            dom = m.group(1).lower()
            if not any(x in dom for x in ['screener', 'bseindia', 'nseindia', 'icra', 'careratings']):
                candidate_domains.append(dom)
            
    if ticker_clean in DOMAIN_MAP:
        candidate_domains.append(DOMAIN_MAP[ticker_clean])

    name_clean = re.sub(r'[^\w\s]', '', company_name.lower())
    words = [w for w in name_clean.split() if w not in ['ltd', 'limited', 'inds', 'industries', 'co', 'corp', 'india', 'the']]
    if ticker:
        t = ticker.lower().replace(' ', '')
        candidate_domains.extend([f"{t}.com", f"{t}.in", f"{t}.co.in"])
    if words:
        brand = "".join(words[:2])
        candidate_domains.extend([
            f"{words[0]}.com", f"{words[0]}.in", f"{words[0]}.co.in",
            f"{brand}.com", f"{brand}.in", f"{brand}.co.in"
        ])

    endpoints = [
        'https://logo.clearbit.com/{domain}',
        'https://t2.gstatic.com/faviconV2?client=SOCIAL&type=FAVICON&fallback_opts=TYPE,SIZE,URL&url=https://{domain}&size=256',
        'https://icon.horse/icon/{domain}'
    ]

    for dom in candidate_domains:
        if any(x in dom for x in ['screener', 'bseindia', 'nseindia']):
            continue
        for ep in endpoints:
            url = ep.format(domain=dom)
            try:
                r = session.get(url, timeout=3.5)
                if r.status_code == 200 and len(r.content) > 600:
                    im = Image.open(BytesIO(r.content)).convert('RGBA')
                    if im.size[0] >= 32 and im.size[1] >= 32:
                        im.save(cache_path, format='PNG')
                        print(f"[Logo Manager] Successfully sourced logo for '{company_name}' ({ticker_clean}) via {dom}")
                        return im
            except Exception:
                continue

    # Fallback to institutional emblem badge
    fallback_im = generate_fallback_badge(company_name, ticker_clean)
    fallback_im.save(cache_path, format='PNG')
    print(f"[Logo Manager] Generated corporate emblem badge for '{company_name}' ({ticker_clean})")
    return fallback_im


def get_target_logo_images(dest_xlsx_path):
    """
    Dynamically identifies all image media paths bound to 'Dupont Analysis'
    and 'Altman's Z Score' across drawing relationships in the Excel workbook.
    """
    import xml.etree.ElementTree as ET
    target_images = set()
    if not os.path.exists(dest_xlsx_path):
        return target_images

    try:
        with zipfile.ZipFile(dest_xlsx_path, 'r') as z:
            if 'xl/workbook.xml' in z.namelist() and 'xl/_rels/workbook.xml.rels' in z.namelist():
                wb_xml = z.read('xl/workbook.xml')
                wb_tree = ET.fromstring(wb_xml)
                wb_rels_xml = z.read('xl/_rels/workbook.xml.rels')
                wb_rels_tree = ET.fromstring(wb_rels_xml)
                
                rel_map = {r.attrib['Id']: r.attrib['Target'] for r in wb_rels_tree}
                
                sheet_targets = {}
                sheets_node = wb_tree.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets')
                if sheets_node is not None:
                    for s in sheets_node:
                        name = s.attrib.get('name')
                        rId = s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                        if name and rId and rId in rel_map:
                            t = rel_map[rId]
                            if t.startswith('/xl/'):
                                t = t[1:]
                            elif not t.startswith('xl/'):
                                t = 'xl/' + t
                            sheet_targets[name] = t

                for sname in ['Dupont Analysis', "Altman's Z Score"]:
                    target = sheet_targets.get(sname)
                    if not target or '/' not in target:
                        continue
                    sheet_dir, sheet_file = target.rsplit('/', 1)
                    rels_path = f"{sheet_dir}/_rels/{sheet_file}.rels"
                    if rels_path in z.namelist():
                        s_rels_tree = ET.fromstring(z.read(rels_path))
                        for r in s_rels_tree:
                            if 'drawing' in r.attrib.get('Type', ''):
                                dwg_target = r.attrib.get('Target', '')
                                if dwg_target.startswith('../'):
                                    dwg_path = 'xl/' + dwg_target.replace('../', '')
                                elif dwg_target.startswith('/xl/'):
                                    dwg_path = dwg_target[1:]
                                else:
                                    dwg_path = f"{sheet_dir}/{dwg_target}"
                                
                                if '/' in dwg_path:
                                    dwg_dir, dwg_file = dwg_path.rsplit('/', 1)
                                    dwg_rels_path = f"{dwg_dir}/_rels/{dwg_file}.rels"
                                    if dwg_rels_path in z.namelist():
                                        d_rels_tree = ET.fromstring(z.read(dwg_rels_path))
                                        for dr in d_rels_tree:
                                            if 'image' in dr.attrib.get('Type', ''):
                                                img_target = dr.attrib.get('Target', '')
                                                if img_target.startswith('../'):
                                                    img_path = 'xl/' + img_target.replace('../', '')
                                                elif img_target.startswith('/xl/'):
                                                    img_path = img_target[1:]
                                                else:
                                                    img_path = f"{dwg_dir}/{img_target}"
                                                target_images.add(img_path)
    except Exception as e_parse:
        print(f"[Logo Manager] Notice: drawing parsing notice: {e_parse}")

    # Standard fallback image targets for template / COM exports
    target_images.add('xl/media/image3.png')
    target_images.add('xl/media/image7.png')
    target_images.add('xl/media/image8.png')
    return target_images


def embed_logo_in_excel(dest_xlsx_path, company_name, ticker="", website=None):
    """
    Embeds the company logo DIRECTLY into the exported Excel model package.
    STRICT USER CONSTRAINT:
    - Modifies ONLY the logo images rendered in 'Dupont Analysis' and 'Altman's Z Score'
      (e.g., xl/media/image7.png, xl/media/image8.png, xl/media/image3.png)
    - NEVER touches xl/media/image2.png or any other header ribbons / sheets!
    Maintains exact original dimensions (330 x 344), aspect ratios, and transparent backgrounds.
    """
    if not os.path.exists(dest_xlsx_path):
        return False

    try:
        raw_logo = fetch_company_logo(company_name, ticker, website)
        if not raw_logo:
            return False

        # Format image (Canvas 330 x 344) strictly for Dupont Analysis and Altman's Z Score
        canvas3 = Image.new('RGBA', (330, 344), (0, 0, 0, 0))
        im3 = raw_logo.copy()
        if im3.mode != 'RGBA':
            im3 = im3.convert('RGBA')
        im3.thumbnail((310, 325), Image.Resampling.LANCZOS)
        x3 = (330 - im3.width) // 2
        y3 = (344 - im3.height) // 2
        canvas3.paste(im3, (x3, y3), im3)
        buf3 = BytesIO()
        canvas3.save(buf3, format='PNG')
        img3_bytes = buf3.getvalue()

        target_imgs = get_target_logo_images(dest_xlsx_path)

        import time
        temp_path = dest_xlsx_path + ".logotmp"
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

        updated_count = 0
        with zipfile.ZipFile(dest_xlsx_path, 'r') as zin:
            with zipfile.ZipFile(temp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    is_target = False
                    if item.filename in target_imgs:
                        is_target = True
                    elif item.filename.startswith('xl/media/image') and item.file_size == 18741:
                        # Exactly matches template ITC logo byte size
                        is_target = True

                    if is_target:
                        # Write with clean filename string so zipfile recalculates proper CRC and size
                        zout.writestr(item.filename, img3_bytes)
                        updated_count += 1
                    else:
                        zout.writestr(item, zin.read(item.filename))

        if updated_count > 0:
            moved = False
            for _ in range(5):
                try:
                    shutil.move(temp_path, dest_xlsx_path)
                    moved = True
                    break
                except Exception:
                    time.sleep(0.3)
            if moved:
                print(f"[Logo Manager] Successfully updated {updated_count} logo(s) strictly in Dupont Analysis & Altman's Z Score for '{company_name}'!")
                return True
            else:
                if os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                    except Exception:
                        pass
                return False
        else:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
            return False

    except Exception as e:
        print(f"[Logo Manager] Warning: Could not embed logo in Excel: {e}")
        return False
