import zipfile, re, os, shutil

def sanitize_xlsx_relationships(xlsx_path):
    if not xlsx_path or not os.path.exists(xlsx_path):
        return
    temp_zip = xlsx_path + '.sanitizerels.tmp'
    try:
        fixed_count = 0
        with zipfile.ZipFile(xlsx_path, 'r') as zin:
            with zipfile.ZipFile(temp_zip, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename.endswith('.rels'):
                        text = data.decode('utf-8', errors='ignore')
                        orig = text
                        parts = item.filename.split('/')
                        # e.g. xl/worksheets/_rels/sheet1.xml.rels
                        if len(parts) >= 3 and parts[0] == 'xl' and parts[-2] == '_rels':
                            text = re.sub(r'Target="/xl/([^"]+)"', r'Target="../\1"', text)
                        elif len(parts) == 3 and parts[0] == 'xl' and parts[1] == '_rels':
                            text = re.sub(r'Target="/xl/([^"]+)"', r'Target="\1"', text)
                        elif len(parts) == 2 and parts[0] == '_rels':
                            text = re.sub(r'Target="/xl/([^"]+)"', r'Target="xl/\1"', text)
                        if text != orig:
                            fixed_count += 1
                            data = text.encode('utf-8')
                    zout.writestr(item, data)
        if fixed_count > 0:
            os.replace(temp_zip, xlsx_path)
            print(f"Sanitized {fixed_count} rels files in {xlsx_path}")
        else:
            if os.path.exists(temp_zip):
                os.remove(temp_zip)
    except Exception as e:
        if os.path.exists(temp_zip):
            try: os.remove(temp_zip)
            except: pass
        print(f"Sanitize error: {e}")

shutil.copyfile('exports/UNITEDTEA_Valuation_Model.xlsx', 'scratch/test_healed_corrupted.xlsx')
sanitize_xlsx_relationships('scratch/test_healed_corrupted.xlsx')

# Now scan bad targets in healed file
with zipfile.ZipFile('scratch/test_healed_corrupted.xlsx', 'r') as z:
    bad_count = 0
    for n in z.namelist():
        if n.endswith('.rels'):
            d = z.read(n).decode('utf-8', errors='ignore')
            targets = re.findall(r'Target="([^"]+)"', d)
            bad_targets = [t for t in targets if t.startswith('/xl/') or t.startswith('/')]
            bad_count += len(bad_targets)
    print(f"scratch/test_healed_corrupted.xlsx now has {bad_count} bad targets!")
