import zipfile, re

for path in ['scratch/test_breaklink.xlsx', 'exports/TEST_MARUTI_COM.xlsx', 'ITC Model.xlsx']:
    with zipfile.ZipFile(path, 'r') as z:
        bad_count = 0
        for n in z.namelist():
            if n.endswith('.rels'):
                data = z.read(n).decode('utf-8', errors='ignore')
                targets = re.findall(r'Target="([^"]+)"', data)
                bad_targets = [t for t in targets if t.startswith('/xl/') or t.startswith('/')]
                bad_count += len(bad_targets)
        print(f"{path} has {bad_count} bad targets!")
