import zipfile, re

path = r'exports\UNITEDTEA_Valuation_Model.xlsx'
with zipfile.ZipFile(path, 'r') as z:
    for n in z.namelist():
        if n.startswith('xl/worksheets/_rels/'):
            data = z.read(n).decode('utf-8', errors='ignore')
            m = re.findall(r'Target="[^"]*drawing[^"]*"', data)
            if m:
                sheet_xml = n.replace('_rels/', '').replace('.rels', '')
                print(sheet_xml, m)
