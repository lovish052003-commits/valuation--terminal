import zipfile

good_path = 'exports/TEST_MARUTI_COM.xlsx'
bad_path = 'exports/UNITEDTEA_Valuation_Model.xlsx'
itc_path = 'ITC Model.xlsx'

print("=== CHECKING EXTERNAL LINKS ===")
for p, label in [(good_path, 'GOOD (MARUTI)'), (bad_path, 'BAD (UNITEDTEA)'), (itc_path, 'TEMPLATE (ITC)')]:
    with zipfile.ZipFile(p, 'r') as z:
        exts = [n for n in z.namelist() if 'external' in n.lower()]
        dwgs = [n for n in z.namelist() if 'drawing' in n.lower()]
        print(f"{label}: externalLinks={exts}, drawings count={len(dwgs)}")

print("\n=== CHECKING DRAWING RELS IN GOOD VS BAD ===")
with zipfile.ZipFile(good_path, 'r') as z_good, zipfile.ZipFile(bad_path, 'r') as z_bad:
    print("GOOD drawing1 rels:")
    if 'xl/drawings/_rels/drawing1.xml.rels' in z_good.namelist():
        print("  ", z_good.read('xl/drawings/_rels/drawing1.xml.rels').decode('utf-8'))
    print("BAD drawing1 rels:")
    if 'xl/drawings/_rels/drawing1.xml.rels' in z_bad.namelist():
        print("  ", z_bad.read('xl/drawings/_rels/drawing1.xml.rels').decode('utf-8'))

    print("GOOD workbook.xml.rels external references:")
    rels_good = z_good.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    for line in rels_good.split('>'):
        if 'external' in line:
            print("  GOOD:", line)
            
    print("BAD workbook.xml.rels external references:")
    rels_bad = z_bad.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    for line in rels_bad.split('>'):
        if 'external' in line:
            print("  BAD:", line)
