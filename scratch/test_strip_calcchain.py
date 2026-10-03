import zipfile, os, re, xml.etree.ElementTree as ET

def remove_calc_chain(xlsx_path):
    temp_zip = xlsx_path + ".notchain.tmp"
    with zipfile.ZipFile(xlsx_path, 'r') as zin:
        with zipfile.ZipFile(temp_zip, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'xl/calcChain.xml':
                    print("Removing xl/calcChain.xml")
                    continue
                data = zin.read(item.filename)
                if item.filename == '[Content_Types].xml':
                    # Remove calcChain Override tag
                    data = re.sub(rb'<Override[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', b'', data)
                    print("Cleaned [Content_Types].xml")
                elif item.filename == 'xl/_rels/workbook.xml.rels':
                    # Remove calcChain Relationship tag
                    data = re.sub(rb'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', b'', data)
                    print("Cleaned xl/_rels/workbook.xml.rels")
                zout.writestr(item, data)
    os.replace(temp_zip, xlsx_path)
    print("Done removing calcChain from", xlsx_path)

test_target = 'scratch/test_populate_ds_sunpharma.xlsx'
remove_calc_chain(test_target)

with zipfile.ZipFile(test_target, 'r') as z:
    names = z.namelist()
    print("Has calcChain:", any('calcChain' in n for n in names))
    # verify xml
    for n in ['[Content_Types].xml', 'xl/_rels/workbook.xml.rels']:
        try:
            ET.fromstring(z.read(n))
            print(f"{n} parsed successfully!")
        except Exception as e:
            print(f"Error parsing {n}: {e}")
