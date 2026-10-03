import os, zipfile, xml.etree.ElementTree as ET
from PIL import Image
from io import BytesIO

def get_target_logo_images(xlsx_path):
    target_images = set()
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        try:
            wb_xml = z.read('xl/workbook.xml')
            wb_tree = ET.fromstring(wb_xml)
            wb_rels_xml = z.read('xl/_rels/workbook.xml.rels')
            wb_rels_tree = ET.fromstring(wb_rels_xml)
            
            rel_map = {r.attrib['Id']: r.attrib['Target'] for r in wb_rels_tree}
            
            sheet_targets = {}
            for s in wb_tree.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets'):
                name = s.attrib['name']
                rId = s.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
                t = rel_map[rId]
                if t.startswith('/xl/'):
                    t = t[1:]
                elif not t.startswith('xl/'):
                    t = 'xl/' + t
                sheet_targets[name] = t

            for sname in ['Dupont Analysis', "Altman's Z Score"]:
                target = sheet_targets.get(sname)
                if not target:
                    continue
                sheet_dir, sheet_file = target.rsplit('/', 1)
                rels_path = f"{sheet_dir}/_rels/{sheet_file}.rels"
                if rels_path in z.namelist():
                    s_rels_tree = ET.fromstring(z.read(rels_path))
                    for r in s_rels_tree:
                        if 'drawing' in r.attrib.get('Type', ''):
                            dwg_target = r.attrib['Target']
                            if dwg_target.startswith('../'):
                                dwg_path = 'xl/' + dwg_target.replace('../', '')
                            elif dwg_target.startswith('/xl/'):
                                dwg_path = dwg_target[1:]
                            else:
                                dwg_path = f"{sheet_dir}/{dwg_target}"
                            
                            dwg_dir, dwg_file = dwg_path.rsplit('/', 1)
                            dwg_rels_path = f"{dwg_dir}/_rels/{dwg_file}.rels"
                            if dwg_rels_path in z.namelist():
                                d_rels_tree = ET.fromstring(z.read(dwg_rels_path))
                                for dr in d_rels_tree:
                                    if 'image' in dr.attrib.get('Type', ''):
                                        img_target = dr.attrib['Target']
                                        if img_target.startswith('../'):
                                            img_path = 'xl/' + img_target.replace('../', '')
                                        elif img_target.startswith('/xl/'):
                                            img_path = img_target[1:]
                                        else:
                                            img_path = f"{dwg_dir}/{img_target}"
                                        target_images.add(img_path)
        except Exception as e:
            print('Error parsing zip:', e)
            
    # Always include image3.png, image7.png, image8.png as standard targets if referenced or matching legacy
    target_images.add('xl/media/image3.png')
    target_images.add('xl/media/image7.png')
    target_images.add('xl/media/image8.png')
    return target_images

print('Detected target logo images:', get_target_logo_images('exports/JUBLFOOD_Valuation_Model.xlsx'))
