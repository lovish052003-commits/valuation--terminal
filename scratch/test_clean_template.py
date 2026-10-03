import zipfile, os, re, xml.etree.ElementTree as ET

src_path = 'ITC Model.xlsx'
clean_path = 'scratch/test_cleaned_model.xlsx'

with zipfile.ZipFile(src_path, 'r') as zin, zipfile.ZipFile(clean_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        filename = item.filename
        
        # 1. Skip externalLink files
        if 'externalLink' in filename:
            continue
            
        data = zin.read(filename)
        
        # 2. In workbook.xml, remove externalReferences
        if filename == 'xl/workbook.xml':
            text = data.decode('utf-8')
            text = re.sub(r'<externalReferences>.*?</externalReferences>', '', text)
            data = text.encode('utf-8')
            
        # 3. In xl/_rels/workbook.xml.rels, remove externalLink relationship and normalize targets
        elif filename == 'xl/_rels/workbook.xml.rels':
            text = data.decode('utf-8')
            text = re.sub(r'<Relationship [^>]*externalLink[^>]*/>', '', text)
            # Normalize /xl/worksheets/ to worksheets/
            text = text.replace('Target="/xl/', 'Target="')
            data = text.encode('utf-8')
            
        # 4. In [Content_Types].xml, remove externalLink override
        elif filename == '[Content_Types].xml':
            text = data.decode('utf-8')
            text = re.sub(r'<Override [^>]*externalLink[^>]*/>', '', text)
            data = text.encode('utf-8')
            
        # 5. In worksheet rels, normalize Target="/xl/drawings/..." to "../drawings/..."
        elif filename.startswith('xl/worksheets/_rels/'):
            text = data.decode('utf-8')
            text = text.replace('Target="/xl/drawings/', 'Target="../drawings/')
            data = text.encode('utf-8')
            
        # 6. In drawing rels, normalize Target="/xl/media/..." to "../media/..." and "/xl/charts/..." to "../charts/..."
        elif filename.startswith('xl/drawings/_rels/'):
            text = data.decode('utf-8')
            text = text.replace('Target="/xl/media/', 'Target="../media/')
            text = text.replace('Target="/xl/charts/', 'Target="../charts/')
            data = text.encode('utf-8')
            
        zout.writestr(item, data)

print("Created cleaned template:", clean_path)

# Test opening with Excel COM
import win32com.client
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(os.path.abspath(clean_path), UpdateLinks=0, ReadOnly=False)
    print("SUCCESS: Opened cleaned workbook with Excel COM without any error!")
    print("Sheets count:", len(wb.Sheets))
    wb.Close(False)
except Exception as e:
    print("FAILED to open with Excel COM:", e)
finally:
    excel.Quit()
