import zipfile, os, re

src = 'exports/TEST_MARUTI_COM.xlsx'
dest = 'scratch/test_pristine_template.xlsx'

with zipfile.ZipFile(src, 'r') as zin, zipfile.ZipFile(dest, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        filename = item.filename
        
        # 1. Skip externalLink files completely
        if 'externalLink' in filename:
            continue
            
        data = zin.read(filename)
        
        # 2. In workbook.xml, remove externalReferences
        if filename == 'xl/workbook.xml':
            text = data.decode('utf-8')
            text = re.sub(r'<externalReferences>.*?</externalReferences>', '', text)
            data = text.encode('utf-8')
            
        # 3. In xl/_rels/workbook.xml.rels, remove externalLink relationship
        elif filename == 'xl/_rels/workbook.xml.rels':
            text = data.decode('utf-8')
            text = re.sub(r'<Relationship [^>]*externalLink[^>]*/>', '', text)
            data = text.encode('utf-8')
            
        # 4. In [Content_Types].xml, remove externalLink override
        elif filename == '[Content_Types].xml':
            text = data.decode('utf-8')
            text = re.sub(r'<Override [^>]*externalLink[^>]*/>', '', text)
            data = text.encode('utf-8')
            
        zout.writestr(item, data)

print("Created pristine template without externalLinks:", dest)

# Test opening with Excel COM
import win32com.client
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(os.path.abspath(dest), UpdateLinks=0, ReadOnly=False)
    print("SUCCESS: Excel COM opened pristine template with NO error!")
    print("Sheets count:", len(wb.Sheets))
    print("Sheet names:", [s.Name for s in wb.Sheets])
    wb.Close(False)
except Exception as e:
    print("FAILED:", e)
finally:
    excel.Quit()
