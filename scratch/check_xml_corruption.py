import openpyxl, os, zipfile, xml.etree.ElementTree as ET

path = r'c:\Users\LENOVO\Downloads\Advance Financial Project\exports\SUNPHARMA_Valuation_Model.xlsx'
with zipfile.ZipFile(path, 'r') as z:
    for name in z.namelist():
        if name.endswith('.xml') or name.endswith('.rels'):
            try:
                data = z.read(name)
                ET.fromstring(data)
            except Exception as e:
                print(f"Error in {name}: {e}")

print("Checking sheets and drawings...")
with zipfile.ZipFile(path, 'r') as z:
    for name in z.namelist():
        if 'drawings' in name:
            print("Drawing file:", name, len(z.read(name)))
        if 'calcChain' in name:
            print("calcChain file:", name, len(z.read(name)))
