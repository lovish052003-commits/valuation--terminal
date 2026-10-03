import zipfile

print("=== TEMPLATE (ITC Model.xlsx) ===")
with zipfile.ZipFile('ITC Model.xlsx', 'r') as z:
    for i in range(1, 9):
        name = f"xl/drawings/drawing{i}.xml"
        if name in z.namelist():
            data = z.read(name).decode('utf-8', errors='ignore')
            print(f"{name} ({len(data)} bytes):")
            print(data[:300])

print("\n=== EXPORT (UNITEDTEA_Valuation_Model.xlsx) ===")
path = r'exports\UNITEDTEA_Valuation_Model.xlsx'
with zipfile.ZipFile(path, 'r') as z:
    for i in range(1, 9):
        name = f"xl/drawings/drawing{i}.xml"
        if name in z.namelist():
            data = z.read(name).decode('utf-8', errors='ignore')
            print(f"{name} ({len(data)} bytes):")
            print(data[:300])
