import requests
import json
import zipfile
import os
from PIL import Image
from io import BytesIO

url = 'http://127.0.0.1:5000/api/analyze'
payload = {
    'company': 'Tata Steel',
    'skipAi': True
}

print("1. Sending analyze request for Tata Steel...")
r = requests.post(url, json=payload, timeout=60)
data = r.json()
print("   Success:", data.get('success'))
print("   Company:", data.get('company', {}).get('name'))
print("   Logo URL:", data.get('company', {}).get('logo_url'))
excel_filename = data.get('excel_filename')
print("   Excel Filename:", excel_filename)

excel_path = os.path.join('exports', excel_filename)
if os.path.exists(excel_path):
    print("\n2. Inspecting media images in exported Excel file:", excel_path)
    with zipfile.ZipFile(excel_path, 'r') as z:
        for m in ['xl/media/image3.png', 'xl/media/image2.png']:
            if m in z.namelist():
                img_data = z.read(m)
                im = Image.open(BytesIO(img_data))
                print(f"   {m}: size={im.size}, mode={im.mode}, length={len(img_data)} bytes")
                
print("\nLogo test completed!")
