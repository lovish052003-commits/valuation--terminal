import urllib.request, zipfile
from io import BytesIO

url = 'http://127.0.0.1:5000/api/download-excel/JUBLFOOD_Valuation_Model.xlsx'
with urllib.request.urlopen(url) as resp:
    data = resp.read()
    print('Download Status:', resp.status, 'Bytes:', len(data))
    with zipfile.ZipFile(BytesIO(data), 'r') as z:
        for name in z.namelist():
            if 'media' in name:
                info = z.getinfo(name)
                is_itc = (info.file_size == 18741)
                tag = '[FAIL: ITC LOGO]' if is_itc else '[PASS: JUBILANT LOGO]'
                print(f"{name}: size={info.file_size} {tag}")
