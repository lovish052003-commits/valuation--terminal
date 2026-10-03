import requests, io, openpyxl

r = requests.post('http://127.0.0.1:5000/api/analyze', json={'company': 'BRITANNIA', 'skipAi': True}, timeout=60)
print('Status:', r.status_code)
res = r.json()
if r.status_code != 200:
    print('Error:', res)
    exit(1)

excel_fn = res.get('excel_filename')
print('Excel filename:', excel_fn)

r2 = requests.get(f'http://127.0.0.1:5000/api/download-excel/{excel_fn}', timeout=30)
print('Downloaded bytes:', len(r2.content))

wb = openpyxl.load_workbook(io.BytesIO(r2.content), read_only=True)
print('Sheet count:', len(wb.sheetnames))
print('Sheets:', wb.sheetnames)
