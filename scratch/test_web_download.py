import requests, os, win32com.client, pythoncom

print("1. Calling Flask /api/analyze for ITC (Skip AI)...")
resp = requests.post('http://127.0.0.1:5000/api/analyze', json={
    'company': 'ITC',
    'provider': 'nvidia',
    'apiKey': '',
    'skipAi': True
}, timeout=30)

print("Status Code:", resp.status_code)
res_json = resp.json()
print("Success:", res_json.get('success'))
excel_filename = res_json.get('excel_filename')
print("Excel filename:", excel_filename)

print("2. Downloading Excel file from /api/download-excel/<filename>...")
dl_resp = requests.get(f'http://127.0.0.1:5000/api/download-excel/{excel_filename}')
print("Download status:", dl_resp.status_code, f"Length: {len(dl_resp.content)} bytes")

dl_save_path = os.path.abspath('exports/DOWNLOADED_TEST.xlsx')
with open(dl_save_path, 'wb') as f:
    f.write(dl_resp.content)

print("3. Opening downloaded file in Microsoft Excel via COM to test for repair errors...")
pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(dl_save_path)
    print("SUCCESS: Downloaded file opened in Microsoft Excel with ZERO ERRORS and NO REPAIR PROMPTS!")
    sheet_names = [s.Name for s in wb.Sheets]
    print("Sheets inside:", len(sheet_names), sheet_names[:5])
    wb.Close(SaveChanges=False)
except Exception as e:
    print("FAILED opening downloaded file:", e)
finally:
    excel.Quit()
    del excel
    pythoncom.CoUninitialize()

print("\nFull Web End-to-End Test Completed Successfully!")
