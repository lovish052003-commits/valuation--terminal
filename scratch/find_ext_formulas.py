import zipfile

with zipfile.ZipFile('ITC Model.xlsx', 'r') as z:
    data = z.read('xl/worksheets/sheet22.xml').decode('utf-8', errors='ignore')
    idx = data.find('Tata Steel')
    if idx != -1:
        print("Found 'Tata Steel' at:", idx)
        print(data[max(0, idx-100):min(len(data), idx+100)])
    else:
        print("'Tata Steel' not found")
    idx2 = data.find('[1]')
    if idx2 != -1:
        print("Found '[1]' at:", idx2)
        print(data[max(0, idx2-100):min(len(data), idx2+100)])
    else:
        print("'[1]' not found")
