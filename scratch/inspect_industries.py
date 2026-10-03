import openpyxl
from collections import Counter

wb = openpyxl.load_workbook('ITC Model.xlsx', read_only=True)
ws = wb['List of Stocks']

industries = Counter()
samples_by_ind = {}

for row in ws.iter_rows(values_only=True):
    if not row or len(row) < 5 or row[1] == 'COMPANY_NAME':
        continue
    c_name = row[1]
    bse = row[2]
    nse = row[3]
    ind = row[4]
    if ind:
        industries[ind] += 1
        if ind not in samples_by_ind and nse and str(nse).strip() not in ('None', 'Not Listed', ''):
            samples_by_ind[ind] = {'name': c_name, 'bse': bse, 'nse': nse, 'industry': ind}

wb.close()

print(f"Total unique industries: {len(industries)}")
print("\nTop 25 Industries by company count:")
for ind, count in industries.most_common(25):
    sample = samples_by_ind.get(ind, {})
    print(f"  {ind:40s}: {count:4d} companies | Sample: {sample.get('nse')} ({sample.get('name')})")
