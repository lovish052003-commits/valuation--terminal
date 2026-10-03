import sys, os
sys.path.insert(0, os.path.abspath('.'))
import screener_client

data = screener_client.fetch_company_data('SUNPHARMA')
print("Company:", data['company_name'])
print("Tables:", list(data['tables'].keys()))
pl = data['tables'].get('profit-loss')
if pl is not None:
    print("P&L columns:", list(pl.columns))
    print("P&L metrics:\n", pl['Metric'].tolist())

print("Schedules fetched:", list(data.get('schedules', {}).keys()))
exp = data.get('schedules', {}).get('Expenses', {})
print("Expenses schedule items:", list(exp.keys()))
mat = data.get('schedules', {}).get('Material Cost %', {})
print("Material Cost % schedule items:", list(mat.keys()))
if mat:
    print("Raw material cost sample:", list(mat.get('Raw material cost', {}).items())[:3])
