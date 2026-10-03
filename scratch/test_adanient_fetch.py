import sys
sys.path.insert(0, '.')
import screener_client

sd = screener_client.fetch_company_data("ADANIENT")
print("Company Name:", sd.get('company_name'))
print("Ticker:", sd.get('ticker'))
print("Tables:", list(sd.get('tables', {}).keys()))
pl = sd.get('tables', {}).get('profit-loss')
if pl is not None:
    print("PL cols:", pl.columns.tolist())
    sales_m = pl[pl['Metric'].str.contains('Sales|Revenue', case=False, na=False)]
    if not sales_m.empty:
        print("Sales row:", sales_m.iloc[0].to_dict())
bs = sd.get('tables', {}).get('balance-sheet')
if bs is not None:
    print("BS cols:", bs.columns.tolist())
    eq_m = bs[bs['Metric'].str.contains('Equity Capital|Share Capital', case=False, na=False)]
    if not eq_m.empty:
        print("Equity row:", eq_m.iloc[0].to_dict())
