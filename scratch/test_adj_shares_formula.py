import sys, os
sys.path.insert(0, os.path.abspath('.'))
from screener_client import fetch_company_data, clean_num
import openpyxl

wb = openpyxl.load_workbook(r'C:\Users\LENOVO\Downloads\Nestle India (2).xlsx', data_only=True)
ws = wb['Data Sheet']

actual_adj_shares = [ws.cell(93, c).value for c in range(2, 12)]
print("Actual Row 93 in Nestle India (2).xlsx:")
print(actual_adj_shares)

data = fetch_company_data('NESTLEIND')
pl_df = data['tables']['profit-loss']
print("\nPL DF Metric rows:")
print(pl_df['Metric'].tolist())

net_profit_row = pl_df[pl_df['Metric'].str.contains('Net Profit', case=False, na=False)].iloc[0]
eps_row = pl_df[pl_df['Metric'].str.contains('EPS', case=False, na=False)].iloc[0]

period_cols = [c for c in pl_df.columns if c != 'Metric'][-10:]
calc_adj_shares = []
for p in period_cols:
    np_val = clean_num(net_profit_row[p])
    eps_val = clean_num(eps_row[p])
    shares = round(np_val / eps_val, 2) if eps_val > 0 else 0.0
    calc_adj_shares.append((p, np_val, eps_val, shares))

print("\nCalculated Adjusted Shares (Net Profit / EPS):")
for p, np_val, eps_val, shares in calc_adj_shares:
    print(f"Period: {p:10s} | NP: {np_val:10.2f} | EPS: {eps_val:6.2f} | Shares Cr: {shares:8.2f}")
