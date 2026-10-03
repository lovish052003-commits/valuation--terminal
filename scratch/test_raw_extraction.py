import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client

data = screener_client.fetch_company_data('ADANIENT')
tables = data.get('tables', {})
schedules = data.get('schedules', {})
meta = data.get('meta_raw', {})

# 1. PBT and Interest from profit-loss
pl_df = tables.get('profit-loss')
pbt_row = pl_df[pl_df['Metric'].str.contains('Profit before tax|PBT', case=False, na=False)]
int_row = pl_df[pl_df['Metric'].str.contains('Interest', case=False, na=False)]
pl_cols = [c for c in pl_df.columns if c != 'Metric' and 'TTM' not in c]
latest_pbt = screener_client.clean_num(pbt_row.iloc[0][pl_cols[-1]]) if not pbt_row.empty else 0
latest_interest = screener_client.clean_num(int_row.iloc[0][pl_cols[-1]]) if not int_row.empty else 0
base_ebit = latest_pbt + latest_interest

print("latest_pbt:", latest_pbt)
print("latest_interest:", latest_interest)
print("base_ebit:", base_ebit)

# 2. Cash from schedules
oa_sched = schedules.get('Other Assets', {})
cash_dict = oa_sched.get('Cash Equivalents', {})
bs_df = tables.get('balance-sheet')
bs_cols = [c for c in bs_df.columns if c != 'Metric']
latest_bs_col = bs_cols[-1]
cash_val = None
for sk, sv in cash_dict.items():
    if sk.strip().lower() in latest_bs_col.lower() or latest_bs_col.lower() in sk.strip().lower():
        cash_val = screener_client.clean_num(sv)
        break
if cash_val is None and cash_dict:
    cash_val = screener_client.clean_num(list(cash_dict.values())[-1])
print("cash_val from schedule:", cash_val)

# 3. Debt from balance sheet
debt_row = bs_df[bs_df['Metric'].str.contains('Borrowings|Total Debt', case=False, na=False)]
total_debt = screener_client.clean_num(debt_row.iloc[0][latest_bs_col]) if not debt_row.empty else 0
print("total_debt:", total_debt)

# 4. Shares from equity capital / face value
eq_row = bs_df[bs_df['Metric'].str.contains('Equity Capital|Share Capital', case=False, na=False)]
latest_eq = screener_client.clean_num(eq_row.iloc[0][latest_bs_col]) if not eq_row.empty else 0
face_value = screener_client.clean_num(data.get('face_value', 1.0)) or 1.0
shares_cr = latest_eq / face_value
print("equity capital:", latest_eq, "face_value:", face_value, "shares_cr:", shares_cr)
