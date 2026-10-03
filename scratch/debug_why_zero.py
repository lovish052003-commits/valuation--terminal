import sys, os
sys.path.insert(0, os.path.abspath('.'))
import screener_client

screener_data = screener_client.fetch_company_data('SUNPHARMA')
schedules = screener_data.get('schedules', {})
pl_df = screener_data['tables'].get('profit-loss')

print("schedules keys:", list(schedules.keys()))
exp_sch = schedules.get('Expenses', {})
mat_sch = schedules.get('Material Cost %', {})
print("exp_sch keys:", list(exp_sch.keys()))
print("mat_sch keys:", list(mat_sch.keys()))

period_cols = [c for c in pl_df.columns if c not in ('Metric', 'TTM')][-10:]
print("period_cols from pl_df:", period_cols)

raw_mat_dict = mat_sch.get('Raw material cost', {})
print("raw_mat_dict keys:", list(raw_mat_dict.keys()))

for col in period_cols:
    in_raw = col in raw_mat_dict
    raw_val = raw_mat_dict.get(col)
    print(f"Col: {col!r} | in raw_mat_dict: {in_raw} | raw_val: {raw_val}")

print("\n--- exp_sch check ---")
for item_k, item_dict in exp_sch.items():
    print(f"item_k: {item_k!r}, keys: {list(item_dict.keys())[:3]}")
    for col in period_cols[:3]:
        val = item_dict.get(col)
        print(f"   col: {col!r} -> val: {val!r}")
