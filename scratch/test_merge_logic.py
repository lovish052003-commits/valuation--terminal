import requests
import pandas as pd
from io import StringIO
from lxml import html
import re

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
})

def extract_df(tree, sec_id):
    sec = tree.xpath(f'//section[@id="{sec_id}"]')
    if not sec:
        return None
    tables = sec[0].xpath('.//table')
    if not tables:
        return None
    df = pd.read_html(StringIO(html.tostring(tables[0]).decode('utf-8')))[0]
    first_col = df.columns[0]
    df.rename(columns={first_col: 'Metric'}, inplace=True)
    df['Metric'] = df['Metric'].astype(str).str.replace(r'[\+\xa0\r\n\t]', '', regex=True).str.strip()
    return df

r_c = session.get('https://www.screener.in/company/NESTLEIND/consolidated/')
r_s = session.get('https://www.screener.in/company/NESTLEIND/')
t_c = html.fromstring(r_c.content)
t_s = html.fromstring(r_s.content)

def clean_period_name(col):
    # e.g. "Mar 2024  15m" -> "Mar 2024 15m", or normalized form
    return re.sub(r'\s+', ' ', str(col)).strip()

def merge_statement_dfs(df_cons, df_std):
    if df_std is None or df_std.empty:
        return df_cons
    if df_cons is None or df_cons.empty:
        return df_std

    # Clean column names in both
    cons_cols = [c for c in df_cons.columns if c not in ('Metric', 'TTM')]
    std_cols = [c for c in df_std.columns if c not in ('Metric', 'TTM')]
    
    # Check if consolidated needs backfilling (e.g. fewer than 10 periods or has NaNs)
    print(f"Cons cols count: {len(cons_cols)}, Std cols count: {len(std_cols)}")
    
    # We want a unified set of periods.
    # Typically, std has the earlier periods, and cons has the latest periods.
    # Let's map normalized period key (e.g. 'mar 2024' from 'mar 2024 15m')
    def get_norm_key(c):
        s = re.sub(r'\s*\d+m\b', '', str(c), flags=re.I).strip().lower()
        return re.sub(r'\s+', ' ', s)

    # Let's align on Metric
    # Create a merged DataFrame starting from std_df
    merged_df = df_std.copy()
    
    # Map metrics in cons to rows in merged_df
    for _, cons_row in df_cons.iterrows():
        c_metric = cons_row['Metric']
        # Find matching row in merged_df
        m_idx = merged_df[merged_df['Metric'].str.lower() == c_metric.lower()].index
        if len(m_idx) > 0:
            target_idx = m_idx[0]
            # Overlay non-null values from cons_row
            for c_col in df_cons.columns:
                if c_col == 'Metric':
                    continue
                c_val = cons_row[c_col]
                if pd.notna(c_val) and str(c_val).strip() not in ('', 'nan', 'NaN', '-'):
                    # Find corresponding column in merged_df
                    # If exact col name exists in merged_df:
                    if c_col in merged_df.columns:
                        merged_df.at[target_idx, c_col] = c_val
                    else:
                        # Match by norm_key (e.g. 'Mar 2024 15m' matching 'Mar 2024')
                        matched = False
                        c_key = get_norm_key(c_col)
                        for s_col in merged_df.columns:
                            if s_col == 'Metric':
                                continue
                            if get_norm_key(s_col) == c_key:
                                merged_df.at[target_idx, s_col] = c_val
                                matched = True
                                break
                        if not matched:
                            # Add new column
                            merged_df[c_col] = None
                            merged_df.at[target_idx, c_col] = c_val

    # Also make sure TTM column is preserved if in cons
    if 'TTM' in df_cons.columns:
        for _, cons_row in df_cons.iterrows():
            c_metric = cons_row['Metric']
            m_idx = merged_df[merged_df['Metric'].str.lower() == c_metric.lower()].index
            if len(m_idx) > 0:
                merged_df.at[m_idx[0], 'TTM'] = cons_row['TTM']

    return merged_df

for sec in ['profit-loss', 'balance-sheet', 'cash-flow']:
    c_df = extract_df(t_c, sec)
    s_df = extract_df(t_s, sec)
    merged = merge_statement_dfs(c_df, s_df)
    print(f"\n=== MERGED {sec.upper()} ===")
    print("Columns:", merged.columns.tolist())
    print(merged.head(3))
