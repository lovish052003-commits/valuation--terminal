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
    df.columns = [re.sub(r'\s+', ' ', str(c)).strip() if c != 'Metric' else 'Metric' for c in df.columns]
    return df

def _norm_period_key(p):
    s = re.sub(r'[\+\xa0\r\n\t]', ' ', str(p))
    s = re.sub(r'\s*\d+m\b', '', s, flags=re.I).strip().lower()
    return re.sub(r'\s+', ' ', s)

def merge_statement_dfs(df_cons, df_std):
    if df_std is None or df_std.empty:
        return df_cons
    if df_cons is None or df_cons.empty:
        return df_std

    merged_df = df_std.copy()
    if 'Metric' in merged_df.columns:
        merged_df['Metric'] = merged_df['Metric'].astype(str).str.replace(r'[\+\xa0\r\n\t]', '', regex=True).str.strip()
    if 'Metric' in df_cons.columns:
        df_cons_clean = df_cons.copy()
        df_cons_clean['Metric'] = df_cons_clean['Metric'].astype(str).str.replace(r'[\+\xa0\r\n\t]', '', regex=True).str.strip()
    else:
        df_cons_clean = df_cons

    for _, c_row in df_cons_clean.iterrows():
        c_metric = str(c_row.get('Metric', '')).strip()
        if not c_metric:
            continue
        m_idx = merged_df[merged_df['Metric'].str.lower() == c_metric.lower()].index
        if len(m_idx) == 0:
            m_idx = merged_df[merged_df['Metric'].str.lower().str.contains(re.escape(c_metric[:8].lower()), case=False, na=False)].index
        if len(m_idx) > 0:
            target_idx = m_idx[0]
            for c_col in df_cons_clean.columns:
                if c_col == 'Metric':
                    continue
                c_val = c_row[c_col]
                if pd.notna(c_val) and str(c_val).strip() not in ('', 'nan', 'NaN', '-'):
                    if c_col in merged_df.columns:
                        merged_df.at[target_idx, c_col] = c_val
                    else:
                        c_key = _norm_period_key(c_col)
                        matched = False
                        for s_col in merged_df.columns:
                            if s_col != 'Metric' and _norm_period_key(s_col) == c_key:
                                merged_df.at[target_idx, s_col] = c_val
                                matched = True
                                break
                        if not matched:
                            merged_df[c_col] = None
                            merged_df.at[target_idx, c_col] = c_val
        else:
            # Append new row
            new_row = {col: None for col in merged_df.columns}
            new_row['Metric'] = c_metric
            for c_col in df_cons_clean.columns:
                if c_col in merged_df.columns:
                    new_row[c_col] = c_row[c_col]
                else:
                    c_key = _norm_period_key(c_col)
                    for s_col in merged_df.columns:
                        if s_col != 'Metric' and _norm_period_key(s_col) == c_key:
                            new_row[s_col] = c_row[c_col]
                            break
            merged_df = pd.concat([merged_df, pd.DataFrame([new_row])], ignore_index=True)

    if 'TTM' in df_cons.columns:
        for _, c_row in df_cons_clean.iterrows():
            c_metric = str(c_row.get('Metric', '')).strip().lower()
            m_idx = merged_df[merged_df['Metric'].str.lower() == c_metric].index
            if len(m_idx) > 0:
                merged_df.at[m_idx[0], 'TTM'] = c_row['TTM']

    return merged_df

r_c = session.get('https://www.screener.in/company/NESTLEIND/consolidated/')
r_s = session.get('https://www.screener.in/company/NESTLEIND/')
t_c = html.fromstring(r_c.content)
t_s = html.fromstring(r_s.content)

print("Testing BS merge:")
bs_c = extract_df(t_c, 'balance-sheet')
bs_s = extract_df(t_s, 'balance-sheet')
bs_m = merge_statement_dfs(bs_c, bs_s)
print("Merged BS columns:", bs_m.columns.tolist())
print(bs_m.to_string())
