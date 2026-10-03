import os, sys, gc
import win32com.client, pythoncom
import pandas as pd
from datetime import datetime
import openpyxl

sys.path.insert(0, os.path.abspath('.'))
from screener_client import fetch_company_data, clean_num

def test_populate_complete_data_sheet():
    symbol = 'NESTLEIND'
    print(f"Fetching data for {symbol}...")
    screener_data = fetch_company_data(symbol)
    
    # 1. Fetch historical prices
    cid = screener_data.get('company_id')
    import requests
    price_series = []
    if cid:
        url = f'https://www.screener.in/api/company/{cid}/chart/?q=Price-DMA50-DMA200-Volume&days=3650'
        r = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'})
        if r.status_code == 200:
            data = r.json()
            for ds in data.get('datasets', []):
                if ds.get('metric') == 'Price':
                    for pt in ds.get('values', []):
                        try:
                            dt = datetime.strptime(pt[0], '%Y-%m-%d')
                            price_series.append((dt, float(pt[1])))
                        except Exception:
                            pass
                    break
    price_series.sort(key=lambda x: x[0])
    print(f"Fetched {len(price_series)} historical price points.")

    # Let's inspect what we write to each cell
    tables = screener_data.get('tables', {})
    pl_df = tables.get('profit-loss')
    bs_df = tables.get('balance-sheet')
    cf_df = tables.get('cash-flow')
    q_df = tables.get('quarters')
    schedules = screener_data.get('schedules', {})

    bs_periods = [c for c in bs_df.columns if c != 'Metric'][-10:]
    pl_periods = [c for c in pl_df.columns if c != 'Metric'][-10:]
    cf_periods = [c for c in cf_df.columns if c != 'Metric'][-10:]
    
    print("BS periods:", bs_periods)
    print("PL periods:", pl_periods)

    # Let's check schedules for Other Assets
    oa = schedules.get('Other Assets', {})
    print("Other Assets schedules:", list(oa.keys()))

    # Calculate Adjusted Shares
    np_row = pl_df[pl_df['Metric'].str.contains('Net Profit', case=False, na=False)].iloc[0]
    eps_row = pl_df[pl_df['Metric'].str.contains('EPS', case=False, na=False)].iloc[0]
    curr_shares = clean_num(screener_data.get('shares_in_cr', 0))

    adj_shares_list = []
    for p in pl_periods:
        np_v = clean_num(np_row.get(p, 0))
        eps_v = clean_num(eps_row.get(p, 0))
        if eps_v > 0 and np_v > 0:
            sh = round(np_v / eps_v, 2)
        else:
            sh = curr_shares
        adj_shares_list.append(sh)
    print("Derived Adjusted Shares:", adj_shares_list)

test_populate_complete_data_sheet()
