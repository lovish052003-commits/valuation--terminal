import os, sys
sys.path.insert(0, os.path.abspath('.'))
from screener_client import load_offline_company_data
from screener_client import clean_num
import numpy as np

def run_test(ticker):
    d = load_offline_company_data(ticker)
    tables = d.get('tables', {})
    pl_df = tables.get('profit-loss')
    bs_df = tables.get('balance-sheet')
    
    # Check what screener_client has
    cash = clean_num(d.get('cash_cr'))
    inv = clean_num(d.get('investments_cr'))
    debt = clean_num(d.get('debt_cr'))
    cmp = clean_num(d.get('current_price'))
    shares = clean_num(d.get('shares_in_cr'))
    print(f"[{ticker}] Screener cached: CMP={cmp}, Shares={shares}, Cash={cash}, Inv={inv}, Debt={debt}")

run_test('NESTLEIND')
run_test('ITC')
