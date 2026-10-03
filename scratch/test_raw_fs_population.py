import os, sys, shutil, json, requests, gc
import win32com.client, pythoncom
from lxml import html
import pandas as pd
from io import StringIO

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from screener_client import fetch_company_data, clean_num, HEADERS
from valuation_engine import calculate_valuation

def fetch_schedules_for_company(company_id, is_consolidated=True):
    """Fetches sub-schedules for Balance Sheet and Cash Flow from Screener.in."""
    schedules = {}
    if not company_id:
        return schedules

    schedule_targets = [
        ('Fixed Assets', 'balance-sheet'),
        ('Other Assets', 'balance-sheet'),
        ('Borrowings', 'balance-sheet'),
        ('Other Liabilities', 'balance-sheet'),
        ('Cash from Operating Activity', 'cash-flow'),
        ('Cash from Investing Activity', 'cash-flow'),
        ('Cash from Financing Activity', 'cash-flow')
    ]
    for parent, section in schedule_targets:
        try:
            cons_param = "&consolidated=" if is_consolidated else ""
            sch_url = f"https://www.screener.in/api/company/{company_id}/schedules/?parent={requests.utils.quote(parent)}&section={section}{cons_param}"
            s_resp = requests.get(sch_url, headers=HEADERS, timeout=8)
            if s_resp.status_code == 200 and s_resp.text.startswith('{'):
                schedules[parent] = s_resp.json()
                print(f"  Fetched schedule '{parent}' with {len(schedules[parent])} items")
        except Exception as e:
            print(f"Notice: Failed fetching schedule {parent}: {e}")
    return schedules

# Test with Tata Motors
symbol = 'Tata Motors'
print(f"1. Fetching company data for {symbol}...")
data = fetch_company_data(symbol)
cid = data.get('company_id')
print(f"Company ID: {cid}, Name: {data['company_name']}")

print("2. Fetching sub-schedules...")
schedules = fetch_schedules_for_company(cid, is_consolidated=True)
data['schedules'] = schedules

print("Schedules fetched successfully!")
