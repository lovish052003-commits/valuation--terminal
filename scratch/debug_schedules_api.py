import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client
import requests

session = screener_client.get_screener_session()

# Let's inspect WIPRO and BRITANNIA company page and schedules
for sym in ['WIPRO', 'BRITANNIA', 'INFY']:
    print(f"\n==================== {sym} ====================")
    resp = session.get(f"https://www.screener.in/company/{sym}/consolidated/")
    if resp.status_code != 200:
        resp = session.get(f"https://www.screener.in/company/{sym}/")
    
    # Get company_id
    import re
    m_id = re.search(r'data-company-id=[\'"](\d+)[\'"]', resp.text)
    cid = m_id.group(1) if m_id else None
    print(f"Company ID: {cid}")
    
    # Check is_consolidated
    is_cons = 'data-consolidated' in resp.text
    print(f"Is consolidated attribute in HTML: {is_cons}")

    schedule_targets = [
        ('Fixed Assets', 'balance-sheet'),
        ('Other Assets', 'balance-sheet'),
        ('Borrowings', 'balance-sheet'),
        ('Other Liabilities', 'balance-sheet'),
        ('Cash from Operating Activity', 'cash-flow'),
        ('Cash from Investing Activity', 'cash-flow'),
        ('Cash from Financing Activity', 'cash-flow'),
        ('Expenses', 'profit-loss'),
        ('Material Cost %', 'profit-loss')
    ]
    
    for parent, sec_name in schedule_targets:
        for cons_flag in ["&consolidated=", ""]:
            sch_url = f"https://www.screener.in/api/company/{cid}/schedules/?parent={requests.utils.quote(parent)}&section={sec_name}{cons_flag}"
            try:
                r = session.get(sch_url, timeout=10)
                is_json = r.text.startswith('{')
                first_50 = r.text[:60].replace('\n', ' ')
                print(f"  {parent} (cons='{cons_flag}'): status={r.status_code}, is_json={is_json}, snippet={first_50}")
            except Exception as e:
                print(f"  {parent} (cons='{cons_flag}'): ERROR {e}")
