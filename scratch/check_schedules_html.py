import screener_client, re

session = screener_client.get_screener_session()
for sym in ['WIPRO', 'BRITANNIA', 'SUZLON', 'TCS', 'HDFCBANK']:
    resp = session.get(f'https://www.screener.in/company/{sym}/consolidated/')
    if resp.status_code != 200:
        resp = session.get(f'https://www.screener.in/company/{sym}/')
    matches = re.findall(r'Company\.showSchedule\([\'"]([^\'"]+)[\'"]\s*,\s*[\'"]([^\'"]+)[\'"]', resp.text)
    print(f'{sym} schedules in HTML ({len(matches)}): {matches}')
