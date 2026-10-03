import re
from screener_client import get_screener_session, fetch_company_data
import sys
sys.stdout.reconfigure(encoding='utf-8')

session = get_screener_session()
r = session.get('https://www.screener.in/company/HINDUNILVR/consolidated/')
pattern = r'Company\.showSchedule\([\'"]([^\'"]+)[\'"]\s*,\s*[\'"]([^\'"]+)[\'"]'
matches = re.findall(pattern, r.text)
print('Matches count:', len(matches))
for m in matches:
    print(' ', m)
