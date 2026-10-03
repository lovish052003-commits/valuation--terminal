import requests

r = requests.post('http://127.0.0.1:5000/api/analyze', json={'company': 'TATASTEEL', 'skipAi': True}, timeout=60)
data = r.json()
print("Success:", data.get('success'))
comp = data.get('company', {})
print("Company Name:", comp.get('name'))
print("Logo URL in company:", comp.get('logo_url'))
assert comp.get('logo_url') is None, "logo_url must NOT be in API response"

val = data.get('valuation', {})
print("Valuation Tax Rate:", val.get('tax_rate'))
assert val.get('tax_rate') == 30.0, f"Tax rate expected 30.0, got {val.get('tax_rate')}"
print("ALL API ASSERTIONS PASSED!")
