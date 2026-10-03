import json
from screener_client import fetch_company_data
from valuation_engine import calculate_valuation
from excel_exporter import build_wacc_peer_companies

data = fetch_company_data('ADANIPOWER')
val_res = calculate_valuation(data)
print("valuation_engine wacc:", val_res.get('wacc'))
print("valuation_engine cost_of_equity:", val_res.get('cost_of_equity'))

comps = build_wacc_peer_companies(data)
print("\nComps in WACC sheet:")
for c in comps[:5]:
    print(c['name'], "beta:", c['beta'], "debt:", c['debt'], "mcap:", c['mcap'])
