import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client
import valuation_engine

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

print("Fetching Screener data for 'OIL'...")
data = screener_client.fetch_company_data('OIL')
print(f"Company: {data.get('company_name')} ({data.get('ticker')})")
print(f"Sector: {data.get('sector')}, Industry: {data.get('industry')}")
print(f"Current Price: {data.get('current_price')}")
print(f"CFO: {data.get('cfo_cr')} or {data.get('cash_flow')}")
print(f"EBIT: {data.get('ebit')} | EBITDA: {data.get('ebitda')}")
print(f"Cash: {data.get('cash')} | Debt: {data.get('debt')}")
print(f"Shares: {data.get('shares_outstanding_cr')}")

val = valuation_engine.calculate_valuation(data)
print("\nValuation Engine Result for OIL:")
print(f"Intrinsic Value: {val.get('intrinsic_value')}")
print(f"CMP: {val.get('current_price')}")
print(f"WACC: {val.get('wacc')}")
fp = val.get('four_pillars', {})
p1 = fp.get('pillar1_fcf', {})
print(f"Pillar 1: CFO={p1.get('cfo')}, PAT={p1.get('net_profit')}, EBIT Margin={p1.get('ebit_margin')}%")
