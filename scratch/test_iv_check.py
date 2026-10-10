import os, sys
sys.path.insert(0, os.path.abspath('.'))
from screener_client import load_offline_company_data
import valuation_engine

d = load_offline_company_data('NESTLEIND')
v = valuation_engine.calculate_valuation(d)
print("NESTLEIND IV:", v['intrinsic_value_per_share'])
print("EV:", v['enterprise_value'])
print("Cash:", v['cash_estimate'])
print("Debt:", v['total_debt'])
print("WACC:", v['wacc'])
