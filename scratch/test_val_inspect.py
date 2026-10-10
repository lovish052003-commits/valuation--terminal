import os, sys
sys.path.insert(0, os.path.abspath('.'))
from screener_client import load_offline_company_data
import valuation_engine

# Let's inspect calculate_valuation with aligned parameters
d = load_offline_company_data('NESTLEIND')

# Custom params test
v = valuation_engine.calculate_valuation(d)
print("NESTLEIND current engine IV:", v['intrinsic_value_per_share'])

# Let's trace why curr_roic was calculated
print("invested_capital:", d.get('invested_capital'))
