import os, sys
sys.path.insert(0, os.path.abspath('.'))
from screener_client import load_offline_company_data
import valuation_engine

# Let's test the math changes directly
d_nestle = load_offline_company_data('NESTLEIND')

# Let's inspect what curr_roic was in valuation_engine for NESTLEIND
# Let's calculate:
latest_ebit = 4799.22 / (1 + 0.052) # 4562.0
eff_tax = 0.2517
# If cash is 2577.90
# Terminal roic:
# For Nestlé, if terminal_roic is curr_roic:
# Let's see:
print("Testing math...")
