import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client
import valuation_engine

sd = screener_client.fetch_company_data('ITC')
val = valuation_engine.calculate_valuation(sd)
print("CMP:", val.get('current_price'))
print("Intrinsic Value:", val.get('intrinsic_value_per_share'))
print("WACC:", val.get('wacc'))
print("Growth Rate:", val.get('growth_rate'))
print("Enterprise Value:", val.get('enterprise_value'))
print("Equity Value:", val.get('equity_value'))
print("Margin of Safety:", val.get('margin_of_safety_pct'))
print("Verdict:", val.get('verdict'))
