import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import screener_client
import valuation_engine

data = screener_client.fetch_company_data('SBIN')
val = valuation_engine.calculate_valuation(data)

u = val.get('universal_engine', {})
cls = u.get('classification', {})
dcf = u.get('dcf', {})
coc = u.get('cost_of_capital', {})

print(f"Company: {data.get('company_name')} ({data.get('ticker')})")
print(f"Classification: is_bank={cls.get('is_bank')}, is_financial={cls.get('is_financial')}, family={cls.get('valuation_family')}")
print(f"Cost of Capital: Ke={coc.get('cost_of_equity')}%, WACC={coc.get('wacc')}%, discount_rate={coc.get('discount_rate')}%")
print(f"Valuation Model: {dcf.get('valuation_model')}")
print(f"Intrinsic Value Per Share: Rs. {dcf.get('intrinsic_value_per_share')}")
print(f"Current Book Value: Rs. {dcf.get('current_book_value_cr')} Cr")
print(f"Equity Value: Rs. {dcf.get('equity_value_cr')} Cr")
print(f"Shares: {dcf.get('shares_outstanding_cr')} Cr")
print(f"Sustainable ROE: {dcf.get('sustainable_roe_pct')}%")
print(f"Excess Return Schedule:")
for row in dcf.get('forecast_schedule', []):
    print(" ", row)
