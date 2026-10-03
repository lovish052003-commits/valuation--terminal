import os
import sys
import json
import pandas as pd

sys.path.insert(0, os.path.abspath('.'))
from valuation_engine import calculate_valuation

tickers = ['TATASTEEL', 'NESTLEIND', 'INFY', 'UNITEDTEA', 'OISL', 'SBIN']
for t in tickers:
    p = os.path.join('exports', '.screener_cache', f'{t}_data.json')
    if not os.path.exists(p):
        continue
    with open(p, 'r', encoding='utf-8') as f:
        data = json.load(f)
    if 'tables' in data:
        for tb_k, tb_v in data['tables'].items():
            if isinstance(tb_v, dict) and 'columns' in tb_v and 'data' in tb_v:
                data['tables'][tb_k] = pd.DataFrame(tb_v['data'], columns=tb_v['columns'])
    
    res = calculate_valuation(data)
    print(f"\n==========================================")
    print(f"Ticker: {res['ticker']} ({res['company_name']}) | Type: {res['company_type']}")
    print(f"CMP: Rs. {res['current_price']} | Intrinsic Value: Rs. {res['intrinsic_value_per_share']} | Verdict: {res['verdict']}")
    print(f"Normalized ROIC: {res['normalized_roic']}%")
    print(f"Expected Growth: {res['expected_growth_rate']}% (Source: {res['growth_source']})")
    print(f"Fundamental Reinvestment Rate: {res['fundamental_reinvestment_rate']}%")
    print(f"Historical Median Reinvestment Rate: {res['historical_median_reinvestment_rate']}%")
    print(f"Terminal ROIC: {res['terminal_roic']}% | Terminal Reinvest: {res['terminal_reinvestment_rate']}%")
    print(f"5-Year DCF Forecast Reinvestment: {[x['reinvestment_rate'] for x in res['dcf_table']]}")
    print(f"Methodology: {res['reinvestment_methodology']}")
    print(f"Confidence: {res['reinvestment_confidence']} | Consistency: {res['growth_roic_consistency']}")
    print(f"ROIC x Growth Sensitivity Grid Rows: {len(res['sensitivity_roic_growth']['rows'])}")
    if res['reinvestment_warnings']:
        print(f"Warnings: {res['reinvestment_warnings']}")
print("\nALL RUNS COMPLETED SUCCESSFULLY!")
