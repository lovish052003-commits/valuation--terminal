import urllib.request
import json
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

for ticker in ['SUNPHARMA', 'SBIN']:
    print(f"\n--- Testing live API for {ticker} ---")
    req = urllib.request.Request(
        'http://127.0.0.1:5000/api/analyze',
        data=json.dumps({'company': ticker, 'skipAi': True}).encode('utf-8'),
        headers={'Content-Type': 'application/json'}
    )
    resp = urllib.request.urlopen(req)
    res = json.loads(resp.read().decode('utf-8'))
    val = res.get('valuation', {})
    print('Company:', res['company']['name'])
    print('Ticker:', res['company']['ticker'])
    print('Intrinsic Value:', val.get('intrinsic_value_per_share'))
    print('Valuation Gap:', val.get('valuation_gap', {}).get('gap_label'))
    print('Classification Family:', val.get('company_classification', {}).get('valuation_family'))
    print('Primary Valuation Method:', val.get('valuation_methods', {}).get('primary_method'))
    print('Model Quality Status:', val.get('model_quality', {}).get('overall_status'))
    p_count = val.get('model_quality', {}).get('pass_count')
    tot_count = val.get('model_quality', {}).get('total_tests')
    print(f'Quality Tests Passed: {p_count}/{tot_count}')
    print('Confidence Level:', val.get('valuation_confidence', {}).get('confidence_level'))
