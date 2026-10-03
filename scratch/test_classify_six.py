import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client
from universal_valuation.company_classifier import CompanyClassificationEngine

tickers = ['ICICIBANK', 'SBIN', 'ONGC', 'SUNPHARMA', 'TATASTEEL', 'VBL']
for ticker in tickers:
    data = screener_client.fetch_company_data(ticker)
    res = CompanyClassificationEngine.classify(data)
    print(f"{ticker:12} -> Type: {res['canonical_company_type']:<26} | Framework: {res['canonical_valuation_framework']}")
