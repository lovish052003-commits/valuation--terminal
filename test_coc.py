from screener_client import fetch_company_data
from universal_valuation import CompanyClassificationEngine, FinancialNormalizationEngine, CostOfCapitalEngine

data = fetch_company_data('ADANIPOWER')
clf = CompanyClassificationEngine.classify(data)
norm = FinancialNormalizationEngine.normalize(data, clf)
coc = CostOfCapitalEngine.calculate(data, clf, norm)
print("CostOfCapitalEngine results:")
for k, v in coc.items():
    if k != 'metadata':
        print(f"  {k}: {v}")
