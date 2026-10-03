import os
import sys
import openpyxl

sys.path.insert(0, os.path.abspath('.'))
import screener_client

stocks = screener_client.load_list_of_stocks()
print(f"Total stocks loaded: {len(stocks)}")
print(f"Total indexed by NSE: {len(screener_client._STOCKS_BY_NSE)}")
print(f"Total indexed by BSE: {len(screener_client._STOCKS_BY_BSE)}")
print(f"Total indexed by Norm Name: {len(screener_client._STOCKS_BY_NAME)}")

# Sample across different sectors and company types
test_queries = [
    "ITC", "Nestle India", "Jubilant FoodWorks", "Force Motors", "Tata Motors",
    "HDFC Bank", "RBL Bank", "Bajaj Finance", "TCS", "Infosys", "Tata Steel",
    "JSW Steel", "Sun Pharma", "Reliance Industries", "State Bank of India",
    "Titan Company", "Avenue Supermarts", "Varun Beverages", "Larsen & Toubro",
    "500033", # BSE code for Force Motors
    "500875"  # BSE code for ITC
]

print("\n--- Testing Name/Ticker Resolution for diverse companies ---")
for q in test_queries:
    res = screener_client.resolve_company(q)
    print(f"Query: {q:25} -> Ticker: {res.get('ticker'):12} | Name: {res.get('company_name'):35} | Type: {res.get('company_type')}")
    assert res.get('resolved') is True, f"Failed to resolve {q}"
    assert res.get('ticker'), f"No ticker for {q}"

print("\n>>> ALL SAMPLE RESOLUTIONS PASSED SUCCESSFULLY!")
