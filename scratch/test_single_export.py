import time
import os
import sys

# Add project root to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from screener_client import fetch_company_data
from valuation_engine import calculate_valuation
from excel_exporter import export_valuation_model, validate_generated_workbook

def test_company(ticker, name, is_fin=False):
    t0 = time.time()
    print(f"\n=======================================================")
    print(f"Testing {ticker} ({name}) - Financial: {is_fin}")
    print(f"=======================================================")
    
    # 1. Data fetch
    print("1. Fetching company data...")
    sd = fetch_company_data(ticker)
    if not sd or not sd.get('company_name'):
        sd['company_name'] = name
        sd['ticker'] = ticker
    print(f"   Fetched {sd.get('company_name')} in {time.time() - t0:.2f}s")
    
    # 2. Valuation
    t1 = time.time()
    print("2. Calculating valuation...")
    val = calculate_valuation(sd)
    print(f"   Valuation calculated in {time.time() - t1:.2f}s")
    
    # 3. Export
    t2 = time.time()
    print("3. Exporting valuation model...")
    dest = export_valuation_model(sd, val)
    print(f"   Exported to {dest} in {time.time() - t2:.2f}s")
    
    # 4. Validation
    t3 = time.time()
    print("4. Validating workbook...")
    validate_generated_workbook(dest, ticker, name, is_financial=is_fin)
    print(f"   Validation passed in {time.time() - t3:.2f}s")
    print(f"TOTAL TIME for {ticker}: {time.time() - t0:.2f}s\n")

if __name__ == '__main__':
    ticker = sys.argv[1] if len(sys.argv) > 1 else 'ADANIPOWER'
    name = sys.argv[2] if len(sys.argv) > 2 else 'Adani Power Limited'
    is_fin = sys.argv[3].lower() == 'true' if len(sys.argv) > 3 else False
    test_company(ticker, name, is_fin)
