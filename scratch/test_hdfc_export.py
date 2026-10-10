import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client
import valuation_engine
import excel_exporter

print("1. Fetching HDFCBANK screener data...")
data = screener_client.fetch_company_data("HDFCBANK")
print(f"Company: {data.get('company_name')}, Type: {data.get('company_type')}")

print("2. Calculating valuation...")
val = valuation_engine.calculate_valuation(data)
print(f"Intrinsic Value: {val.get('intrinsic_value_per_share')}")

print("3. Testing export_valuation_model...")
import time
t0 = time.time()
try:
    path = excel_exporter.export_valuation_model(data, val, "")
    print(f"Export succeeded in {time.time()-t0:.2f}s -> {path}")
except Exception as e:
    import traceback
    print(f"Export failed in {time.time()-t0:.2f}s: {e}")
    traceback.print_exc()
