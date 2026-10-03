import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from screener_client import fetch_company_data
from valuation_engine import calculate_valuation, compute_fundamental_reinvestment_engine

def run_archetypes():
    archetypes = [
        ('TATASTEEL', 'Tata Steel Ltd', 'Cyclical Industrial'),
        ('NESTLEIND', 'Nestle India Ltd', 'Mature Consumer Company'),
        ('TRENT', 'Trent Ltd', 'High-Growth Company'),
        ('SUNPHARMA', 'Sun Pharmaceutical Industries Ltd', 'Pharma Company'),
        ('INFY', 'Infosys Ltd', 'IT Services Company'),
        ('SUZLON', 'Suzlon Energy Ltd', 'Negative/Volatile ROIC Company'),
        ('SBIN', 'State Bank of India', 'Financial Institution')
    ]
    
    print("\n================ UNIVERSAL REINVESTMENT ENGINE MULTI-ARCHETYPE TEST ================\n")
    
    for ticker, name, archetype in archetypes:
        try:
            data = fetch_company_data(ticker)
            val = calculate_valuation(data)
            res = compute_fundamental_reinvestment_engine(data)
            
            dcf_val = val.get('target_price') or val.get('dcf_target_price', 0.0)
            
            print(f"--------------------------------------------------------------------------------")
            print(f"Archetype: {archetype}")
            print(f"Target Company: {name} ({ticker})")
            print(f"Historical Median Reinvestment: {res['historical_median_reinvestment_rate']*100:.2f}%" if res['historical_median_reinvestment_rate'] is not None else "Historical Median Reinvestment: N/A")
            print(f"Normalized ROIC: {res['normalized_roic']*100:.2f}%")
            print(f"Expected Growth: {res['expected_growth_rate']*100:.2f}% (Source: {res['growth_source']})")
            print(f"Fundamental Reinvestment: {res['fundamental_reinvestment_rate']*100:.2f}%")
            print(f"Terminal ROIC: {res['terminal_roic']*100:.2f}%")
            print(f"Terminal Reinvestment: {res['terminal_reinvestment_rate']*100:.2f}%")
            print(f"Year 1 DCF Reinvestment: {res['forecast_reinvest_rates'][0]:.2f}%")
            print(f"Year 5 DCF Reinvestment: {res['forecast_reinvest_rates'][4]:.2f}%")
            print(f"DCF Value/Share: Rs. {dcf_val:.2f}")
            print(f"Growth-ROIC Check: {res['growth_roic_consistency']}")
            print(f"Reinvestment Confidence: {res['reinvestment_confidence']}")
            print(f"Warnings: {res['reinvestment_warnings']}")
            
            # Universal assertions
            assert res['forecast_reinvest_rates'][0] == round(res['fundamental_reinvestment_rate'] * 100, 2), "Year 1 Reinvestment != Fundamental Reinvestment!"
            assert res['forecast_reinvest_rates'][4] == round(res['terminal_reinvestment_rate'] * 100, 2), "Year 5 Reinvestment != Terminal Reinvestment!"
            print(f"VERIFIED: Year 1 ({res['forecast_reinvest_rates'][0]:.2f}%) == Fundamental ({res['fundamental_reinvestment_rate']*100:.2f}%)")
            
        except Exception as e:
            print(f"Error on {ticker}: {e}")
            import traceback
            traceback.print_exc()

if __name__ == '__main__':
    run_archetypes()
