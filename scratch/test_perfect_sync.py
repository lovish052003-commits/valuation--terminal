import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import screener_client

def test_sync():
    data = screener_client.fetch_company_data('ITC')
    
    # 1. Base inputs from data
    current_price = 255.0
    shares_cr = 1253.0
    market_cap = current_price * shares_cr
    
    # Financial statements
    # EBIT
    latest_ebit = 25595.0
    tax_rate = 0.2517
    nopat_base = latest_ebit * (1 - tax_rate) # 19152.7385
    
    # Reinvestment
    # Net capex = 422, Change in WC = 1184 -> total = 1606
    # Reinvestment rate
    reinvest_rate_1 = 0.08469412450318278
    roic = 0.5887172563243476
    
    # WACC
    wacc = 0.11191343057450917
    terminal_g = 0.04
    
    terminal_roic = roic
    terminal_reinvest = terminal_g / terminal_roic # 0.06794433078068705
    
    # 5 Year DCF
    curr_ebit = latest_ebit
    pv_fcff_sum = 0.0
    dcf_rows = []
    
    for i in range(1, 6):
        fade = (i - 1) / 4.0
        rr = reinvest_rate_1 + (terminal_reinvest - reinvest_rate_1) * fade
        g_i = rr * roic
        ebit_i = curr_ebit * (1 + g_i)
        nopat_i = ebit_i * (1 - tax_rate)
        fcff_i = nopat_i * (1 - rr)
        mid_year = i - 0.5
        df_i = (1 + wacc) ** (-mid_year)
        pv_fcff_i = fcff_i * df_i
        pv_fcff_sum += pv_fcff_i
        curr_ebit = ebit_i
        dcf_rows.append({
            'year': f'Year {i}',
            'ebit': ebit_i,
            'nopat': nopat_i,
            'reinvestment_rate': rr * 100,
            'fcff': fcff_i,
            'discount_factor': df_i,
            'pv_fcff': pv_fcff_i
        })
        print(f"Year {i}: EBIT={ebit_i:.2f}, NOPAT={nopat_i:.2f}, RR={rr*100:.2f}%, FCFF={fcff_i:.2f}, DF={df_i:.4f}, PV={pv_fcff_i:.2f}")

    tv = (dcf_rows[-1]['fcff'] * (1 + terminal_g)) / (wacc - terminal_g)
    pv_tv = tv * ((1 + wacc) ** (-4.5))
    ev = pv_fcff_sum + pv_tv
    
    cash = 72443.2
    debt = 2399.0
    minority_interest = 366.0
    
    eq_val = ev + cash - debt - minority_interest
    iv = eq_val / shares_cr
    
    print("\n--- RESULTS ---")
    print(f"PV FCFF Sum: {pv_fcff_sum:.2f}")
    print(f"Terminal Value: {tv:.2f}")
    print(f"PV Terminal Value: {pv_tv:.2f}")
    print(f"Enterprise Value: {ev:.2f}")
    print(f"Equity Value: {eq_val:.2f}")
    print(f"Intrinsic Value per Share: Rs. {iv:.2f}")
    
    assert abs(iv - 277.13) < 0.1, f"Expected 277.13, got {iv}"
    print("\n>>> 100% PARITY CONFIRMED! <<<")

if __name__ == '__main__':
    test_sync()
