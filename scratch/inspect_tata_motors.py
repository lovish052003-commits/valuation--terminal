import sys, os
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__) + '/..'))
import screener_client

try:
    d = screener_client.fetch_company_data('Tata Motors')
    print('Company Name:', d.get('company_name'))
    print('Ticker:', d.get('ticker'))
    print('Shares (Cr):', d.get('shares_in_cr'))
    print('Face Value:', d.get('face_value'))
    print('Current Price:', d.get('current_price'))
    print('Market Cap:', d.get('market_cap_cr'))
    
    pdf = d.get('peers_df')
    if pdf is not None and not pdf.empty:
        print('Peers Columns:', list(pdf.columns))
        for _, r in pdf.iterrows():
            print('Peer:', r.get('Name') or r.get('Company'), '| CMP:', r.get('CMP Rs.') or r.get('CMP  Rs.'), '| Mcap:', r.get('Mar Cap Rs.Cr.') or r.get('Mar Cap  Rs.Cr.'))
    else:
        print('Peers DF is empty!')
        
    oa = d.get('schedules', {}).get('Other Assets', {})
    print('Other Assets keys:', list(oa.keys()))
    print('Cash Equivalents:', oa.get('Cash Equivalents'))
    
    bs = d.get('tables', {}).get('balance-sheet')
    if bs is not None and not bs.empty:
        print('BS Metrics:', bs['Metric'].tolist())
        for _, r in bs.iterrows():
            if any(k in r['Metric'].lower() for k in ['equity', 'share', 'borrow', 'asset']):
                print(r['Metric'], '-->', r.iloc[-1])
except Exception as e:
    import traceback
    traceback.print_exc()
