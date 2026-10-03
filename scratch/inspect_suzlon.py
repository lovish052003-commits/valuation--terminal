import sys, os
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__) + '/..'))
import screener_client, excel_exporter

d = screener_client.fetch_company_data('Suzlon')
print('Company Name:', d.get('company_name'))
print('Ticker:', d.get('ticker'))
print('Sector Key in excel_exporter:', excel_exporter.get_sector_key(d))

pdf = d.get('peers_df')
if pdf is not None and not pdf.empty:
    print('Peers DF columns:', list(pdf.columns))
    for _, r in pdf.iterrows():
        pname = r.get('Company') or r.get('Name')
        print('Peer:', pname, '| Sales:', r.get('Sales Qtr  Rs.Cr.') or r.get('Sales Qtr Rs.Cr.'), '| NP:', r.get('NP Qtr  Rs.Cr.') or r.get('NP Qtr Rs.Cr.'))
else:
    print('Peers DF is empty!')

eff_peers = excel_exporter.get_effective_peers(d)
print('\nEffective Peers in excel_exporter:')
for p in eff_peers:
    print(p['name'], '| CMP:', p['cmp'], '| Mcap:', p['mcap'], '| Sales:', p['sales'], '| EBITDA:', p['ebitda'])
