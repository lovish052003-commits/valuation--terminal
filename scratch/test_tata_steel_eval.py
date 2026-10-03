import sys
sys.path.insert(0, '.')
import screener_client
import valuation_engine

print("Fetching Tata Steel...")
sd = screener_client.fetch_company_data("Tata Steel")
print("Company Name:", sd.get('company_name'))
print("Ticker:", sd.get('ticker'))
print("Current Price:", sd.get('current_price'))
print("Market Cap:", sd.get('market_cap_cr'))
print("Tables:", list(sd.get('tables', {}).keys()))
pl = sd.get('tables', {}).get('profit-loss')
if pl is not None:
    print("PL metrics sample:", pl['Metric'].tolist()[:8])
    print("PL latest cols:", pl.columns.tolist()[-5:])
    print("PL sales row:", pl[pl['Metric'].str.contains('Sales|Revenue', case=False, na=False)].to_dict(orient='records'))

print("\nPeers sample:")
pdf = sd.get('peers_df')
if pdf is not None and not pdf.empty:
    print(pdf[['Company', 'CMP  Rs.', 'Mar Cap  Rs.Cr.']].head(5))

val = valuation_engine.calculate_valuation(sd)
print("\nValuation result:")
print("Intrinsic Value:", val.get('intrinsic_value_per_share'))
print("Current Price:", val.get('current_price'))
print("WACC:", val.get('wacc'))
print("Growth Rate:", val.get('growth_rate'))
print("Enterprise Value:", val.get('enterprise_value'))
print("Equity Value:", val.get('equity_value'))
print("Verdict:", val.get('verdict'))
