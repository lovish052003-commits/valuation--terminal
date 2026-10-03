import openpyxl

# Load evaluated values from Excel
wb = openpyxl.load_workbook('exports/ADANIENT_Valuation_Model.xlsx', data_only=True)
ws_dcf = wb['DCF']
ws_data = wb['Data Sheet']

excel_vals = {
    'ebit_base': ws_dcf['H8'].value,
    'tax_rate': ws_dcf['H9'].value,
    'growth_rate': ws_dcf['D18'].value,
    'terminal_growth': ws_dcf['D19'].value,
    'wacc': ws_dcf['D20'].value,
    'terminal_reinvest': ws_dcf['D21'].value,
    'base_reinvest': ws_dcf['I11'].value,
    'pv_fcff_sum': ws_dcf['D33'].value,
    'terminal_value': ws_dcf['D29'].value,
    'pv_terminal_value': ws_dcf['D34'].value,
    'enterprise_value': ws_dcf['D35'].value,
    'cash': ws_dcf['D37'].value,
    'debt': ws_dcf['D38'].value,
    'equity_val': ws_dcf['D39'].value,
    'shares_cr': ws_dcf['D40'].value,
    'intrinsic_val': ws_dcf['D42'].value,
    'cmp': ws_dcf['D44'].value
}

print("=== EXCEL ACTUAL VALUES ===")
for k, v in excel_vals.items():
    print(f"{k}: {v}")

# Now compute using pure Python with the harmonized logic:
base_ebit = float(excel_vals['ebit_base'])
tax_rate = float(excel_vals['tax_rate'])
growth_rate = float(excel_vals['growth_rate'])
terminal_g = float(excel_vals['terminal_growth'])
wacc = float(excel_vals['wacc'])
terminal_reinvest = terminal_g / wacc
base_reinvest = float(excel_vals['base_reinvest'])

dcf_schedule = []
current_ebit = base_ebit
pv_fcff_sum = 0.0

for t in range(1, 6):
    proj_ebit = current_ebit * (1.0 + growth_rate)
    proj_nopat = proj_ebit * (1.0 - tax_rate)
    proj_rr = base_reinvest + (terminal_reinvest - base_reinvest) * ((t - 1) / 4.0)
    eff_rr = min(0.85, proj_rr)
    proj_fcff = proj_nopat * (1.0 - eff_rr)
    mid_year = t - 0.5
    df = 1.0 / ((1.0 + wacc) ** mid_year)
    pv_fcff = proj_fcff * df
    pv_fcff_sum += pv_fcff
    dcf_schedule.append({
        't': t, 'ebit': proj_ebit, 'nopat': proj_nopat,
        'rr': proj_rr, 'fcff': proj_fcff, 'df': df, 'pv_fcff': pv_fcff
    })
    current_ebit = proj_ebit

fcff_5 = dcf_schedule[-1]['fcff']
fcff_term = fcff_5 * (1.0 + terminal_g)
tv = fcff_term / (wacc - terminal_g)
df_term = 1.0 / ((1.0 + wacc) ** 4.5)
pv_tv = tv * df_term

ev = pv_fcff_sum + pv_tv
cash = float(excel_vals['cash'])
debt = float(excel_vals['debt'])
eq_val = ev + cash - debt
shares = float(excel_vals['shares_cr'])
intrinsic_p = eq_val / shares

print("\n=== PYTHON HARMONIZED VALUES ===")
print("PV FCFF Sum:", round(pv_fcff_sum, 2), "vs Excel:", round(excel_vals['pv_fcff_sum'], 2))
print("Terminal Value:", round(tv, 2), "vs Excel:", round(excel_vals['terminal_value'], 2))
print("PV Terminal Value:", round(pv_tv, 2), "vs Excel:", round(excel_vals['pv_terminal_value'], 2))
print("Enterprise Value:", round(ev, 2), "vs Excel:", round(excel_vals['enterprise_value'], 2))
print("Equity Value:", round(eq_val, 2), "vs Excel:", round(excel_vals['equity_val'], 2))
print("Intrinsic Value:", round(intrinsic_p, 2), "vs Excel:", round(excel_vals['intrinsic_val'], 2))

diff_pct = abs(intrinsic_p - excel_vals['intrinsic_val']) / excel_vals['intrinsic_val'] * 100
print(f"\nDiscrepancy: {diff_pct:.4f}%")
