import openpyxl

wb = openpyxl.load_workbook('master_model_template.xlsx')

# 1. Intrinsic Valuation Sheet
if 'Intrinsic Valuation' in wb.sheetnames:
    ws_iv = wb['Intrinsic Valuation']
    cols_iv = ['H', 'I', 'J', 'K', 'L']
    cols_ds = ['G', 'H', 'I', 'J', 'K']

    # Row 6: Dates
    for c_iv, c_ds in zip(cols_iv, cols_ds):
        ws_iv[f'{c_iv}6'] = f"='Data Sheet'!{c_ds}16"

    # Current Assets (Rows 8-13)
    ws_iv['B9'] = "Inventory"
    ws_iv['B10'] = "Trade Receivables"
    ws_iv['B11'] = "Cash & Cash Equivalents"
    ws_iv['B12'] = "Other Current Assets"
    for c_iv, c_ds in zip(cols_iv, cols_ds):
        ws_iv[f'{c_iv}9'] = f"='Data Sheet'!{c_ds}68"
        ws_iv[f'{c_iv}10'] = f"='Data Sheet'!{c_ds}67"
        ws_iv[f'{c_iv}11'] = f"='Data Sheet'!{c_ds}69"
        ws_iv[f'{c_iv}12'] = f"='Data Sheet'!{c_ds}65"

    # Current Liabilities (Rows 15-19)
    ws_iv['B16'] = "Trade Payables"
    ws_iv['B17'] = "Other Current Liabilities"
    ws_iv['B18'] = "Short Term Borrowings"
    for c_iv, c_ds in zip(cols_iv, cols_ds):
        ws_iv[f'{c_iv}16'] = f"='Data Sheet'!{c_ds}60*0.4"
        ws_iv[f'{c_iv}17'] = f"='Data Sheet'!{c_ds}60*0.4"
        ws_iv[f'{c_iv}18'] = f"='Data Sheet'!{c_ds}60*0.2"

    # Non-Current Assets (Rows 23-35)
    ws_iv['B24'] = "Fixed Assets"
    ws_iv['B25'] = "Capital Work in Progress (CWIP)"
    ws_iv['B26'] = "Investments"
    for c_iv, c_ds in zip(cols_iv, cols_ds):
        ws_iv[f'{c_iv}24'] = f"='Data Sheet'!{c_ds}62"
        ws_iv[f'{c_iv}25'] = f"='Data Sheet'!{c_ds}63"
        ws_iv[f'{c_iv}26'] = f"='Data Sheet'!{c_ds}64"
        for r_clear in range(27, 35):
            ws_iv[f'B{r_clear}'] = None
            ws_iv[f'{c_iv}{r_clear}'] = 0.0

    # Net Non-Current Assets (Row 35)
    for c_iv in cols_iv:
        ws_iv[f'{c_iv}35'] = f"=SUM({c_iv}24:{c_iv}34)"

    # EBIT (Row 38)
    for c_iv, c_ds in zip(cols_iv, cols_ds):
        ws_iv[f'{c_iv}38'] = f"='Data Sheet'!{c_ds}32-'Data Sheet'!{c_ds}26"

    # Reinvestment Dates (Row 42)
    for c_iv in cols_iv:
        ws_iv[f'{c_iv}42'] = f"={c_iv}6"

    # Net Capex (Row 44)
    for c_iv, c_ds in zip(cols_iv, cols_ds):
        ws_iv[f'{c_iv}44'] = f"=-'Data Sheet'!{c_ds}83"

    # Growth Dates (Row 57)
    for c_iv in cols_iv:
        ws_iv[f'{c_iv}57'] = f"={c_iv}6"

# 2. Comp_Valuation Sheet
if 'Comp_Valuation' in wb.sheetnames:
    ws_cv = wb['Comp_Valuation']
    # Clear peer rows (Rows 12-21)
    for r in range(12, 22):
        for col_l in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']:
            ws_cv[f'{col_l}{r}'].value = None

    # Target company rows (Rows 32-39)
    ws_cv['I32'] = '=IF(OR(\'Data Sheet\'!K17<=0, I25="N/A", ISBLANK(I25)), "N/A", \'Data Sheet\'!K17*I25)'
    ws_cv['J32'] = '=IF(OR(\'Data Sheet\'!K32<=0, J25="N/A", ISBLANK(J25)), "N/A", \'Data Sheet\'!K32*J25)'
    ws_cv['O32'] = '=IF(OR(\'Data Sheet\'!K17<=0, O25="N/A", ISBLANK(O25)), "N/A", \'Data Sheet\'!K17*O25)'
    ws_cv['P32'] = '=IF(OR(\'Data Sheet\'!K32<=0, P25="N/A", ISBLANK(P25)), "N/A", \'Data Sheet\'!K32*P25)'
    ws_cv['Q32'] = '=IF(OR(\'Data Sheet\'!K30<=0, Q25="N/A", ISBLANK(Q25)), "N/A", \'Data Sheet\'!K30*Q25)'

    ws_cv['I33'] = "='Data Sheet'!K59-'Data Sheet'!K69"
    ws_cv['J33'] = "='Data Sheet'!K59-'Data Sheet'!K69"
    ws_cv['O33'] = "='Data Sheet'!K59-'Data Sheet'!K69"
    ws_cv['P33'] = "='Data Sheet'!K59-'Data Sheet'!K69"
    ws_cv['Q33'] = 0

    ws_cv['I35'] = "='Data Sheet'!B6"
    ws_cv['J35'] = "='Data Sheet'!B6"
    ws_cv['O35'] = "='Data Sheet'!B6"
    ws_cv['P35'] = "='Data Sheet'!B6"
    ws_cv['Q35'] = "='Data Sheet'!B6"

    ws_cv['O38'] = "='Data Sheet'!B8"
    ws_cv['P38'] = "='Data Sheet'!B8"
    ws_cv['Q38'] = "='Data Sheet'!B8"

    ws_cv['B39'] = "Price vs Implied Variance"
    ws_cv['I39'] = '=IF(OR(I37="N/A", ISBLANK(I37), \'Data Sheet\'!B8<=0), "N/A", IFERROR((I37-\'Data Sheet\'!B8)/\'Data Sheet\'!B8, "N/A"))'
    ws_cv['J39'] = '=IF(OR(J37="N/A", ISBLANK(J37), \'Data Sheet\'!B8<=0), "N/A", IFERROR((J37-\'Data Sheet\'!B8)/\'Data Sheet\'!B8, "N/A"))'
    ws_cv['O39'] = '=IF(OR(O37="N/A", ISBLANK(O37), \'Data Sheet\'!B8<=0), "N/A", IFERROR((O37-\'Data Sheet\'!B8)/\'Data Sheet\'!B8, "N/A"))'
    ws_cv['P39'] = '=IF(OR(P37="N/A", ISBLANK(P37), \'Data Sheet\'!B8<=0), "N/A", IFERROR((P37-\'Data Sheet\'!B8)/\'Data Sheet\'!B8, "N/A"))'
    ws_cv['Q39'] = '=IF(OR(Q37="N/A", ISBLANK(Q37), \'Data Sheet\'!B8<=0), "N/A", IFERROR((Q37-\'Data Sheet\'!B8)/\'Data Sheet\'!B8, "N/A"))'

wb.save('master_model_template.xlsx')
print("[SUCCESS] master_model_template.xlsx cleaned of all Raw FS references!")
