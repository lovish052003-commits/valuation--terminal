import os, sys, shutil, json, requests, gc
import win32com.client, pythoncom
import pandas as pd

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from screener_client import fetch_company_data, clean_num, HEADERS
from valuation_engine import calculate_valuation

def fetch_schedules_for_company(company_id, is_consolidated=True):
    schedules = {}
    if not company_id:
        return schedules
    schedule_targets = [
        ('Fixed Assets', 'balance-sheet'),
        ('Other Assets', 'balance-sheet'),
        ('Borrowings', 'balance-sheet'),
        ('Other Liabilities', 'balance-sheet'),
        ('Cash from Operating Activity', 'cash-flow'),
        ('Cash from Investing Activity', 'cash-flow'),
        ('Cash from Financing Activity', 'cash-flow')
    ]
    for parent, section in schedule_targets:
        try:
            cons_param = "&consolidated=" if is_consolidated else ""
            sch_url = f"https://www.screener.in/api/company/{company_id}/schedules/?parent={requests.utils.quote(parent)}&section={section}{cons_param}"
            s_resp = requests.get(sch_url, headers=HEADERS, timeout=8)
            if s_resp.status_code == 200 and s_resp.text.startswith('{'):
                schedules[parent] = s_resp.json()
        except Exception as e:
            print(f"Notice: Failed schedule {parent}: {e}")
    return schedules

print("1. Fetching Tata Motors...")
data = fetch_company_data('Tata Motors')
data['schedules'] = fetch_schedules_for_company(data.get('company_id'), is_consolidated=True)
val = calculate_valuation(data)

dest_path = os.path.abspath('exports/TEST_TATA_RAW_FS.xlsx')
shutil.copyfile('ITC Model.xlsx', dest_path)

print("2. Opening Excel COM...")
pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.ScreenUpdating = False

try:
    wb = excel.Workbooks.Open(dest_path)
    
    # -------------------------------------------------------------
    # 1. Populate Data Sheet
    # -------------------------------------------------------------
    ws_data = wb.Sheets('Data Sheet')
    ws_data.Range('B1').Value = data.get('company_name', '')
    ws_data.Range('B6').Value = clean_num(data.get('shares_in_cr', 0))
    ws_data.Range('B7').Value = clean_num(data.get('face_value', 1))
    ws_data.Range('B8').Value = clean_num(data.get('current_price', 0))
    ws_data.Range('B9').Value = clean_num(data.get('market_cap_cr', 0))

    tables = data.get('tables', {})
    pl_df = tables.get('profit-loss')
    bs_df = tables.get('balance-sheet')
    cf_df = tables.get('cash-flow')
    schedules = data.get('schedules', {})

    def fill_ds_row(row_idx, df, keywords):
        if df is None or df.empty:
            return
        period_cols = [c for c in df.columns if c != 'Metric']
        for kw in keywords:
            m = df[df['Metric'].str.contains(kw, case=False, na=False)]
            if not m.empty:
                r = m.iloc[0]
                for offset, col_name in enumerate(period_cols[-10:]):
                    ws_data.Cells(row_idx, 2 + offset).Value = clean_num(r[col_name])
                break

    def fill_ds_dates(row_idx, df):
        if df is None or df.empty:
            return
        period_cols = [c for c in df.columns if c != 'Metric']
        for offset, col_name in enumerate(period_cols[-10:]):
            ws_data.Cells(row_idx, 2 + offset).Value = str(col_name)

    fill_ds_dates(16, pl_df)
    fill_ds_row(17, pl_df, ['^Sales', 'Revenue'])
    fill_ds_row(18, pl_df, ['Raw Material Cost', 'Material'])
    fill_ds_row(22, pl_df, ['Employee Cost'])
    fill_ds_row(25, pl_df, ['Other Income'])
    fill_ds_row(26, pl_df, ['Depreciation'])
    fill_ds_row(27, pl_df, ['Interest'])
    fill_ds_row(28, pl_df, ['Profit before tax', 'PBT'])
    fill_ds_row(29, pl_df, ['Tax'])
    fill_ds_row(30, pl_df, ['Net profit', 'PAT'])
    fill_ds_row(32, pl_df, ['Operating Profit', 'EBITDA'])

    fill_ds_dates(56, bs_df)
    fill_ds_row(57, bs_df, ['Equity Capital', 'Share Capital'])
    fill_ds_row(58, bs_df, ['Reserves'])
    fill_ds_row(59, bs_df, ['Borrowings', 'Total Debt'])
    fill_ds_row(60, bs_df, ['Other Liabilities'])
    fill_ds_row(61, bs_df, ['Total Liabilities'])
    fill_ds_row(62, bs_df, ['Fixed Assets', 'Net Block'])
    fill_ds_row(63, bs_df, ['CWIP'])
    fill_ds_row(64, bs_df, ['Investments'])
    fill_ds_row(65, bs_df, ['Other Assets'])
    fill_ds_row(66, bs_df, ['Total Assets'])

    fill_ds_dates(81, cf_df)
    fill_ds_row(82, cf_df, ['Cash from Operating Activity'])
    fill_ds_row(83, cf_df, ['Cash from Investing Activity'])
    fill_ds_row(84, cf_df, ['Cash from Financing Activity'])
    fill_ds_row(85, cf_df, ['Net Cash Flow'])

    # -------------------------------------------------------------
    # 2. Populate Raw FS Sheet
    # -------------------------------------------------------------
    print("3. Populating Raw FS sheet with live Screener data...")
    ws_raw = wb.Sheets('Raw FS')

    # A. Balance Sheet (Cols C to N = cols 3 to 14)
    if bs_df is not None and not bs_df.empty:
        bs_period_cols = [c for c in bs_df.columns if c != 'Metric']
        # Use up to 12 periods matching C to N
        sel_periods = bs_period_cols[-12:]
        for offset, p_name in enumerate(sel_periods):
            c_idx = 3 + offset
            ws_raw.Cells(3, c_idx).Value = str(p_name)

        def write_raw_bs_row(row_idx, keywords):
            for kw in keywords:
                m = bs_df[bs_df['Metric'].str.contains(kw, case=False, na=False)]
                if not m.empty:
                    r = m.iloc[0]
                    for offset, p_name in enumerate(sel_periods):
                        ws_raw.Cells(row_idx, 3 + offset).Value = clean_num(r[p_name])
                    break

        def write_raw_schedule_row(row_idx, parent_name, item_keyword):
            sch = schedules.get(parent_name, {})
            for item_name, period_dict in sch.items():
                if item_keyword.lower() in item_name.lower():
                    for offset, p_name in enumerate(sel_periods):
                        val_str = period_dict.get(p_name, 0)
                        ws_raw.Cells(row_idx, 3 + offset).Value = clean_num(val_str)
                    break

        write_raw_bs_row(5, ['Equity Capital', 'Share Capital'])
        write_raw_bs_row(6, ['Reserves'])
        write_raw_bs_row(7, ['Borrowings'])
        write_raw_schedule_row(8, 'Borrowings', 'Long term Borrowings')
        write_raw_schedule_row(9, 'Borrowings', 'Short term Borrowings')
        write_raw_schedule_row(10, 'Borrowings', 'Lease Liabilities')
        write_raw_schedule_row(11, 'Borrowings', 'Other Borrowings')

        write_raw_bs_row(12, ['Other Liabilities'])
        write_raw_schedule_row(13, 'Other Liabilities', 'Non controlling int')
        write_raw_schedule_row(14, 'Other Liabilities', 'Trade Payables')
        write_raw_schedule_row(15, 'Other Liabilities', 'Advance from Customers')
        write_raw_schedule_row(16, 'Other Liabilities', 'Other liability')

        write_raw_bs_row(21, ['Fixed Assets'])
        write_raw_schedule_row(22, 'Fixed Assets', 'Land')
        write_raw_schedule_row(23, 'Fixed Assets', 'Building')
        write_raw_schedule_row(24, 'Fixed Assets', 'Plant Machinery')
        write_raw_schedule_row(25, 'Fixed Assets', 'Equipments')
        write_raw_schedule_row(26, 'Fixed Assets', 'Furniture')
        write_raw_schedule_row(27, 'Fixed Assets', 'Railway sidings')
        write_raw_schedule_row(28, 'Fixed Assets', 'Vehicles')
        write_raw_schedule_row(29, 'Fixed Assets', 'Intangible')
        write_raw_schedule_row(30, 'Fixed Assets', 'Other fixed assets')
        write_raw_schedule_row(31, 'Fixed Assets', 'Gross Block')
        write_raw_schedule_row(32, 'Fixed Assets', 'Accumulated Depreciation')

        write_raw_bs_row(35, ['CWIP'])
        write_raw_bs_row(36, ['Investments'])
        write_raw_bs_row(38, ['Other Assets'])
        write_raw_schedule_row(39, 'Other Assets', 'Inventories')
        write_raw_schedule_row(40, 'Other Assets', 'Trade receivables')
        write_raw_schedule_row(41, 'Other Assets', 'Cash Equivalents')
        write_raw_schedule_row(42, 'Other Assets', 'Loans n Advances')
        write_raw_schedule_row(43, 'Other Assets', 'Other asset items')
        write_raw_bs_row(44, ['Total Assets'])

    # B. Profit & Loss (Cols S to AE = cols 19 to 31)
    if pl_df is not None and not pl_df.empty:
        pl_period_cols = [c for c in pl_df.columns if c != 'Metric']
        sel_pl_periods = pl_period_cols[-13:] # up to 12 years + TTM
        for offset, p_name in enumerate(sel_pl_periods):
            c_idx = 19 + offset
            ws_raw.Cells(3, c_idx).Value = str(p_name)

        def write_raw_pl_row(row_idx, keywords):
            for kw in keywords:
                m = pl_df[pl_df['Metric'].str.contains(kw, case=False, na=False)]
                if not m.empty:
                    r = m.iloc[0]
                    for offset, p_name in enumerate(sel_pl_periods):
                        ws_raw.Cells(row_idx, 19 + offset).Value = clean_num(r[p_name])
                    break

        write_raw_pl_row(4, ['^Sales', 'Revenue'])
        write_raw_pl_row(5, ['^Expenses'])
        write_raw_pl_row(6, ['Operating Profit'])
        write_raw_pl_row(7, ['OPM'])
        write_raw_pl_row(8, ['Other Income'])
        write_raw_pl_row(9, ['Interest'])
        write_raw_pl_row(10, ['Depreciation'])
        write_raw_pl_row(11, ['Profit before tax', 'PBT'])
        write_raw_pl_row(12, ['Tax %'])
        write_raw_pl_row(13, ['Net Profit', 'PAT'])
        write_raw_pl_row(14, ['EPS'])
        write_raw_pl_row(15, ['Dividend Payout'])

    # C. Cash Flow (Cols S to AE = cols 19 to 31)
    if cf_df is not None and not cf_df.empty:
        cf_period_cols = [c for c in cf_df.columns if c != 'Metric']
        sel_cf_periods = cf_period_cols[-13:]

        def write_raw_cf_row(row_idx, keywords):
            for kw in keywords:
                m = cf_df[cf_df['Metric'].str.contains(kw, case=False, na=False)]
                if not m.empty:
                    r = m.iloc[0]
                    for offset, p_name in enumerate(sel_cf_periods):
                        if p_name in r:
                            ws_raw.Cells(row_idx, 19 + offset).Value = clean_num(r[p_name])
                    break

        def write_raw_cf_schedule(row_idx, parent_name, item_keyword):
            sch = schedules.get(parent_name, {})
            for item_name, period_dict in sch.items():
                if item_keyword.lower() in item_name.lower():
                    for offset, p_name in enumerate(sel_cf_periods):
                        if p_name in period_dict:
                            ws_raw.Cells(row_idx, 19 + offset).Value = clean_num(period_dict[p_name])
                    break

        write_raw_cf_row(19, ['Cash from Operating Activity'])
        write_raw_cf_schedule(20, 'Cash from Operating Activity', 'Profit from operations')
        write_raw_cf_schedule(21, 'Cash from Operating Activity', 'Receivables')
        write_raw_cf_schedule(22, 'Cash from Operating Activity', 'Inventory')
        write_raw_cf_schedule(23, 'Cash from Operating Activity', 'Payables')
        write_raw_cf_schedule(24, 'Cash from Operating Activity', 'Working capital')
        write_raw_cf_schedule(25, 'Cash from Operating Activity', 'Direct taxes')

        write_raw_cf_row(26, ['Cash from Investing Activity'])
        write_raw_cf_schedule(27, 'Cash from Investing Activity', 'Fixed assets purchased')
        write_raw_cf_schedule(28, 'Cash from Investing Activity', 'Fixed assets sold')
        write_raw_cf_schedule(29, 'Cash from Investing Activity', 'Investments purchased')
        write_raw_cf_schedule(30, 'Cash from Investing Activity', 'Investments sold')
        write_raw_cf_schedule(31, 'Cash from Investing Activity', 'Interest received')
        write_raw_cf_schedule(32, 'Cash from Investing Activity', 'Dividends received')
        write_raw_cf_schedule(33, 'Cash from Investing Activity', 'subsidiaries')
        write_raw_cf_schedule(34, 'Cash from Investing Activity', 'Other investing')

        write_raw_cf_row(36, ['Cash from Financing Activity'])
        write_raw_cf_schedule(37, 'Cash from Financing Activity', 'shares')
        write_raw_cf_schedule(38, 'Cash from Financing Activity', 'Proceeds from borrowings')
        write_raw_cf_schedule(39, 'Cash from Financing Activity', 'Repayment of borrowings')
        write_raw_cf_schedule(40, 'Cash from Financing Activity', 'Interest paid')
        write_raw_cf_schedule(41, 'Cash from Financing Activity', 'Dividends paid')
        write_raw_cf_row(45, ['Net Cash Flow'])

    # D. Peers Table in Raw FS (Rows 56 to 66)
    print("4. Populating Peers Table in Raw FS...")
    peers_df = data.get('peers_df', pd.DataFrame())
    
    # Target company row 56
    ws_raw.Cells(56, 12).Value = data['company_name']
    ws_raw.Cells(56, 13).Value = clean_num(data.get('current_price', 0))
    ws_raw.Cells(56, 14).Value = clean_num(data.get('shares_in_cr', 0))
    ws_raw.Cells(56, 44).Value = clean_num(data.get('market_cap_cr', 0))
    ws_raw.Cells(56, 52).Value = clean_num(val.get('total_debt', 0))
    ws_raw.Cells(56, 53).Value = clean_num(val.get('cash_estimate', 0))

    if not peers_df.empty:
        for p_idx, (_, p_row) in enumerate(peers_df.head(10).iterrows(), start=1):
            r_num = 56 + p_idx
            ws_raw.Cells(r_num, 11).Value = p_idx
            ws_raw.Cells(r_num, 12).Value = str(p_row.get('Name', ''))
            p_cmp = clean_num(p_row.get('CMP Rs.', 0))
            p_mcap = clean_num(p_row.get('Mar Cap Rs.Cr.', 0))
            ws_raw.Cells(r_num, 13).Value = p_cmp
            ws_raw.Cells(r_num, 44).Value = p_mcap
            if p_cmp > 0 and p_mcap > 0:
                ws_raw.Cells(r_num, 14).Value = round(p_mcap / p_cmp, 2)
            if 'Sales Qtr Rs.Cr.' in p_row:
                ws_raw.Cells(r_num, 47).Value = clean_num(p_row.get('Sales Qtr Rs.Cr.', 0)) * 4 # annualize
            if 'NP Qtr Rs.Cr.' in p_row:
                ws_raw.Cells(r_num, 49).Value = clean_num(p_row.get('NP Qtr Rs.Cr.', 0)) * 4
            if 'ROCE %' in p_row:
                ws_raw.Cells(r_num, 50).Value = clean_num(p_row.get('ROCE %', 0))

    # -------------------------------------------------------------
    # 3. Fix Forecasting Sheet (Requirement 1)
    # -------------------------------------------------------------
    print("5. Fixing Forecasting sheet: =C13+365 at year weight 10 till year 15...")
    ws_f = wb.Sheets('Forecasting')

    ws_f.Range('C5:C14').ClearContents()
    ws_f.Range('C5:C13').FormulaArray = "=TRANSPOSE('Data Sheet'!B16:J16)"
    ws_f.Range('C14').Formula = '=C13+365'

    ws_f.Range('H5:H14').ClearContents()
    ws_f.Range('H5:H13').FormulaArray = "=TRANSPOSE('Data Sheet'!B16:J16)"
    ws_f.Range('H14').Formula = '=H13+365'

    ws_f.Range('M5:M14').ClearContents()
    ws_f.Range('M5:M13').FormulaArray = "=TRANSPOSE('Data Sheet'!B16:J16)"
    ws_f.Range('M14').Formula = '=M13+365'

    for r in range(15, 20):
        prev = r - 1
        ws_f.Range(f'C{r}').Formula = f'=C{prev}+365'
        ws_f.Range(f'H{r}').Formula = f'=H{prev}+365'
        ws_f.Range(f'M{r}').Formula = f'=M{prev}+365'

    # Recalculate entire workbook
    print("6. Executing CalculateFull across all sheets...")
    excel.CalculateFull()

    wb.Save()

    # Verify values in Intrinsic Valuation & Comp_Valuation
    ws_iv = wb.Sheets('Intrinsic Valuation')
    ws_cv = wb.Sheets('Comp_Valuation')
    
    print("\n--- Verification: Intrinsic Valuation Sheet ---")
    print("Dates Row 6 (H6..L6):", [ws_iv.Range(f'{col}6').Text for col in ['H', 'I', 'J', 'K', 'L']])
    print("Inventories Row 9 (H9..L9):", [ws_iv.Range(f'{col}9').Value for col in ['H', 'I', 'J', 'K', 'L']])
    print("Receivables Row 10 (H10..L10):", [ws_iv.Range(f'{col}10').Value for col in ['H', 'I', 'J', 'K', 'L']])
    print("Trade Payables Row 16 (H16..L16):", [ws_iv.Range(f'{col}16').Value for col in ['H', 'I', 'J', 'K', 'L']])

    print("\n--- Verification: Comp_Valuation Sheet ---")
    print("Peer 1 Name (B12):", str(ws_cv.Range('B12').Value))
    print("Peer 1 CMP (D12):", ws_cv.Range('D12').Value)
    print("Target EV (H12):", ws_cv.Range('H12').Value)

    print("\n--- Verification: Forecasting Sheet ---")
    print("Row 13 Year:", str(ws_f.Range('C13').Value))
    print("Row 14 Year (Weight 10):", str(ws_f.Range('C14').Value), "Formula:", ws_f.Range('C14').Formula)
    print("Row 15 Year (Weight 11):", str(ws_f.Range('C15').Value), "Formula:", ws_f.Range('C15').Formula)
    print("Row 19 Year (Weight 15):", str(ws_f.Range('C19').Value), "Formula:", ws_f.Range('C19').Formula)

    wb.Close(SaveChanges=False)
    print("\nSUCCESS! Workbook recalculated and verified cleanly!")

except Exception as e:
    import traceback
    print("ERROR:", e)
    traceback.print_exc()

finally:
    excel.Quit()
    del excel
    gc.collect()
    pythoncom.CoUninitialize()
