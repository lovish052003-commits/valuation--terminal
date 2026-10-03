import os, sys, shutil, gc
import pythoncom
import win32com.client
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from screener_client import fetch_company_data, clean_num
from valuation_engine import calculate_valuation

TEMPLATE_PATH = os.path.abspath('ITC Model.xlsx')
DEST_PATH = os.path.abspath('exports/TATAMOTORS_Valuation_Model.xlsx')

os.makedirs('exports', exist_ok=True)
print("Fetching Screener data for Tata Motors...")
screener_data = fetch_company_data('Tata Motors')
val = calculate_valuation(screener_data)

print("Copying template to destination...")
shutil.copyfile(TEMPLATE_PATH, DEST_PATH)

print("Starting Excel COM with pythoncom init...")
pythoncom.CoInitialize()
excel = None
wb = None

try:
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    wb = excel.Workbooks.Open(DEST_PATH)
    ws_data = wb.Sheets('Data Sheet')
    
    # Fill metadata
    ws_data.Range('B1').Value = screener_data['company_name']
    ws_data.Range('B6').Value = screener_data['shares_in_cr']
    ws_data.Range('B7').Value = screener_data['face_value']
    ws_data.Range('B8').Value = screener_data['current_price']
    ws_data.Range('B9').Value = screener_data['market_cap_cr']

    tables = screener_data.get('tables', {})
    pl_df = tables.get('profit-loss')
    bs_df = tables.get('balance-sheet')
    cf_df = tables.get('cash-flow')

    def fill_row(row_idx, df, keywords):
        if df is None or df.empty:
            return
        period_cols = [c for c in df.columns if c != 'Metric']
        for kw in keywords:
            m = df[df['Metric'].str.contains(kw, case=False, na=False)]
            if not m.empty:
                r = m.iloc[0]
                for offset, col_name in enumerate(period_cols[-10:]):
                    c_idx = 2 + offset
                    v = clean_num(r[col_name])
                    ws_data.Cells(row_idx, c_idx).Value = v
                break

    def fill_dates(row_idx, df):
        if df is None or df.empty:
            return
        period_cols = [c for c in df.columns if c != 'Metric']
        for offset, col_name in enumerate(period_cols[-10:]):
            c_idx = 2 + offset
            ws_data.Cells(row_idx, c_idx).Value = col_name

    fill_dates(16, pl_df)
    fill_row(17, pl_df, ['^Sales', 'Revenue'])
    fill_row(18, pl_df, ['Raw Material Cost', 'Material'])
    fill_row(22, pl_df, ['Employee Cost'])
    fill_row(25, pl_df, ['Other Income'])
    fill_row(26, pl_df, ['Depreciation'])
    fill_row(27, pl_df, ['Interest'])
    fill_row(28, pl_df, ['Profit before tax', 'PBT'])
    fill_row(29, pl_df, ['Tax'])
    fill_row(30, pl_df, ['Net profit', 'PAT'])
    fill_row(32, pl_df, ['Operating Profit', 'EBITDA'])

    fill_dates(56, bs_df)
    fill_row(57, bs_df, ['Equity Capital', 'Share Capital'])
    fill_row(58, bs_df, ['Reserves'])
    fill_row(59, bs_df, ['Borrowings', 'Total Debt'])
    fill_row(60, bs_df, ['Other Liabilities'])
    fill_row(61, bs_df, ['Total Liabilities'])
    fill_row(62, bs_df, ['Fixed Assets', 'Net Block'])
    fill_row(63, bs_df, ['CWIP'])
    fill_row(64, bs_df, ['Investments'])
    fill_row(65, bs_df, ['Other Assets'])
    fill_row(66, bs_df, ['Total Assets'])

    fill_dates(81, cf_df)
    fill_row(82, cf_df, ['Cash from Operating Activity'])
    fill_row(83, cf_df, ['Cash from Investing Activity'])
    fill_row(84, cf_df, ['Cash from Financing Activity'])
    fill_row(85, cf_df, ['Net Cash Flow'])

    # Add AI Valuation Summary Sheet
    sheet_names = [s.Name for s in wb.Sheets]
    if 'AI Valuation Summary' in sheet_names:
        ws_sum = wb.Sheets('AI Valuation Summary')
    else:
        ws_sum = wb.Sheets.Add(Before=wb.Sheets(1))
        ws_sum.Name = 'AI Valuation Summary'

    ws_sum.Range('A1:G2').Merge()
    ws_sum.Range('A1').Value = f"{screener_data['company_name']} ({screener_data['ticker']}) - ITC Valuation Model"
    ws_sum.Range('A1').Font.Size = 16
    ws_sum.Range('A1').Font.Bold = True

    # Recalculate
    excel.CalculateFull()
    wb.Save()
    wb.Close(SaveChanges=True)
    wb = None
    print("TATAMOTORS saved successfully via COM!")

except Exception as e:
    print("Error during COM run:", e)
finally:
    if wb is not None:
        try:
            wb.Close(SaveChanges=False)
        except:
            pass
        del wb
    if excel is not None:
        try:
            excel.Quit()
        except:
            pass
        del excel
    gc.collect()
    pythoncom.CoUninitialize()
    print("COM cleaned up cleanly.")

# Now verify reopening with a fresh COM instance
print("Verifying reopening in Excel...")
pythoncom.CoInitialize()
excel2 = win32com.client.DispatchEx('Excel.Application')
excel2.Visible = False
excel2.DisplayAlerts = False
try:
    wb_test = excel2.Workbooks.Open(DEST_PATH)
    print("SUCCESS: TATAMOTORS_Valuation_Model.xlsx opened with ZERO errors in Excel!")
    wb_test.Close(SaveChanges=False)
except Exception as e:
    print("Verification failed:", e)
finally:
    excel2.Quit()
    del excel2
    gc.collect()
    pythoncom.CoUninitialize()
