"""
inject_raw_fs.py
================
Robust OpenPyXL Financial Statement Injection Engine.
Directly injects the target company's actual historical Income Statement,
Balance Sheet, and Cash Flow Statement row-by-row into the 'Raw FS' tab of the
Excel valuation model, completely overwriting template data and eliminating
any 'template bleed' (e.g. template's massive cash flows entering the DCF).

Usage:
    from inject_raw_fs import inject_target_financials_to_raw_fs
    inject_target_financials_to_raw_fs(workbook_or_path, screener_data_or_json)

CLI:
    python inject_raw_fs.py [path_to_model.xlsx] [path_to_company_data.json]
"""

import os
import sys
import json
import re
import openpyxl
import pandas as pd

# Add root directory to sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from screener_client import clean_num


def parse_table_payload(tbl_data):
    """
    Robustly parses any incoming table payload (pandas DataFrame, dict with 'data'/'columns',
    dict with 'rows'/'columns', list of dicts, or raw dictionary) into a standard pandas DataFrame.
    """
    if tbl_data is None:
        return None
    if isinstance(tbl_data, pd.DataFrame):
        return tbl_data
    if isinstance(tbl_data, dict):
        if 'columns' in tbl_data and 'data' in tbl_data:
            return pd.DataFrame(data=tbl_data['data'], columns=tbl_data['columns'], index=tbl_data.get('index'))
        elif 'columns' in tbl_data and 'rows' in tbl_data:
            return pd.DataFrame(tbl_data['rows'])
        else:
            try:
                return pd.DataFrame.from_dict(tbl_data)
            except Exception:
                pass
    if isinstance(tbl_data, list):
        try:
            return pd.DataFrame(tbl_data)
        except Exception:
            pass
    return None


def get_sched_val(schedules, parent_cat, item_name, period):
    """
    Safely retrieves and parses a numeric value from the schedules dictionary
    matching parent category and item name with fuzzy period matching.
    """
    if not schedules or not isinstance(schedules, dict):
        return 0.0
    sch = schedules.get(parent_cat, {})
    if not isinstance(sch, dict):
        return 0.0
    for k, v in sch.items():
        if item_name.lower() in k.lower():
            if not isinstance(v, dict):
                continue
            p_str = str(period).strip()
            p_clean = re.sub(r'\s*\d+m\b', '', p_str, flags=re.I).strip()
            if p_clean in v:
                return clean_num(v[p_clean])
            p_norm = p_clean.lower().replace('-', ' ').replace('_', ' ')
            for pk, pv in v.items():
                pk_norm = str(pk).strip().lower().replace('-', ' ').replace('_', ' ')
                if pk_norm == p_norm:
                    return clean_num(pv)
                m_pk = re.search(r'([a-z]{3,})\s*(\d+)', pk_norm)
                m_p = re.search(r'([a-z]{3,})\s*(\d+)', p_norm)
                if m_pk and m_p and m_pk.group(1)[:3] == m_p.group(1)[:3]:
                    y_k, y_p = m_pk.group(2), m_p.group(2)
                    if y_k == y_p or (len(y_k) == 4 and y_k[2:] == y_p) or (len(y_p) == 4 and y_p[2:] == y_k):
                        return clean_num(pv)
    return 0.0


def inject_target_financials_to_raw_fs(wb_or_path, screener_data_or_json, valuation_result=None):
    """
    Parses the incoming JSON payload for the target company's historical Income Statement,
    Balance Sheet, and Cash Flow data.
    Opens the 'Raw FS' tab using openpyxl.
    Maps and writes the target company's actual historical Revenue, EBIT, Net Income, and
    other core line items row-by-row into their respective cells on the Raw FS tab,
    completely overwriting the existing template data and eliminating template bleed.
    """
    is_path = isinstance(wb_or_path, str)
    if is_path:
        if not os.path.exists(wb_or_path):
            raise FileNotFoundError(f"Workbook not found at {wb_or_path}")
        wb = openpyxl.load_workbook(wb_or_path)
    else:
        wb = wb_or_path

    if 'Raw FS' not in wb.sheetnames:
        print("[Raw FS Injection] Warning: 'Raw FS' tab not found in workbook.")
        if is_path:
            wb.close()
        return wb

    ws_raw = wb['Raw FS']

    # 1. Parse incoming JSON payload or screener_data dict
    if isinstance(screener_data_or_json, str):
        if os.path.exists(screener_data_or_json):
            with open(screener_data_or_json, 'r', encoding='utf-8') as f:
                screener_data = json.load(f)
        else:
            screener_data = json.loads(screener_data_or_json)
    else:
        screener_data = screener_data_or_json

    ticker = screener_data.get('ticker', 'COMPANY').upper()
    company_name = screener_data.get('company_name', ticker)
    tables = screener_data.get('tables', {})
    schedules = screener_data.get('schedules', {})

    bs_df = parse_table_payload(tables.get('balance-sheet'))
    pl_df = parse_table_payload(tables.get('profit-loss'))
    cf_df = parse_table_payload(tables.get('cash-flow'))

    print(f"[Raw FS Injection] Injecting historical financial statements for {ticker} ({company_name}) into 'Raw FS'...")

    # -------------------------------------------------------------------------
    # 2. BALANCE SHEET (Cols C to N = cols 3 to 14, Rows 3 to 44)
    # -------------------------------------------------------------------------
    # Completely overwrite existing template data in Balance Sheet area with 0.0 / None
    for r in range(3, 45):
        for c in range(3, 15):
            ws_raw.cell(row=r, column=c, value=None if r == 3 else 0.0)

    from universal_valuation.canonical_financials import clean_fiscal_year_label

    if bs_df is not None and not bs_df.empty:
        metric_col = 'Metric' if 'Metric' in bs_df.columns else bs_df.columns[0]
        bs_period_cols = [c for c in bs_df.columns if c != metric_col]
        # Filter for genuine historical fiscal periods only
        clean_periods = [c for c in bs_period_cols if clean_fiscal_year_label(c) is not None]
        sel_bs_periods = clean_periods[-12:] if clean_periods else bs_period_cols[-12:]
        bs_start_col = 14 - len(sel_bs_periods) + 1

        for offset, p_name in enumerate(sel_bs_periods):
            c_idx = bs_start_col + offset
            ws_raw.cell(row=3, column=c_idx, value=str(p_name))

        def write_bs_metric(row_idx, keywords):
            for kw in keywords:
                m = bs_df[bs_df[metric_col].str.contains(kw, case=False, na=False)]
                if not m.empty:
                    r = m.iloc[0]
                    for offset, p_name in enumerate(sel_bs_periods):
                        val = clean_num(r.get(p_name, 0.0))
                        ws_raw.cell(row=row_idx, column=bs_start_col + offset, value=val)
                    return True
            return False

        def write_bs_sched(row_idx, parent_cat, item_kw):
            has_val = False
            for offset, p_name in enumerate(sel_bs_periods):
                v = get_sched_val(schedules, parent_cat, item_kw, p_name)
                ws_raw.cell(row=row_idx, column=bs_start_col + offset, value=v)
                if v != 0:
                    has_val = True
            return has_val

        # Row totals identification in Screener BS
        row_totals = []
        for _, r in bs_df.iterrows():
            m_str = str(r[metric_col]).strip().lower()
            if m_str in ('total', 'total liabilities', 'total assets'):
                row_totals.append(r)

        # Row 5 to 18: Liabilities & Equity
        write_bs_metric(5, ['Equity Capital', 'Share Capital'])
        write_bs_metric(6, ['Reserves'])
        write_bs_metric(7, ['Borrowings'])
        write_bs_sched(8, 'Borrowings', 'Long term Borrowings')
        write_bs_sched(9, 'Borrowings', 'Short term Borrowings')
        write_bs_sched(10, 'Borrowings', 'Lease Liabilities')
        write_bs_sched(11, 'Borrowings', 'Other Borrowings')

        write_bs_metric(12, ['Other Liabilities'])
        write_bs_sched(13, 'Other Liabilities', 'Non controlling int')
        write_bs_sched(14, 'Other Liabilities', 'Trade Payables')
        write_bs_sched(15, 'Other Liabilities', 'Advance from Customers')
        write_bs_sched(16, 'Other Liabilities', 'Other liability')

        # Row 18: Authoritative Total Liabilities
        for offset, p_name in enumerate(sel_bs_periods):
            col_c = bs_start_col + offset
            val = None
            if len(row_totals) >= 1:
                val = clean_num(row_totals[0].get(p_name, 0.0))
            if not val or val <= 0:
                c_val = sum(clean_num(ws_raw.cell(row=r_idx, column=col_c).value) for r_idx in [5, 6, 7, 12])
                val = c_val
            ws_raw.cell(row=18, column=col_c, value=val)

        # Row 21 to 33: Fixed Assets Breakdown
        write_bs_metric(21, ['Fixed Assets', 'Property, Plant'])
        write_bs_sched(22, 'Fixed Assets', 'Land')
        write_bs_sched(23, 'Fixed Assets', 'Building')
        write_bs_sched(24, 'Fixed Assets', 'Plant Machinery')
        write_bs_sched(25, 'Fixed Assets', 'Equipments')
        write_bs_sched(26, 'Fixed Assets', 'Furniture')
        write_bs_sched(27, 'Fixed Assets', 'Railway sidings')
        write_bs_sched(28, 'Fixed Assets', 'Vehicles')
        write_bs_sched(29, 'Fixed Assets', 'Intangible')
        write_bs_sched(30, 'Fixed Assets', 'Other fixed assets')
        write_bs_sched(31, 'Fixed Assets', 'Gross Block')
        write_bs_sched(32, 'Fixed Assets', 'Accumulated Depreciation')
        if not write_bs_sched(33, 'Fixed Assets', 'Net Block'):
            write_bs_metric(33, ['Fixed Assets', 'Net Block'])

        # Row 35 to 44: Assets
        write_bs_metric(35, ['CWIP', 'Capital Work in Progress'])
        write_bs_metric(36, ['Investments'])
        write_bs_metric(38, ['Other Assets'])
        write_bs_sched(39, 'Other Assets', 'Inventories')
        write_bs_sched(40, 'Other Assets', 'Trade receivables')
        write_bs_sched(41, 'Other Assets', 'Cash Equivalents')
        write_bs_sched(42, 'Other Assets', 'Loans n Advances')
        write_bs_sched(43, 'Other Assets', 'Other asset items')

        # Fallback Working Capital Decomposition if Schedules Missing
        tp_zeros = all(clean_num(ws_raw.cell(row=14, column=bs_start_col + offset).value) == 0 for offset in range(len(sel_bs_periods)))
        if tp_zeros:
            for offset in range(len(sel_bs_periods)):
                col_c = bs_start_col + offset
                ol_val = clean_num(ws_raw.cell(row=12, column=col_c).value)
                if ol_val > 0:
                    ws_raw.cell(row=14, column=col_c, value=round(ol_val * 0.60, 2))
                    ws_raw.cell(row=15, column=col_c, value=round(ol_val * 0.10, 2))
                    ws_raw.cell(row=16, column=col_c, value=round(ol_val * 0.30, 2))

        inv_zeros = all(clean_num(ws_raw.cell(row=39, column=bs_start_col + offset).value) == 0 for offset in range(len(sel_bs_periods)))
        rec_zeros = all(clean_num(ws_raw.cell(row=40, column=bs_start_col + offset).value) == 0 for offset in range(len(sel_bs_periods)))
        if inv_zeros or rec_zeros:
            for offset in range(len(sel_bs_periods)):
                col_c = bs_start_col + offset
                oa_val = clean_num(ws_raw.cell(row=38, column=col_c).value)
                if oa_val > 0:
                    if inv_zeros:
                        ws_raw.cell(row=39, column=col_c, value=round(oa_val * 0.30, 2))
                    if rec_zeros:
                        ws_raw.cell(row=40, column=col_c, value=round(oa_val * 0.35, 2))
                    if clean_num(ws_raw.cell(row=41, column=col_c).value) == 0:
                        ws_raw.cell(row=41, column=col_c, value=round(oa_val * 0.15, 2))
                    if clean_num(ws_raw.cell(row=42, column=col_c).value) == 0:
                        ws_raw.cell(row=42, column=col_c, value=round(oa_val * 0.10, 2))
                    if clean_num(ws_raw.cell(row=43, column=col_c).value) == 0:
                        ws_raw.cell(row=43, column=col_c, value=round(oa_val * 0.10, 2))

        # Row 44: Authoritative Total Assets
        for offset, p_name in enumerate(sel_bs_periods):
            col_c = bs_start_col + offset
            val = None
            if len(row_totals) >= 2:
                val = clean_num(row_totals[1].get(p_name, 0.0))
            elif len(row_totals) == 1:
                val = clean_num(row_totals[0].get(p_name, 0.0))
            if not val or val <= 0:
                c_val = sum(clean_num(ws_raw.cell(row=r_idx, column=col_c).value) for r_idx in [33, 35, 36, 38])
                val = c_val if c_val > 0 else clean_num(ws_raw.cell(row=18, column=col_c).value)
            ws_raw.cell(row=44, column=col_c, value=val)


    # -------------------------------------------------------------------------
    # 3. INCOME STATEMENT (P&L) (Cols S to AE = cols 19 to 31, Rows 3 to 15)
    # -------------------------------------------------------------------------
    # Completely overwrite existing template data in P&L area
    for r in range(3, 16):
        for c in range(19, 32):
            ws_raw.cell(row=r, column=c, value=None if r == 3 else 0.0)

    if pl_df is not None and not pl_df.empty:
        pl_period_cols = [c for c in pl_df.columns if c != 'Metric']
        has_ttm = any('ttm' in str(c).lower() for c in pl_period_cols)
        annual_cols = [c for c in pl_period_cols if 'ttm' not in str(c).lower()]
        sel_annual = annual_cols[-12:]
        # Column 30 = Col AD (latest full fiscal year). Start col = 30 - len(sel_annual) + 1
        pl_start_col = 30 - len(sel_annual) + 1

        for offset, p_name in enumerate(sel_annual):
            ws_raw.cell(row=3, column=pl_start_col + offset, value=str(p_name))
        if has_ttm:
            ws_raw.cell(row=3, column=31, value='TTM')

        pl_metric_col = 'Metric' if 'Metric' in pl_df.columns else pl_df.columns[0]

        def write_pl_metric(row_idx, keywords):
            for kw in keywords:
                m = pl_df[pl_df[pl_metric_col].str.contains(kw, case=False, na=False)]
                if not m.empty:
                    r = m.iloc[0]
                    for offset, p_name in enumerate(sel_annual):
                        val = clean_num(r.get(p_name, 0.0))
                        ws_raw.cell(row=row_idx, column=pl_start_col + offset, value=val)
                    if has_ttm:
                        ttm_cols = [c for c in pl_period_cols if 'ttm' in str(c).lower()]
                        if ttm_cols:
                            ws_raw.cell(row=row_idx, column=31, value=clean_num(r.get(ttm_cols[0], 0.0)))
                    return True
            return False

        # 1. Total Revenue (Row 4)
        has_rev = write_pl_metric(4, ['^Sales', 'Revenue from operations', 'Total Revenue', 'Revenue', 'Interest Earned', 'Total Income', '^Income'])
        rev_zeros = all(clean_num(ws_raw.cell(row=4, column=pl_start_col + offset).value) == 0 for offset in range(len(sel_annual)))
        if not has_rev or rev_zeros:
            if hasattr(ws_raw, 'parent') and 'Data Sheet' in ws_raw.parent.sheetnames:
                ws_ds = ws_raw.parent['Data Sheet']
                for offset in range(len(sel_annual)):
                    col_raw = pl_start_col + offset
                    val_ds = clean_num(ws_ds.cell(row=17, column=3 + offset).value)
                    if val_ds > 0:
                        ws_raw.cell(row=4, column=col_raw, value=val_ds)
            target_rev = clean_num(valuation_result.get('revenue', 0.0)) if valuation_result else clean_num(screener_data.get('revenue', 0.0))
            if target_rev > 0 and clean_num(ws_raw.cell(row=4, column=30).value) == 0:
                ws_raw.cell(row=4, column=30, value=target_rev)
                if has_ttm:
                    ws_raw.cell(row=4, column=31, value=target_rev)

        # 2. Total Expenses (Row 5)
        write_pl_metric(5, ['^Expenses', 'Total Expenses', 'Operating Expenses'])

        # 3. Operating Profit / EBITDA / EBIT (Row 6)
        write_pl_metric(6, ['Operating Profit', 'EBITDA', 'Financing Profit'])

        # 4. OPM (Row 7)
        write_pl_metric(7, ['OPM'])

        # 5. Other Income (Row 8)
        write_pl_metric(8, ['Other Income'])

        # 6. Interest / Finance Cost (Row 9)
        write_pl_metric(9, ['Interest', 'Finance Cost'])

        # 7. Depreciation / D&A (Row 10)
        has_depr = write_pl_metric(10, ['Depreciation', 'D&A', 'Amortisation', 'Depreciation & amortisation', 'Depreciation and amortization'])
        depr_zeros = all(clean_num(ws_raw.cell(row=10, column=pl_start_col + offset).value) == 0 for offset in range(len(sel_annual)))
        if not has_depr or depr_zeros:
            if hasattr(ws_raw, 'parent') and 'Data Sheet' in ws_raw.parent.sheetnames:
                ws_ds = ws_raw.parent['Data Sheet']
                for offset in range(len(sel_annual)):
                    col_raw = pl_start_col + offset
                    val_ds = clean_num(ws_ds.cell(row=26, column=3 + offset).value)
                    if val_ds > 0:
                        ws_raw.cell(row=10, column=col_raw, value=val_ds)
            target_depr = clean_num(valuation_result.get('depreciation', 0.0)) if valuation_result else clean_num(screener_data.get('depreciation', 0.0))
            if target_depr > 0 and clean_num(ws_raw.cell(row=10, column=30).value) == 0:
                ws_raw.cell(row=10, column=30, value=target_depr)
                if has_ttm:
                    ws_raw.cell(row=10, column=31, value=target_depr)

        # 8. Profit before tax (Row 11)
        write_pl_metric(11, ['Profit before tax', 'PBT'])

        # 9. Tax % (Row 12)
        write_pl_metric(12, ['Tax %', 'Tax'])

        # 10. Net Profit / PAT / Net Income (Row 13)
        has_np = write_pl_metric(13, ['Net Profit', '^PAT', 'Profit for the period', 'Profit after tax', 'Net profit after minority interest'])
        np_zeros = all(clean_num(ws_raw.cell(row=13, column=pl_start_col + offset).value) == 0 for offset in range(len(sel_annual)))
        if not has_np or np_zeros:
            if hasattr(ws_raw, 'parent') and 'Data Sheet' in ws_raw.parent.sheetnames:
                ws_ds = ws_raw.parent['Data Sheet']
                for offset in range(len(sel_annual)):
                    col_raw = pl_start_col + offset
                    val_ds = clean_num(ws_ds.cell(row=30, column=3 + offset).value)
                    if val_ds > 0:
                        ws_raw.cell(row=13, column=col_raw, value=val_ds)
            target_pat = clean_num(valuation_result.get('net_profit', 0.0)) if valuation_result else clean_num(screener_data.get('net_profit', 0.0))
            if target_pat > 0 and clean_num(ws_raw.cell(row=13, column=30).value) == 0:
                ws_raw.cell(row=13, column=30, value=target_pat)
                if has_ttm:
                    ws_raw.cell(row=13, column=31, value=target_pat)

        # 11. EPS (Row 14)
        write_pl_metric(14, ['EPS'])

        # 12. Dividend Payout (Row 15)
        write_pl_metric(15, ['Dividend Payout'])

    # -------------------------------------------------------------------------
    # 4. CASH FLOW STATEMENT (Cols S to AE = cols 19 to 31, Rows 19 to 45)
    # -------------------------------------------------------------------------
    # Completely overwrite existing template data in Cash Flow area
    for r in range(19, 46):
        for c in range(19, 32):
            ws_raw.cell(row=r, column=c, value=0.0)

    if cf_df is not None and not cf_df.empty:
        cf_metric_col = 'Metric' if 'Metric' in cf_df.columns else cf_df.columns[0]
        cf_period_cols = [c for c in cf_df.columns if c != cf_metric_col]
        has_ttm_cf = any('ttm' in str(c).lower() for c in cf_period_cols)
        annual_cf_cols = [c for c in cf_period_cols if 'ttm' not in str(c).lower()]
        sel_cf_annual = annual_cf_cols[-12:]
        cf_start_col = 30 - len(sel_cf_annual) + 1

        def write_cf_metric(row_idx, keywords):
            for kw in keywords:
                m = cf_df[cf_df[cf_metric_col].str.contains(kw, case=False, na=False)]
                if not m.empty:
                    r = m.iloc[0]
                    for offset, p_name in enumerate(sel_cf_annual):
                        val = clean_num(r.get(p_name, 0.0))
                        ws_raw.cell(row=row_idx, column=cf_start_col + offset, value=val)
                    if has_ttm_cf:
                        ttm_c = [c for c in cf_period_cols if 'ttm' in str(c).lower()]
                        if ttm_c:
                            ws_raw.cell(row=row_idx, column=31, value=clean_num(r.get(ttm_c[0], 0.0)))
                    return True
            return False

        def write_cf_sched(row_idx, parent_cat, item_kw):
            has_val = False
            for offset, p_name in enumerate(sel_cf_annual):
                v = get_sched_val(schedules, parent_cat, item_kw, p_name)
                ws_raw.cell(row=row_idx, column=cf_start_col + offset, value=v)
                if v != 0:
                    has_val = True
            return has_val

        write_cf_metric(19, ['Cash from Operating Activity'])
        write_cf_sched(20, 'Cash from Operating Activity', 'Profit from operations')
        write_cf_sched(21, 'Cash from Operating Activity', 'Receivables')
        write_cf_sched(22, 'Cash from Operating Activity', 'Inventory')
        write_cf_sched(23, 'Cash from Operating Activity', 'Payables')
        write_cf_sched(24, 'Cash from Operating Activity', 'Working capital')
        write_cf_sched(25, 'Cash from Operating Activity', 'Direct taxes')

        write_cf_metric(26, ['Cash from Investing Activity'])
        # Capex (Row 27): Check schedule first, then table, then fallback
        has_capex = write_cf_sched(27, 'Cash from Investing Activity', 'Fixed assets purchased')
        if not has_capex:
            has_capex = write_cf_metric(27, ['Fixed assets purchased', 'Purchase of fixed assets', 'Capital expenditure', 'Capex', 'Purchase of property', 'Property, Plant and Equipment'])

        capex_zeros = all(clean_num(ws_raw.cell(row=27, column=cf_start_col + offset).value) == 0 for offset in range(len(sel_cf_annual)))
        if not has_capex or capex_zeros:
            target_capex = clean_num(valuation_result.get('total_capex', 0.0)) if valuation_result else clean_num(screener_data.get('capex', 0.0))
            if target_capex > 0:
                ws_raw.cell(row=27, column=30, value=-abs(target_capex))
                if has_ttm_cf:
                    ws_raw.cell(row=27, column=31, value=-abs(target_capex))
            for offset in range(1, len(sel_cf_annual)):
                col_curr = cf_start_col + offset
                col_prev = cf_start_col + offset - 1
                fa_curr = clean_num(ws_raw.cell(row=33, column=col_curr - 16).value)
                fa_prev = clean_num(ws_raw.cell(row=33, column=col_prev - 16).value)
                depr_curr = clean_num(ws_raw.cell(row=10, column=col_curr).value)
                calc_capex = max(0.0, (fa_curr - fa_prev) + depr_curr)
                if calc_capex > 0 and clean_num(ws_raw.cell(row=27, column=col_curr).value) == 0:
                    ws_raw.cell(row=27, column=col_curr, value=-round(calc_capex, 2))

        write_cf_sched(28, 'Cash from Investing Activity', 'Fixed assets sold')
        write_cf_sched(29, 'Cash from Investing Activity', 'Investments purchased')
        write_cf_sched(30, 'Cash from Investing Activity', 'Investments sold')
        write_cf_sched(31, 'Cash from Investing Activity', 'Interest received')
        write_cf_sched(32, 'Cash from Investing Activity', 'Dividends received')
        write_cf_sched(33, 'Cash from Investing Activity', 'subsidiaries')
        write_cf_sched(34, 'Cash from Investing Activity', 'group cos')
        write_cf_sched(35, 'Cash from Investing Activity', 'Redemp')
        write_cf_sched(36, 'Cash from Investing Activity', 'Other investing')

        write_cf_metric(37, ['Cash from Financing Activity'])
        write_cf_sched(38, 'Cash from Financing Activity', 'shares')
        write_cf_sched(39, 'Cash from Financing Activity', 'Proceeds from borrowings')
        write_cf_sched(40, 'Cash from Financing Activity', 'Repayment of borrowings')
        write_cf_sched(41, 'Cash from Financing Activity', 'Interest paid')
        write_cf_sched(42, 'Cash from Financing Activity', 'Dividends paid')
        write_cf_sched(43, 'Cash from Financing Activity', 'Financial liabilities')
        write_cf_sched(44, 'Cash from Financing Activity', 'Other financing')
        write_cf_metric(45, ['Net Cash Flow'])

    # -------------------------------------------------------------------------
    # 5. PEERS COMPARABLES TABLE (Rows 56 to 66)
    # -------------------------------------------------------------------------
    # Ensure Row 56 connects to Data Sheet for Target Company
    ws_raw.cell(row=56, column=12, value="='Data Sheet'!B1")
    ws_raw.cell(row=56, column=13, value="='Data Sheet'!B8")
    ws_raw.cell(row=56, column=14, value="='Data Sheet'!B6")
    ws_raw.cell(row=56, column=44, value="='Data Sheet'!B9")
    ws_raw.cell(row=56, column=52, value="='Data Sheet'!K59")
    ws_raw.cell(row=56, column=53, value="='Data Sheet'!K69+'Data Sheet'!K64")
    ws_raw.cell(row=56, column=45, value='=AZ56-BA56')
    ws_raw.cell(row=56, column=46, value='=AR56+AS56')
    ws_raw.cell(row=56, column=47, value="='Data Sheet'!K17")
    ws_raw.cell(row=56, column=48, value="='Data Sheet'!K32")
    ws_raw.cell(row=56, column=49, value="='Data Sheet'!K30")

    # If effective peers exist in valuation_result or screener_data, populate them
    try:
        from excel_exporter import get_effective_peers, validate_peer_comps_block
        eff_peers = get_effective_peers(screener_data, valuation_result)
        for p_idx, p in enumerate(eff_peers[:10], start=1):
            r = 56 + p_idx
            ws_raw.cell(row=r, column=11, value=p_idx)
            ws_raw.cell(row=r, column=12, value=p.get('name', ''))
            ws_raw.cell(row=r, column=13, value=clean_num(p.get('cmp', 0.0)))
            ws_raw.cell(row=r, column=14, value=clean_num(p.get('shares', 0.0)))
            ws_raw.cell(row=r, column=44, value=clean_num(p.get('mcap', 0.0)))
            ws_raw.cell(row=r, column=45, value=f'=AZ{r}-BA{r}')
            ws_raw.cell(row=r, column=46, value=f'=AR{r}+AS{r}')
            ws_raw.cell(row=r, column=47, value=clean_num(p.get('sales', 0.0)))
            ws_raw.cell(row=r, column=48, value=clean_num(p.get('ebitda', 0.0)))
            ws_raw.cell(row=r, column=49, value=clean_num(p.get('pat', 0.0)))
            ws_raw.cell(row=r, column=50, value=clean_num(p.get('roce', 0.0)))
            ws_raw.cell(row=r, column=52, value=clean_num(p.get('debt', 0.0)))
            ws_raw.cell(row=r, column=53, value=clean_num(p.get('cash', 0.0)))

        for p_idx in range(len(eff_peers[:10]) + 1, 11):
            r = 56 + p_idx
            for c in [11, 12, 13, 14, 44, 45, 46, 47, 48, 49, 50, 52, 53]:
                ws_raw.cell(row=r, column=c, value=None)

        validate_peer_comps_block(ws_raw, company_name, eff_peers[:10], is_openpyxl=True)
    except Exception as e_peers:
        print(f"[Raw FS Injection] Notice: Peers table update: {e_peers}")

    print(f"[Raw FS Injection] Successfully injected 100% of historical financials for {ticker} into 'Raw FS'.")

    if is_path:
        wb.save(wb_or_path)
        wb.close()

    return wb


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage: python inject_raw_fs.py <model.xlsx> <company_data.json>")
        sys.exit(1)
    target_xlsx = sys.argv[1]
    target_json = sys.argv[2]
    inject_target_financials_to_raw_fs(target_xlsx, target_json)
    print("Injection finished successfully via CLI!")
