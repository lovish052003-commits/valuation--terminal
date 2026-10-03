"""
universal_valuation/workbook_map.py
===================================
Dynamic Workbook Mapping & Named Reference Layer.

Implements:
- Phase 4: Dynamic Data Sheet Mapping (DATA_SHEET_MAP dynamically resolved).
- Phase 23: Excel Formula Architecture (decoupled from fixed row numbers).
- Named Ranges support (MARKET_CAP, SHARES_OUTSTANDING, TOTAL_DEBT, CASH, REVENUE, EBITDA, EBIT, NET_INCOME, TOTAL_ASSETS, TOTAL_EQUITY, CURRENT_ASSETS, CURRENT_LIABILITIES).
"""

import re
from typing import Dict, Any, List, Optional, Tuple, Union
import openpyxl
from openpyxl.utils import get_column_letter


# Standard fallback row indices matching master template layout
DEFAULT_DATA_SHEET_ROWS = {
    # Meta (Columns B)
    "company_name": 1,
    "shares_formula": 6,
    "face_value": 7,
    "current_price": 8,
    "market_cap": 9,
    # Profit & Loss (Columns B to K)
    "pnl_date": 16,
    "revenue": 17,
    "raw_material": 18,
    "change_in_inventory": 19,
    "power_fuel": 20,
    "other_mfr_exp": 21,
    "employee_cost": 22,
    "selling_admin": 23,
    "other_expenses": 24,
    "other_income": 25,
    "depreciation": 26,
    "interest": 27,
    "pbt": 28,
    "tax": 29,
    "net_income": 30,
    "dividend_amount": 31,
    "ebitda": 32,
    # Quarters (Columns B to K)
    "quarters_date": 41,
    "q_revenue": 42,
    "q_expenses": 43,
    "q_other_income": 44,
    "q_depreciation": 45,
    "q_interest": 46,
    "q_pbt": 47,
    "q_tax": 48,
    "q_net_income": 49,
    "q_operating_profit": 50,
    # Balance Sheet (Columns B to K)
    "bs_date": 56,
    "equity_capital": 57,
    "reserves": 58,
    "debt": 59,
    "other_liabilities": 60,
    "total_liabilities": 61,
    "fixed_assets": 62,
    "cwip": 63,
    "investments": 64,
    "other_assets": 65,
    "total_assets": 66,
    "receivables": 67,
    "inventory": 68,
    "cash": 69,
    "shares_raw": 70,
    "face_value_bs": 72,
    "minority_interest": 73,
    "net_other_assets": 74,
    # Cash Flow (Columns B to K)
    "cf_date": 81,
    "cfo": 82,
    "cfi": 83,
    "cff": 84,
    "net_cash_flow": 85,
    # Price
    "historical_price": 90,
    "adjusted_equity_shares": 93
}


class WorkbookMap:
    """
    Dynamic Data Sheet Resolver & Workbook Schema Engine.
    Discovers exact row and column coordinates dynamically by scanning
    Column A labels in 'Data Sheet'.
    """

    def __init__(self, sheet_or_workbook=None):
        self.row_map: Dict[str, int] = dict(DEFAULT_DATA_SHEET_ROWS)
        self.period_cols: Dict[str, str] = {}  # e.g. {'Mar 2024': 'K', 'latest': 'K'}
        self.latest_col: str = "K"
        self.data_sheet_name: str = "Data Sheet"
        
        if sheet_or_workbook is not None:
            self.discover(sheet_or_workbook)

    def discover(self, sheet_or_workbook):
        """
        Scans Data Sheet to dynamically map row numbers from Column A text
        and column letters from row 16 / 56 dates.
        """
        ws = None
        if hasattr(sheet_or_workbook, 'sheetnames'):
            if 'Data Sheet' in sheet_or_workbook.sheetnames:
                ws = sheet_or_workbook['Data Sheet']
            elif 'DataSheet' in sheet_or_workbook.sheetnames:
                ws = sheet_or_workbook['DataSheet']
                self.data_sheet_name = 'DataSheet'
        elif hasattr(sheet_or_workbook, 'title'):
            ws = sheet_or_workbook
            self.data_sheet_name = ws.title

        if ws is None:
            return

        # 1. Scan Column A for row labels (rows 1 to 95)
        total_rows = min(100, ws.max_row if hasattr(ws, 'max_row') and ws.max_row else 95)
        found_liab_total = False
        
        for r in range(1, total_rows + 1):
            lbl_val = ws.cell(row=r, column=1).value
            if not lbl_val:
                continue
            lbl = str(lbl_val).strip().lower()

            # Meta
            if 'current price' in lbl or lbl == 'price':
                self.row_map['current_price'] = r
            elif 'market cap' in lbl:
                self.row_map['market_cap'] = r
            elif 'number of shares' in lbl and r < 15:
                self.row_map['shares_formula'] = r
            elif 'face value' in lbl and r < 15:
                self.row_map['face_value'] = r

            # P&L
            elif ('sales' in lbl or 'revenue' in lbl) and 15 < r < 20:
                self.row_map['revenue'] = r
            elif 'raw material' in lbl:
                self.row_map['raw_material'] = r
            elif 'depreciation' in lbl and r < 35:
                self.row_map['depreciation'] = r
            elif 'interest' in lbl and r < 35:
                self.row_map['interest'] = r
            elif 'profit before tax' in lbl and r < 35:
                self.row_map['pbt'] = r
            elif lbl == 'tax' and r < 35:
                self.row_map['tax'] = r
            elif 'net profit' in lbl and r < 35:
                self.row_map['net_income'] = r
            elif 'ebitda' in lbl and r < 35:
                self.row_map['ebitda'] = r

            # Balance Sheet
            elif 'equity share capital' in lbl or ('equity capital' in lbl and r > 50):
                self.row_map['equity_capital'] = r
            elif 'reserves' in lbl and r > 50:
                self.row_map['reserves'] = r
            elif ('borrowings' in lbl or 'total debt' in lbl) and r > 50:
                self.row_map['debt'] = r
            elif 'other liabilities' in lbl and r > 50:
                self.row_map['other_liabilities'] = r
            elif lbl == 'total' and 58 < r < 63 and not found_liab_total:
                self.row_map['total_liabilities'] = r
                found_liab_total = True
            elif 'net block' in lbl or 'fixed assets' in lbl:
                self.row_map['fixed_assets'] = r
            elif 'cwip' in lbl or 'capital work in progress' in lbl:
                self.row_map['cwip'] = r
            elif 'investments' in lbl and r > 55:
                self.row_map['investments'] = r
            elif 'other assets' in lbl and 60 < r < 66:
                self.row_map['other_assets'] = r
            elif (lbl == 'total' or 'total assets' in lbl) and 64 < r < 68:
                self.row_map['total_assets'] = r
            elif 'receivables' in lbl and r > 60:
                self.row_map['receivables'] = r
            elif 'inventory' in lbl and r > 60:
                self.row_map['inventory'] = r
            elif ('cash & bank' in lbl or 'cash & equivalents' in lbl or lbl == 'cash' or 'bank balance' in lbl) and 60 < r < 75:
                self.row_map['cash'] = r
            elif 'no. of equity shares' in lbl and r > 65:
                self.row_map['shares_raw'] = r

            # Cash Flow
            elif 'cash from operating' in lbl:
                self.row_map['cfo'] = r
            elif 'cash from investing' in lbl:
                self.row_map['cfi'] = r
            elif 'cash from financing' in lbl:
                self.row_map['cff'] = r
            elif 'net cash flow' in lbl:
                self.row_map['net_cash_flow'] = r

        # 2. Discover period columns from Row 16 or Row 56
        date_row = self.row_map.get('pnl_date', 16)
        active_cols = []
        for col_idx in range(2, 12):  # Col B to K
            c_val = ws.cell(row=date_row, column=col_idx).value
            if not c_val:
                c_val = ws.cell(row=self.row_map.get('bs_date', 56), column=col_idx).value
            if c_val is not None and str(c_val).strip() not in ('', 'None', 'Jan-00', '1900-01-00'):
                col_letter = get_column_letter(col_idx)
                active_cols.append(col_letter)
                self.period_cols[str(c_val).strip()] = col_letter

        if active_cols:
            self.latest_col = active_cols[-1]
        else:
            self.latest_col = "K"
        self.period_cols['latest'] = self.latest_col

    def get_row(self, field_name: str) -> int:
        """Returns the dynamic row number for a given canonical field."""
        f_norm = field_name.strip().lower()
        # Aliases
        if f_norm in ('sales', 'total_revenue'):
            f_norm = 'revenue'
        elif f_norm in ('total_debt', 'borrowings'):
            f_norm = 'debt'
        elif f_norm in ('cash_and_equivalents', 'cash_and_bank'):
            f_norm = 'cash'
        elif f_norm in ('pat', 'net_profit'):
            f_norm = 'net_income'
        elif f_norm in ('cmp', 'stock_price'):
            f_norm = 'current_price'
        elif f_norm in ('shares', 'shares_out'):
            f_norm = 'shares_formula'
        elif f_norm in ('equity', 'net_worth'):
            f_norm = 'reserves'  # Reserves row contains bulk of equity

        return self.row_map.get(f_norm, DEFAULT_DATA_SHEET_ROWS.get(f_norm, 17))

    def get_cell(self, field_name: str, period: str = 'latest') -> str:
        """Returns cell coordinates like 'K17' or 'B8'."""
        f_norm = field_name.strip().lower()
        if f_norm in ('current_price', 'market_cap', 'face_value', 'shares_formula'):
            # Meta cells are always in column B
            r = self.get_row(f_norm)
            return f"B{r}"
        
        col_letter = self.period_cols.get(period, self.latest_col)
        row_num = self.get_row(f_norm)
        return f"{col_letter}{row_num}"

    def get_ref(self, field_name: str, period: str = 'latest') -> str:
        """Returns fully qualified Excel reference like \"='Data Sheet'!K17\"."""
        cell_coord = self.get_cell(field_name, period)
        return f"='{self.data_sheet_name}'!{cell_coord}"

    def get_data_sheet_map(self) -> Dict[str, str]:
        """
        Phase 4: Exports the authoritative DATA_SHEET_MAP dictionary.
        """
        fields = [
            "revenue", "ebitda", "depreciation", "interest", "pbt", "tax", "net_income",
            "cash", "debt", "total_assets", "total_liabilities", "total_equity",
            "current_assets", "current_liabilities", "receivables", "inventory",
            "fixed_assets", "shares_outstanding", "market_cap", "current_price", "cfo", "cfi"
        ]
        res = {}
        for f in fields:
            if f == 'shares_outstanding':
                res[f] = f"='{self.data_sheet_name}'!B{self.row_map.get('shares_formula', 6)}"
            elif f == 'total_equity':
                res[f] = f"='{self.data_sheet_name}'!{self.latest_col}{self.row_map.get('reserves', 58)}"
            elif f == 'current_assets':
                res[f] = f"='{self.data_sheet_name}'!{self.latest_col}{self.row_map.get('other_assets', 65)}"
            elif f == 'current_liabilities':
                res[f] = f"='{self.data_sheet_name}'!{self.latest_col}{self.row_map.get('other_liabilities', 60)}"
            else:
                res[f] = self.get_ref(f, 'latest')
        return res

    def register_named_ranges_openpyxl(self, wb):
        """
        Phase 4 & 23: Registers standardized Excel Named Ranges on Data Sheet
        so formulas can reference named ranges directly.
        """
        import openpyxl.workbook.defined_name
        from openpyxl.workbook.defined_name import DefinedName

        ds_name = self.data_sheet_name
        k = self.latest_col
        
        named_targets = {
            "CURRENT_PRICE": f"'{ds_name}'!$B${self.row_map.get('current_price', 8)}",
            "MARKET_CAP": f"'{ds_name}'!$B${self.row_map.get('market_cap', 9)}",
            "SHARES_OUTSTANDING": f"'{ds_name}'!$B${self.row_map.get('shares_formula', 6)}",
            "TOTAL_DEBT": f"'{ds_name}'!${k}${self.row_map.get('debt', 59)}",
            "CASH": f"'{ds_name}'!${k}${self.row_map.get('cash', 69)}",
            "REVENUE": f"'{ds_name}'!${k}${self.row_map.get('revenue', 17)}",
            "EBITDA": f"'{ds_name}'!${k}${self.row_map.get('ebitda', 32)}",
            "NET_INCOME": f"'{ds_name}'!${k}${self.row_map.get('net_income', 30)}",
            "TOTAL_ASSETS": f"'{ds_name}'!${k}${self.row_map.get('total_assets', 66)}",
            "TOTAL_EQUITY": f"'{ds_name}'!${k}${self.row_map.get('reserves', 58)}",
            "TOTAL_LIABILITIES": f"'{ds_name}'!${k}${self.row_map.get('total_liabilities', 61)}",
            "CURRENT_ASSETS": f"'{ds_name}'!${k}${self.row_map.get('other_assets', 65)}",
            "CURRENT_LIABILITIES": f"'{ds_name}'!${k}${self.row_map.get('other_liabilities', 60)}"
        }

        for name, ref in named_targets.items():
            try:
                dn = DefinedName(name=name, attr_text=ref)
                wb.defined_names[name] = dn
            except Exception as e:
                pass

    def register_named_ranges_excel_com(self, workbook_com):
        """
        Registers named ranges using native Windows Excel COM.
        """
        ds_name = self.data_sheet_name
        k = self.latest_col

        named_targets = {
            "CURRENT_PRICE": f"='{ds_name}'!$B${self.row_map.get('current_price', 8)}",
            "MARKET_CAP": f"='{ds_name}'!$B${self.row_map.get('market_cap', 9)}",
            "SHARES_OUTSTANDING": f"='{ds_name}'!$B${self.row_map.get('shares_formula', 6)}",
            "TOTAL_DEBT": f"='{ds_name}'!${k}${self.row_map.get('debt', 59)}",
            "CASH": f"='{ds_name}'!${k}${self.row_map.get('cash', 69)}",
            "REVENUE": f"='{ds_name}'!${k}${self.row_map.get('revenue', 17)}",
            "EBITDA": f"='{ds_name}'!${k}${self.row_map.get('ebitda', 32)}",
            "NET_INCOME": f"='{ds_name}'!${k}${self.row_map.get('net_income', 30)}",
            "TOTAL_ASSETS": f"='{ds_name}'!${k}${self.row_map.get('total_assets', 66)}",
            "TOTAL_EQUITY": f"='{ds_name}'!${k}${self.row_map.get('reserves', 58)}",
            "TOTAL_LIABILITIES": f"='{ds_name}'!${k}${self.row_map.get('total_liabilities', 61)}",
            "CURRENT_ASSETS": f"='{ds_name}'!${k}${self.row_map.get('other_assets', 65)}",
            "CURRENT_LIABILITIES": f"='{ds_name}'!${k}${self.row_map.get('other_liabilities', 60)}"
        }

        for name, ref in named_targets.items():
            try:
                workbook_com.Names.Add(Name=name, RefersTo=ref)
            except Exception:
                pass
