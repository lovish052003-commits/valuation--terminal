import sys, os
sys.path.insert(0, os.path.abspath('.'))
import excel_exporter
print("Has populate_cash_flow_statement_sheet_openpyxl:", hasattr(excel_exporter, 'populate_cash_flow_statement_sheet_openpyxl'))
print("Has populate_ratio_analysis_sheet_openpyxl:", hasattr(excel_exporter, 'populate_ratio_analysis_sheet_openpyxl'))
print("Has populate_dupont_altman_sheets_openpyxl:", hasattr(excel_exporter, 'populate_dupont_altman_sheets_openpyxl'))
