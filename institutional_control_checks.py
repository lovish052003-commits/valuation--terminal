"""
institutional_control_checks.py
Provides institutional upgrades to universal valuation workbooks:
1. 'Control' sheet: Central assumptions panel with blue input cells, country table, scenario selector, and dynamic latest column offset.
2. 'Checks' sheet: 10-point automated model validation layer with 'MODEL STATUS: OK / CHECK REQUIRED' summary.
3. 2D Sensitivity Grid: WACC vs Terminal Growth Rate matrix.
4. Dynamic link wiring across DCF, WACC, Raw Data, Comp_Valuation, Beta-Regression, Dupont, Altman, and AI Valuation Summary.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Standard Investment Banking Color Palette
FILL_HEADER_DARK = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
FILL_SECTION = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
FILL_INPUT_BLUE = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid") # Classic IB Blue Input
FILL_PASS_GREEN = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
FILL_WARN_AMBER = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
FILL_MUTED = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

FONT_HEADER_WHITE = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONT_TITLE_DARK = Font(name="Calibri", size=13, bold=True, color="1F497D")
FONT_SECTION_BOLD = Font(name="Calibri", size=11, bold=True, color="1F497D")
FONT_INPUT_BLUE = Font(name="Calibri", size=11, bold=True, color="002060") # Classic IB Blue Text
FONT_BOLD = Font(name="Calibri", size=11, bold=True)
FONT_REGULAR = Font(name="Calibri", size=11, bold=False)
FONT_MUTED = Font(name="Calibri", size=10, italic=True, color="595959")
FONT_STATUS_OK = Font(name="Calibri", size=12, bold=True, color="375623")
FONT_STATUS_CHECK = Font(name="Calibri", size=12, bold=True, color="C65911")

BORDER_THIN = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
BORDER_HEADER = Border(
    left=Side(style='thin', color='1F497D'),
    right=Side(style='thin', color='1F497D'),
    top=Side(style='thin', color='1F497D'),
    bottom=Side(style='medium', color='1F497D')
)
BORDER_DOUBLE_BOTTOM = Border(
    top=Side(style='thin', color='000000'),
    bottom=Side(style='double', color='000000')
)

FMT_PCT = "0.0%"
FMT_PCT_2 = "0.00%"
FMT_NUM = "#,##0"
FMT_CURR = '"₹"\\ #,##0.00;\\("₹"\\ #,##0.00\\);\\-'
FMT_DATE = "YYYY-MM-DD"


def add_or_get_sheet(wb, title, target_index=None):
    """Retrieves sheet if exists, or creates at designated index."""
    if title in wb.sheetnames:
        return wb[title]
    if target_index is not None and target_index < len(wb.sheetnames):
        return wb.create_sheet(title, target_index)
    return wb.create_sheet(title)


def build_control_sheet(wb, screener_data=None, valuation_result=None):
    """
    Constructs the central 'Control' sheet:
    - Input cells for macroeconomic & company parameters (Blue filled).
    - Dynamic column index finding the latest reporting period.
    - Global Country reference table with Damodaran baseline metrics.
    - Scenario Analysis parameters table (Bear, Base, Bull).
    """
    ws = add_or_get_sheet(wb, 'Control', 1)
    ws.views.sheetView[0].showGridLines = True

    # Column widths
    ws.column_dimensions['A'].width = 4
    ws.column_dimensions['B'].width = 38
    ws.column_dimensions['C'].width = 22
    ws.column_dimensions['D'].width = 4
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 16
    ws.column_dimensions['G'].width = 16
    ws.column_dimensions['H'].width = 18
    ws.column_dimensions['I'].width = 4
    ws.column_dimensions['J'].width = 18
    ws.column_dimensions['K'].width = 16
    ws.column_dimensions['L'].width = 16
    ws.column_dimensions['M'].width = 16

    # 1. Main Header
    ws['B2'] = "INSTITUTIONAL MODEL CONTROL PANEL & ASSUMPTIONS"
    ws['B2'].font = FONT_TITLE_DARK
    ws['B3'] = "Single Source of Truth for Global Valuation Inputs, Scenarios & Route Configuration"
    ws['B3'].font = FONT_MUTED

    # Section 1: General & Reporting Inputs
    ws['B5'] = "1. MODEL METADATA & REPORTING SETTINGS"
    ws['B5'].font = FONT_SECTION_BOLD
    ws['B5'].fill = FILL_SECTION
    ws['C5'].fill = FILL_SECTION

    inputs_general = [
        (6, "Valuation Date", '=TODAY()', FMT_DATE, "Model execution date"),
        (7, "Reporting Currency", "INR", "@", "Primary currency"),
        (8, "Financial Reporting Unit", "Crores", "@", "Data Sheet financial scale"),
        (9, "Unit Scale Factor", 10000000, FMT_NUM, "Divisor for per-share calculation (10M for Cr)"),
        (10, "Target Country", "India", "@", "Drives Risk-Free Rate & ERP"),
        (11, "Target Sector / Industry", "Corporate / Non-Financial", "@", "Industry classification"),
        (12, "Valuation Framework Switch", "FCFF DCF", "@", "Options: FCFF DCF | Residual Income | Revenue-Led"),
        (13, "Benchmark Market Index", "NIFTY 50", "@", "Index used for regression beta"),
        (14, "Active Valuation Scenario", 2, "0", "1 = Bear | 2 = Base | 3 = Bull"),
    ]

    for row_idx, label, default_val, num_fmt, note in inputs_general:
        ws.cell(row_idx, 2, label).font = FONT_REGULAR
        cell_c = ws.cell(row_idx, 3, default_val)
        cell_c.font = FONT_INPUT_BLUE
        cell_c.fill = FILL_INPUT_BLUE
        cell_c.number_format = num_fmt
        cell_c.border = BORDER_THIN
        cell_c.alignment = Alignment(horizontal='right')

    # Section 2: Valuation & Cost of Capital Assumptions
    ws['B16'] = "2. COST OF CAPITAL & DCF ASSUMPTIONS"
    ws['B16'].font = FONT_SECTION_BOLD
    ws['B16'].fill = FILL_SECTION
    ws['C16'].fill = FILL_SECTION

    inputs_valuation = [
        (17, "Risk-Free Rate (Rf)", '=VLOOKUP(C10, E6:H12, 2, FALSE())', FMT_PCT, "Derived from Country Table"),
        (18, "Equity Risk Premium (ERP)", '=VLOOKUP(C10, E6:H12, 3, FALSE())', FMT_PCT, "Derived from Country Table"),
        (19, "Marginal Corporate Tax Rate", '=VLOOKUP(C10, E6:H12, 4, FALSE())', FMT_PCT, "Statutory marginal rate"),
        (20, "Effective Dynamic Tax Rate", '=MEDIAN(MAX(0.15, MIN(0.35, INDEX(\'Data Sheet\'!$B29:$K29, C26)/MAX(1, INDEX(\'Data Sheet\'!$B28:$K28, C26)))), C19)', FMT_PCT, "Tied to reported PBT/Tax"),
        (21, "Terminal Growth Rate (g)", 0.040, FMT_PCT, "Long-term GDP bound growth"),
        (22, "Reinvestment Rate Ceiling", 0.850, FMT_PCT, "Cap for volatile cycles"),
        (23, "Investments Fair-Value Haircut", 0.100, FMT_PCT, "Discount on non-operating holdings"),
        (24, "Minority Interest Method", "Book Value", "@", "Book Value | Market Multiple"),
    ]

    for row_idx, label, default_val, num_fmt, note in inputs_valuation:
        ws.cell(row_idx, 2, label).font = FONT_REGULAR
        cell_c = ws.cell(row_idx, 3, default_val)
        cell_c.font = FONT_INPUT_BLUE
        cell_c.fill = FILL_INPUT_BLUE
        cell_c.number_format = num_fmt
        cell_c.border = BORDER_THIN
        cell_c.alignment = Alignment(horizontal='right')

    # Section 3: Dynamic Period & Column Finder (Fixes Column K Hardcode!)
    ws['B25'] = "3. DYNAMIC STATEMENT PERIOD SELECTOR"
    ws['B25'].font = FONT_SECTION_BOLD
    ws['B25'].fill = FILL_SECTION
    ws['C25'].fill = FILL_SECTION

    ws.cell(26, 2, "Latest Historical Period Column Index").font = FONT_REGULAR
    cell_latest_col = ws.cell(26, 3, "=MATCH(MAX('Data Sheet'!$B56:$K56), 'Data Sheet'!$B56:$K56, 0)")
    cell_latest_col.font = FONT_BOLD
    cell_latest_col.fill = FILL_MUTED
    cell_latest_col.number_format = "0"
    cell_latest_col.border = BORDER_THIN

    ws.cell(27, 2, "Latest Statement Header Date").font = FONT_REGULAR
    cell_latest_hdr = ws.cell(27, 3, "=INDEX('Data Sheet'!$B56:$K56, C26)")
    cell_latest_hdr.font = FONT_BOLD
    cell_latest_hdr.fill = FILL_MUTED
    cell_latest_hdr.number_format = FMT_DATE
    cell_latest_hdr.border = BORDER_THIN

    ws.cell(28, 2, "Latest Stock Price (Data Sheet)").font = FONT_REGULAR
    cell_price = ws.cell(28, 3, "='Data Sheet'!B8")
    cell_price.font = FONT_BOLD
    cell_price.fill = FILL_MUTED
    cell_price.number_format = FMT_CURR
    cell_price.border = BORDER_THIN

    ws.cell(29, 2, "Default Credit Spread over Rf (used when debt is negligible)").font = FONT_REGULAR
    cell_cs = ws.cell(29, 3, 0.015)
    cell_cs.font = FONT_INPUT_BLUE
    cell_cs.fill = FILL_INPUT_BLUE
    cell_cs.number_format = FMT_PCT
    cell_cs.border = BORDER_THIN
    cell_cs.alignment = Alignment(horizontal='right')

    ws.cell(30, 2, "Max Credit Spread Cap (calc outside Rf..Rf+cap is rejected)").font = FONT_REGULAR
    cell_cap = ws.cell(30, 3, 0.050)
    cell_cap.font = FONT_INPUT_BLUE
    cell_cap.fill = FILL_INPUT_BLUE
    cell_cap.number_format = FMT_PCT
    cell_cap.border = BORDER_THIN
    cell_cap.alignment = Alignment(horizontal='right')

    # Section 4: Country Reference Table (Damodaran Institutional Baselines)
    ws['E5'] = "COUNTRY"
    ws['F5'] = "RISK-FREE (Rf)"
    ws['G5'] = "ERP"
    ws['H5'] = "MARGINAL TAX"
    for col_c in ['E', 'F', 'G', 'H']:
        ws[f'{col_c}5'].font = FONT_HEADER_WHITE
        ws[f'{col_c}5'].fill = FILL_HEADER_DARK
        ws[f'{col_c}5'].alignment = Alignment(horizontal='center')

    country_table = [
        ("India", 0.0680, 0.0650, 0.2517),
        ("United States", 0.0420, 0.0500, 0.2100),
        ("United Kingdom", 0.0440, 0.0550, 0.2500),
        ("Germany", 0.0250, 0.0500, 0.3000),
        ("Japan", 0.0110, 0.0550, 0.3062),
        ("Singapore", 0.0300, 0.0500, 0.1700),
        ("UAE", 0.0450, 0.0550, 0.0900),
    ]

    for idx, (cntry, rf_val, erp_val, tax_val) in enumerate(country_table, start=6):
        ws.cell(idx, 5, cntry).font = FONT_REGULAR
        ws.cell(idx, 6, rf_val).number_format = FMT_PCT
        ws.cell(idx, 7, erp_val).number_format = FMT_PCT
        ws.cell(idx, 8, tax_val).number_format = FMT_PCT
        for c_idx in range(5, 9):
            ws.cell(idx, c_idx).border = BORDER_THIN

    # Section 5: Scenario Matrix Table
    ws['J5'] = "SCENARIO"
    ws['K5'] = "GROWTH MULT"
    ws['L5'] = "MARGIN ADJ (NOT WIRED)"
    ws['M5'] = "WACC SPREAD"
    for col_c in ['J', 'K', 'L', 'M']:
        ws[f'{col_c}5'].font = FONT_HEADER_WHITE
        ws[f'{col_c}5'].fill = FILL_HEADER_DARK
        ws[f'{col_c}5'].alignment = Alignment(horizontal='center')

    scenarios = [
        ("1 - Bear Case", 0.70, -0.020, 0.010),
        ("2 - Base Case", 1.00, 0.000, 0.000),
        ("3 - Bull Case", 1.30, 0.020, -0.005),
    ]
    for idx, (sc_name, g_m, m_adj, w_sp) in enumerate(scenarios, start=6):
        ws.cell(idx, 10, sc_name).font = FONT_REGULAR
        ws.cell(idx, 11, g_m).number_format = "0.00"
        ws.cell(idx, 12, m_adj).number_format = "+0.0%;-0.0%;0.0%"
        ws.cell(idx, 13, w_sp).number_format = "+0.00%;-0.00%;0.00%"
        for c_idx in range(10, 14):
            ws.cell(idx, c_idx).border = BORDER_THIN

    ws['J10'] = "ACTIVE MULTIPLIERS:"
    ws['J10'].font = FONT_SECTION_BOLD
    ws['K10'] = "=CHOOSE(C14, K6, K7, K8)"
    ws['K10'].number_format = "0.00"
    ws['K10'].font = FONT_BOLD
    ws['L10'] = "=CHOOSE(C14, L6, L7, L8)"
    ws['L10'].number_format = "+0.0%;-0.0%;0.0%"
    ws['L10'].font = FONT_BOLD
    ws['M10'] = "=CHOOSE(C14, M6, M7, M8)"
    ws['M10'].number_format = "+0.00%;-0.00%;0.00%"
    ws['M10'].font = FONT_BOLD

    # Section 6: Optional Alternate 10-Year DCF Inputs (rows 32-36)
    ws['B32'] = "4. OPTIONAL ALTERNATE 10-YEAR DCF (parallel view - base DCF unchanged)"
    ws['B32'].font = FONT_SECTION_BOLD
    ws['B32'].fill = FILL_SECTION
    ws['C32'].fill = FILL_SECTION

    ws.cell(33, 2, "Year-1 EBIT Growth Override (blank = use model)").font = FONT_REGULAR
    cell_g_ovr = ws.cell(33, 3, None)
    cell_g_ovr.font = FONT_INPUT_BLUE
    cell_g_ovr.fill = FILL_INPUT_BLUE
    cell_g_ovr.number_format = FMT_PCT
    cell_g_ovr.border = BORDER_THIN

    ws.cell(34, 2, "Exit EV/EBIT Multiple (x)").font = FONT_REGULAR
    cell_m_ovr = ws.cell(34, 3, 18)
    cell_m_ovr.font = FONT_INPUT_BLUE
    cell_m_ovr.fill = FILL_INPUT_BLUE
    cell_m_ovr.number_format = "0"
    cell_m_ovr.border = BORDER_THIN
    cell_m_ovr.alignment = Alignment(horizontal='right')

    ws.cell(35, 2, "Terminal Method (Gordon Growth / Exit Multiple)").font = FONT_REGULAR
    cell_tm_ovr = ws.cell(35, 3, "Gordon Growth")
    cell_tm_ovr.font = FONT_INPUT_BLUE
    cell_tm_ovr.fill = FILL_INPUT_BLUE
    cell_tm_ovr.number_format = "@"
    cell_tm_ovr.border = BORDER_THIN
    cell_tm_ovr.alignment = Alignment(horizontal='right')

    ws.cell(36, 2, "Incremental ROIC Override (blank = model ROIC)").font = FONT_REGULAR
    cell_r_ovr = ws.cell(36, 3, None)
    cell_r_ovr.font = FONT_INPUT_BLUE
    cell_r_ovr.fill = FILL_INPUT_BLUE
    cell_r_ovr.number_format = FMT_PCT
    cell_r_ovr.border = BORDER_THIN

    return ws


def build_checks_sheet(wb, screener_data=None):
    """
    Constructs the automated 'Checks' validation sheet:
    - 10 automated mathematical and structural tests.
    - Top integrity banner: 'MODEL STATUS: OK' or 'MODEL STATUS: CHECK REQUIRED'.
    """
    ws = add_or_get_sheet(wb, 'Checks', 2)
    ws.views.sheetView[0].showGridLines = True

    ws.column_dimensions['A'].width = 4
    ws.column_dimensions['B'].width = 38
    ws.column_dimensions['C'].width = 48
    ws.column_dimensions['D'].width = 22
    ws.column_dimensions['E'].width = 16
    ws.column_dimensions['F'].width = 16

    # 1. Header & Global Banner
    ws['B2'] = "MODEL AUDIT & INTEGRITY CHECK PANEL"
    ws['B2'].font = FONT_TITLE_DARK

    ws['B3'] = "OVERALL MODEL STATUS:"
    ws['B3'].font = FONT_BOLD
    cell_status = ws['C3']
    cell_status.value = '=IF(AND(F6="PASS", F7="PASS", F8="PASS", F9="PASS", F10="PASS", F11="PASS", F12="PASS", F13="PASS", F14="PASS", F15="PASS"), "MODEL STATUS: OK", "MODEL STATUS: CHECK REQUIRED")'
    cell_status.font = FONT_STATUS_OK
    cell_status.fill = FILL_PASS_GREEN
    cell_status.alignment = Alignment(horizontal='center')
    cell_status.border = Border(top=Side(style='medium', color='375623'), bottom=Side(style='medium', color='375623'))

    # Table Header
    headers = [
        (5, 2, "CHECK ITEM"),
        (5, 3, "VALIDATION LOGIC & FORMULA"),
        (5, 4, "CURRENT VALUE / VARIANCE"),
        (5, 5, "THRESHOLD"),
        (5, 6, "STATUS")
    ]
    for r, c, txt in headers:
        cell = ws.cell(r, c, txt)
        cell.font = FONT_HEADER_WHITE
        cell.fill = FILL_HEADER_DARK
        cell.alignment = Alignment(horizontal='center' if c >= 4 else 'left')

    checks = [
        (6, "1. Balance Sheet Identity",
         "Total liabilities side ties to its components and to total assets",
         "=MAX(ABS(INDEX('Data Sheet'!$B61:$K61,Control!$C$26)-(INDEX('Data Sheet'!$B57:$K57,Control!$C$26)+INDEX('Data Sheet'!$B58:$K58,Control!$C$26)+INDEX('Data Sheet'!$B59:$K59,Control!$C$26)+INDEX('Data Sheet'!$B60:$K60,Control!$C$26))),ABS(INDEX('Data Sheet'!$B61:$K61,Control!$C$26)-INDEX('Data Sheet'!$B66:$K66,Control!$C$26)))",
         "< ₹ 5.0 Cr",
         '=IF(D6<5, "PASS", "CHECK")'),

        (7, "2. PBT Mathematical Tie",
         "Sales less all expenses plus other income less depreciation and interest ties to PBT",
         "=ABS(INDEX('Data Sheet'!$B17:$K17,Control!$C$26)-SUM(INDEX('Data Sheet'!$B$18:$K$24,0,Control!$C$26))+INDEX('Data Sheet'!$B25:$K25,Control!$C$26)-INDEX('Data Sheet'!$B26:$K26,Control!$C$26)-INDEX('Data Sheet'!$B27:$K27,Control!$C$26)-INDEX('Data Sheet'!$B28:$K28,Control!$C$26))/MAX(1,INDEX('Data Sheet'!$B28:$K28,Control!$C$26))",
         "< 1.5%",
         '=IF(D7<0.015, "PASS", "CHECK")'),

        (8, "3. Debt Schedule Concordance",
         "Data Sheet Borrowings ties to the independently pasted Raw FS borrowings (latest year)",
         "=ABS(INDEX('Data Sheet'!$B59:$K59,Control!$C$26)-LOOKUP(2,1/ISNUMBER('Raw FS'!$C7:$N7),'Raw FS'!$C7:$N7))",
         "< ₹ 5.0 Cr",
         '=IF(D8<5, "PASS", "CHECK")'),

        (9, "4. P&L Expenses Agreement",
         "Data Sheet total expenses (rows 18-24) tie to Raw FS total expenses",
         "=ABS(SUM(INDEX('Data Sheet'!$B$18:$K$24,0,Control!$C$26))-'Raw FS'!AD5)",
         "< ₹ 10.0 Cr",
         '=IF(D9<10, "PASS", "CHECK")'),

        (10, "5. Formula Error Counter",
         "Zero formula errors on Summary, DCF, WACC, Intrinsic Valuation, Comps and Beta tabs",
         "=SUMPRODUCT(--ISERROR('AI Valuation Summary'!A4:H47))+SUMPRODUCT(--ISERROR(DCF!A1:M55))+SUMPRODUCT(--ISERROR(WACC!A1:K46))+SUMPRODUCT(--ISERROR('Intrinsic Valuation'!A1:L65))+SUMPRODUCT(--ISERROR(Comp_Valuation!A1:Q39))+SUMPRODUCT(--ISERROR('Beta-Regression'!A1:O302))",
         "0 Errors",
         '=IF(D10=0, "PASS", "CHECK")'),

        (11, "6. Terminal Value EV Share",
         "Terminal Value percentage of Enterprise Value is sustainable",
         "=DCF!D34 / MAX(1, DCF!D35)",
         "< 75.0%",
         '=IF(D11<0.75, "PASS", "CHECK")'),

        (12, "7. Growth vs WACC Boundary",
         "Terminal growth rate is less than WACC and within Risk-Free rate",
         "=DCF!D20 - DCF!D19",
         "> 0.5%",
         '=IF(AND(D12>0.005, DCF!D19<=Control!$C$17), "PASS", "CHECK")'),

        (13, "8. Trading Price Freshness",
         "Market trading price quote is within acceptable institutional window",
         "=TODAY()-SUMPRODUCT(MAX(IFERROR(DATEVALUE('Raw Data'!G6:G252),IF(ISNUMBER('Raw Data'!G6:G252),'Raw Data'!G6:G252,0))))",
         "<= 45 Days",
         '=IF(D13<=45, "PASS", "CHECK")'),

        (14, "9. Peer Group Sample Size",
         "Comparable peer valuation set contains minimum robust observations",
         "=COUNTA(Comp_Valuation!C12:C26)",
         ">= 4 Peers",
         '=IF(D14>=4, "PASS", "CHECK")'),

        (15, "10. DCF Bridge Mathematical Tie",
         "Equity Value equals EV + Cash + Investments - Debt - Minority Interest",
         "=ABS(DCF!D40 - (DCF!D35 + DCF!D37 - DCF!D38 - DCF!D39))",
         "< ₹ 0.01",
         '=IF(D15<0.01, "PASS", "CHECK")'),
    ]

    for row_idx, title, desc, form_val, thresh, status_form in checks:
        ws.cell(row_idx, 2, title).font = FONT_BOLD
        ws.cell(row_idx, 3, desc).font = FONT_REGULAR
        
        c_val = ws.cell(row_idx, 4, form_val)
        c_val.font = FONT_REGULAR
        c_val.alignment = Alignment(horizontal='right')
        c_val.number_format = FMT_NUM if row_idx in [6, 8, 9, 10, 13, 14, 15] else (FMT_PCT if row_idx in [7, 11, 12] else "General")

        c_thresh = ws.cell(row_idx, 5, thresh)
        c_thresh.font = FONT_MUTED
        c_thresh.alignment = Alignment(horizontal='center')

        c_st = ws.cell(row_idx, 6, status_form)
        c_st.font = FONT_BOLD
        c_st.alignment = Alignment(horizontal='center')

        for c_i in range(2, 7):
            ws.cell(row_idx, c_i).border = BORDER_THIN

    # Advisory Checks Section (Rows 17 to 21)
    ws['B17'] = "ADVISORY CHECKS (informational - NOT included in overall status)"
    ws['B17'].font = FONT_SECTION_BOLD
    ws['B17'].fill = FILL_SECTION
    for c_hdr in ['C', 'D', 'E', 'F']:
        ws[f'{c_hdr}17'].fill = FILL_SECTION

    advisory_checks = [
        (18, "A1. DCF vs Market Divergence",
         "Absolute gap between DCF value per share and market price",
         "=ABS(DCF!D46)",
         "< 50%",
         '=IF(D18<0.5,"PASS","REVIEW")'),

        (19, "A2. WACC Sanity Range",
         "WACC sits inside a plausible band for the target market",
         "=DCF!D20",
         "8% - 16%",
         '=IF(AND(D19>=0.08,D19<=0.16),"PASS","REVIEW")'),

        (20, "A3. Market Cap Tie",
         "Market cap used in WACC weights ties to Price x Shares on Data Sheet",
         "=ABS(WACC!C35/'Data Sheet'!B9-1)",
         "< 2%",
         '=IF(D20<0.02,"PASS","REVIEW")'),

        (21, "A4. Reinvestment History Spike",
         "Largest historical reinvestment rate (blank-year / acquisition spikes distort median)",
         "=MAX('Intrinsic Valuation'!H52:L52)",
         "<= 100%",
         '=IF(D21<=1,"PASS","REVIEW")'),
    ]

    for row_idx, title, desc, form_val, thresh, status_form in advisory_checks:
        ws.cell(row_idx, 2, title).font = FONT_BOLD
        ws.cell(row_idx, 3, desc).font = FONT_REGULAR
        c_val = ws.cell(row_idx, 4, form_val)
        c_val.font = FONT_REGULAR
        c_val.alignment = Alignment(horizontal='right')
        c_val.number_format = FMT_PCT
        c_thresh = ws.cell(row_idx, 5, thresh)
        c_thresh.font = FONT_MUTED
        c_thresh.alignment = Alignment(horizontal='center')
        c_st = ws.cell(row_idx, 6, status_form)
        c_st.font = FONT_BOLD
        c_st.alignment = Alignment(horizontal='center')
        for c_i in range(2, 7):
            ws.cell(row_idx, c_i).border = BORDER_THIN

    return ws


def safe_set_cell(ws, coord, value, font=None, fill=None, alignment=None, border=None, number_format=None):
    """Safely writes a cell value, skipping non-top-left cells of merged ranges."""
    try:
        cell = ws[coord]
        if type(cell).__name__ == 'MergedCell':
            # Skip non-top-left merged cells to avoid openpyxl read-only exception
            return None
        cell.value = value
        if font: cell.font = font
        if fill: cell.fill = fill
        if alignment: cell.alignment = alignment
        if border: cell.border = border
        if number_format: cell.number_format = number_format
        return cell
    except Exception:
        return None


def apply_institutional_wiring(wb, screener_data=None, valuation_result=None):
    """
    Universally points hardcoded cells in existing sheets to 'Control' and 'Checks':
    1. DCF: Terminal Growth, Reinvestment Cap, 2D Sensitivity Table, Growth fading, Institutional Bridge.
    2. WACC: Dynamic 2-year average Debt/Interest formula using INDEX, Rf and ERP pointing to Control.
    3. Raw Data: Tax rate pointing to Control, dynamic peer columns without arbitrary fallbacks, target text.
    4. Beta-Regression: Minimum trading history guard (< 120 days), benchmark index pointing to Control.
    5. Dupont & Altman: 52-week High/Low dynamic formula, news blocks labeled 'manual, refresh per run'.
    6. AI Valuation Summary: Title tied to Data Sheet, MODEL STATUS banner tied to Checks.
    7. Comp_Valuation: Dynamic title, 15 peer rows, Target excluded from peer median, SOTP integration.
    """
    # 1. Build Control & Checks
    build_control_sheet(wb, screener_data, valuation_result)
    build_checks_sheet(wb, screener_data)

    # 2. Wire AI Valuation Summary
    if 'AI Valuation Summary' in wb.sheetnames:
        ws_sum = wb['AI Valuation Summary']
        safe_set_cell(ws_sum, 'A1', "='Data Sheet'!B1&\" - Institutional Equity Valuation Model\"")
        # Use Row 3 for status to avoid merged title block A1:G2 in certain templates
        safe_set_cell(ws_sum, 'A3', "MODEL AUDIT STATUS:", font=FONT_BOLD)
        safe_set_cell(ws_sum, 'B3', "=Checks!C3", font=FONT_STATUS_OK, fill=FILL_PASS_GREEN)
        safe_set_cell(ws_sum, 'A5', "=DCF!D45")
        safe_set_cell(ws_sum, 'B5', "=DCF!D43")
        safe_set_cell(ws_sum, 'C5', '=IFERROR((B5-A5)/A5,"n/a")')
        safe_set_cell(ws_sum, 'D5', '=IFERROR((B5-A5)/B5,"n/a")')
        safe_set_cell(ws_sum, 'E5', '=IF(NOT(ISNUMBER(C5)),"n/a - check price link",IF(C5>0.15,"UNDERVALUED / BUY",IF(C5<-0.15,"OVERVALUED / SELL","FAIRLY VALUED / HOLD")))')
        safe_set_cell(ws_sum, 'I4', "Verdict Confidence")
        safe_set_cell(ws_sum, 'I5', '=IF(NOT(ISNUMBER(C5)),"n/a",IF(ABS(C5)>0.5,"REVIEW: DCF >50% from price","OK"))')
        safe_set_cell(ws_sum, 'D10', "% (Control input; India 10Y G-Sec proxy)")
        safe_set_cell(ws_sum, 'D12', "Peer-median relevered beta")
        safe_set_cell(ws_sum, 'C38', "=DCF!D45")
        safe_set_cell(ws_sum, 'B41', "Intrinsic Value rolled to Valuation Date (info only)")
        safe_set_cell(ws_sum, 'C41', '=IFERROR(C37*(1+DCF!D20)^((Control!C6-DCF!H6)/365),"n/a")', number_format=FMT_CURR)
        safe_set_cell(ws_sum, 'D41', "₹")
        safe_set_cell(ws_sum, 'B42', "Alt 10-Year DCF Value (info only; set on Control rows 33-36)")
        safe_set_cell(ws_sum, 'C42', '=DCF!D81', number_format=FMT_CURR)
        safe_set_cell(ws_sum, 'D42', "₹")
        safe_set_cell(ws_sum, 'C44', "=IF(ISNUMBER(Comp_Valuation!Q25), Comp_Valuation!Q25*INDEX('Data Sheet'!$B30:$K30,Control!$C$26)/'Data Sheet'!B6, 'Data Sheet'!B8)")
        safe_set_cell(ws_sum, 'C45', "=IF(ISNUMBER(Comp_Valuation!P25), (Comp_Valuation!P25*INDEX('Data Sheet'!$B32:$K32,Control!$C$26)+DCF!D37-DCF!D38-DCF!D39)/'Data Sheet'!B6, 'Data Sheet'!B8)")
        safe_set_cell(ws_sum, 'C46', "=DCF!D43")
        safe_set_cell(ws_sum, 'C47', "=AVERAGE(C44:C46)")

    # 3. Wire DCF
    if 'DCF' in wb.sheetnames:
        ws_dcf = wb['DCF']
        # Sector / Model Switch Guard in B2
        safe_set_cell(ws_dcf, 'B2', '=IF(Control!$C$12="Residual Income", "FINANCIAL INSTITUTION: Refer to Residual Income Schedule", IF(OR(INDEX(\'Data Sheet\'!$B26:$K26, Control!$C$26)<=0, \'Intrinsic Valuation\'!L40<=0), "NOTICE: Negative EBIT / ROIC. Standard FCFF DCF Not Applicable. Use Relative Valuation.", "ACTIVE FCFF DCF MODEL"))', font=FONT_MUTED)

        # Terminal growth points to Control!C21
        safe_set_cell(ws_dcf, 'D19', "=Control!C21")

        # Terminal ROIC bounded by WACC
        safe_set_cell(ws_dcf, 'B22', "Terminal ROIC (Stable)")
        safe_set_cell(ws_dcf, 'D22', "=MAX('Intrinsic Valuation'!L40,D20)")
        safe_set_cell(ws_dcf, 'D21', "=D19/D22")

        # Fundamental growth: g = Reinvestment Rate * ROIC * Scenario Multiplier
        safe_set_cell(ws_dcf, 'H8', "='Intrinsic Valuation'!L38")
        for col_l in ['I', 'J', 'K', 'L', 'M']:
            prev_l = chr(ord(col_l) - 1)
            safe_set_cell(ws_dcf, f'{col_l}8', f"={prev_l}8*(1+{col_l}11*'Intrinsic Valuation'!$L$40*Control!$K$10)")

        # Reinvestment Rate fading to terminal rate D21 with Control cap
        safe_set_cell(ws_dcf, 'I11', "=MIN(Control!$C$22, MAX(0, 'Intrinsic Valuation'!$L$55))")
        safe_set_cell(ws_dcf, 'J11', "=MIN(Control!$C$22, $I$11+($M$11-$I$11)/4*1)")
        safe_set_cell(ws_dcf, 'K11', "=MIN(Control!$C$22, $I$11+($M$11-$I$11)/4*2)")
        safe_set_cell(ws_dcf, 'L11', "=MIN(Control!$C$22, $I$11+($M$11-$I$11)/4*3)")
        safe_set_cell(ws_dcf, 'M11', "=MIN(Control!$C$22, D21)")

        # Standardized Institutional Equity Bridge (Rows 35 to 46)
        safe_set_cell(ws_dcf, 'B37', "Add: Cash & Liquid Investments")
        safe_set_cell(ws_dcf, 'D37', "=INDEX('Data Sheet'!$B69:$K69, Control!$C$26) + INDEX('Data Sheet'!$B64:$K64, Control!$C$26)*(1-Control!$C$23)")
        safe_set_cell(ws_dcf, 'B38', "Less: Debt")
        safe_set_cell(ws_dcf, 'D38', "=INDEX('Data Sheet'!$B59:$K59, Control!$C$26)")
        safe_set_cell(ws_dcf, 'B39', "Less: Minority Interest")
        safe_set_cell(ws_dcf, 'D39', "=IF(Control!$C$24=\"Market Multiple\", INDEX('Data Sheet'!$B73:$K73, Control!$C$26)*1.2, INDEX('Data Sheet'!$B73:$K73, Control!$C$26))")
        safe_set_cell(ws_dcf, 'B40', "Equity Value")
        safe_set_cell(ws_dcf, 'D40', "=D35+D37-D38-D39")
        safe_set_cell(ws_dcf, 'B41', "No. of Shares")
        safe_set_cell(ws_dcf, 'D41', "=INDEX('Data Sheet'!$B70:$K70, Control!$C$26)/Control!$C$9")
        safe_set_cell(ws_dcf, 'B43', "Equity Value per Share")
        safe_set_cell(ws_dcf, 'D43', "=D40/D41")
        safe_set_cell(ws_dcf, 'B45', "Share Price")
        safe_set_cell(ws_dcf, 'D45', "='Data Sheet'!B8")
        safe_set_cell(ws_dcf, 'B46', "Upside / (Downside)")
        safe_set_cell(ws_dcf, 'D46', "=D43/D45-1")

        # 2D SENSITIVITY GRID (Rows 49 to 57): WACC vs Terminal Growth Rate
        safe_set_cell(ws_dcf, 'B49', "2D SENSITIVITY ANALYSIS: IMPLIED SHARE PRICE (INR)", font=FONT_SECTION_BOLD)
        safe_set_cell(ws_dcf, 'B50', "WACC \\ Terminal Growth (g)", font=FONT_MUTED)

        # Terminal growth headers (Cols E to I)
        g_steps = [-0.010, -0.005, 0, 0.005, 0.010]
        for col_idx, g_step in enumerate(g_steps, start=5):
            cl = chr(ord('A') + col_idx - 1)
            g_val_str = f"+({g_step})" if g_step != 0 else "+(0)"
            safe_set_cell(ws_dcf, f'{cl}50', f"=D19{g_val_str}", font=FONT_HEADER_WHITE, fill=FILL_HEADER_DARK, alignment=Alignment(horizontal='center'), number_format=FMT_PCT)

        # WACC rows (Rows 51 to 55)
        wacc_steps = [-0.010, -0.005, 0, 0.005, 0.010]
        for r_offset, w_step in enumerate(wacc_steps, start=51):
            w_val_str = f"+({w_step})" if w_step != 0 else "+(0)"
            safe_set_cell(ws_dcf, f'B{r_offset}', f"=D20{w_val_str}", font=FONT_HEADER_WHITE, fill=FILL_HEADER_DARK, alignment=Alignment(horizontal='center'), number_format=FMT_PCT)

            # Implied per-share formula across matrix: Dynamic DCF discounting with n/a guard
            for col_idx, g_step in enumerate(g_steps, start=5):
                cl = chr(ord('A') + col_idx - 1)
                safe_set_cell(ws_dcf, f'{cl}{r_offset}', f"=IF($B{r_offset}<={cl}$50,\"n/a\",MAX(0,(SUMPRODUCT($I$12:$M$12,(1+$B{r_offset})^(-$I$13:$M$13))+($M$10*(1+{cl}$50)*(1-MIN(Control!$C$22,{cl}$50/$D$22)))/($B{r_offset}-{cl}$50)*(1+$B{r_offset})^(-$M$13)+$D$37-$D$38-$D$39)/$D$41))", border=BORDER_THIN, alignment=Alignment(horizontal='right'), number_format=FMT_CURR)

        # Alternate 10-Year DCF (Rows 58 to 83)
        safe_set_cell(ws_dcf, 'B58', "ALTERNATE 10-YEAR DCF (parallel view - does not feed the Summary verdict; inputs on Control rows 33-36)", font=FONT_SECTION_BOLD)
        safe_set_cell(ws_dcf, 'B60', "Year-1 EBIT growth")
        safe_set_cell(ws_dcf, 'D60', "=IF(ISNUMBER(Control!$C$33),Control!$C$33,$I$8/$H$8-1)", number_format=FMT_PCT)
        safe_set_cell(ws_dcf, 'B61', "Starting incremental ROIC")
        safe_set_cell(ws_dcf, 'D61', "=IF(ISNUMBER(Control!$C$36),Control!$C$36,'Intrinsic Valuation'!$L$40)", number_format=FMT_PCT)

        safe_set_cell(ws_dcf, 'B64', "Year")
        for y_idx in range(1, 11):
            cl_y = chr(ord('D') + y_idx - 1)
            safe_set_cell(ws_dcf, f'{cl_y}64', y_idx, alignment=Alignment(horizontal='center'))
            safe_set_cell(ws_dcf, f'{cl_y}65', f"=$D$60+($D$19-$D$60)*({cl_y}64-1)/9", number_format=FMT_PCT)
            safe_set_cell(ws_dcf, f'{cl_y}66', f"=$D$61+($D$22-$D$61)*({cl_y}64-1)/9", number_format=FMT_PCT)
            safe_set_cell(ws_dcf, f'{cl_y}67', f"=MIN(Control!$C$22,MAX(0,{cl_y}65/{cl_y}66))", number_format=FMT_PCT)

        safe_set_cell(ws_dcf, 'B65', "EBIT growth (fades to terminal g)")
        safe_set_cell(ws_dcf, 'B66', "ROIC (converges to terminal ROIC)")
        safe_set_cell(ws_dcf, 'B67', "Reinvestment rate (g / ROIC, capped)")

        safe_set_cell(ws_dcf, 'B68', "EBIT")
        safe_set_cell(ws_dcf, 'D68', "=$H$8*(1+D65)", number_format=FMT_CURR)
        for y_idx in range(2, 11):
            cl_curr = chr(ord('D') + y_idx - 1)
            cl_prev = chr(ord('D') + y_idx - 2)
            safe_set_cell(ws_dcf, f'{cl_curr}68', f"={cl_prev}68*(1+{cl_curr}65)", number_format=FMT_CURR)

        safe_set_cell(ws_dcf, 'B69', "NOPAT")
        safe_set_cell(ws_dcf, 'B70', "FCFF")
        safe_set_cell(ws_dcf, 'B71', "Discounting factor (mid-year)")
        safe_set_cell(ws_dcf, 'B72', "PV of FCFF")
        for y_idx in range(1, 11):
            cl_y = chr(ord('D') + y_idx - 1)
            safe_set_cell(ws_dcf, f'{cl_y}69', f"={cl_y}68*(1-$H$9)", number_format=FMT_CURR)
            safe_set_cell(ws_dcf, f'{cl_y}70', f"={cl_y}69*(1-{cl_y}67)", number_format=FMT_CURR)
            safe_set_cell(ws_dcf, f'{cl_y}71', f"=(1+$D$20)^(-({cl_y}64-0.5))", number_format="0.0000")
            safe_set_cell(ws_dcf, f'{cl_y}72', f"={cl_y}70*{cl_y}71", number_format=FMT_CURR)

        safe_set_cell(ws_dcf, 'B74', "Terminal value - Gordon Growth")
        safe_set_cell(ws_dcf, 'D74', "=M69*(1+$D$19)*(1-MIN(Control!$C$22,$D$19/$D$22))/($D$20-$D$19)", number_format=FMT_CURR)
        safe_set_cell(ws_dcf, 'B75', "Terminal value - Exit Multiple (x Year-10 EBIT)")
        safe_set_cell(ws_dcf, 'D75', "=Control!$C$34*M68", number_format=FMT_CURR)
        safe_set_cell(ws_dcf, 'B76', "Terminal method used")
        safe_set_cell(ws_dcf, 'D76', "=Control!$C$35")
        safe_set_cell(ws_dcf, 'B77', "PV of terminal value")
        safe_set_cell(ws_dcf, 'D77', '=IF(D76="Exit Multiple",D75*(1+$D$20)^(-10),D74*M71)', number_format=FMT_CURR)
        safe_set_cell(ws_dcf, 'B78', "PV of 10-year FCFF")
        safe_set_cell(ws_dcf, 'D78', "=SUM(D72:M72)", number_format=FMT_CURR)
        safe_set_cell(ws_dcf, 'B79', "Enterprise value")
        safe_set_cell(ws_dcf, 'D79', "=D77+D78", number_format=FMT_CURR)
        safe_set_cell(ws_dcf, 'B80', "Equity value")
        safe_set_cell(ws_dcf, 'D80', "=D79+$D$37-$D$38-$D$39", number_format=FMT_CURR)
        safe_set_cell(ws_dcf, 'B81', "Alt DCF value per share", font=FONT_BOLD)
        safe_set_cell(ws_dcf, 'D81', "=D80/$D$41", font=FONT_BOLD, number_format=FMT_CURR)
        safe_set_cell(ws_dcf, 'B82', "Upside / (Downside) vs price")
        safe_set_cell(ws_dcf, 'D82', "=D81/$D$45-1", number_format=FMT_PCT)
        safe_set_cell(ws_dcf, 'B83', "Terminal value % of EV")
        safe_set_cell(ws_dcf, 'D83', "=D77/D79", number_format=FMT_PCT)

    # 4. Wire WACC
    if 'WACC' in wb.sheetnames:
        ws_wacc = wb['WACC']
        # Cost of Debt with dynamic credit spread fallback for zero/distorted debt
        safe_set_cell(ws_wacc, 'E26', "=IF(AVERAGE(INDEX('Data Sheet'!$B59:$K59, Control!$C$26-1), INDEX('Data Sheet'!$B59:$K59, Control!$C$26))<=0, Control!$C$17+Control!$C$29, IF(OR((INDEX('Data Sheet'!$B27:$K27, Control!$C$26)/AVERAGE(INDEX('Data Sheet'!$B59:$K59, Control!$C$26-1), INDEX('Data Sheet'!$B59:$K59, Control!$C$26)))<Control!$C$17, (INDEX('Data Sheet'!$B27:$K27, Control!$C$26)/AVERAGE(INDEX('Data Sheet'!$B59:$K59, Control!$C$26-1), INDEX('Data Sheet'!$B59:$K59, Control!$C$26)))>Control!$C$17+Control!$C$30), Control!$C$17+Control!$C$29, INDEX('Data Sheet'!$B27:$K27, Control!$C$26)/AVERAGE(INDEX('Data Sheet'!$B59:$K59, Control!$C$26-1), INDEX('Data Sheet'!$B59:$K59, Control!$C$26))))")
        safe_set_cell(ws_wacc, 'K26', "=Control!C17")
        safe_set_cell(ws_wacc, 'K27', "=Control!C18")
        safe_set_cell(ws_wacc, 'E27', "=Control!C20")
        safe_set_cell(ws_wacc, 'K46', "=(K40*K41)+(K43*K44)+Control!$M$10")
        safe_set_cell(ws_wacc, 'G28', "Levered Beta (peer-median unlevered, relevered at target D/E)")

        # Exclude target company (Row 16) from peer average and median
        for c_let in ['G', 'H', 'I', 'J', 'K']:
            safe_set_cell(ws_wacc, f'{c_let}20', f"=AVERAGE({c_let}14:{c_let}15,{c_let}17:{c_let}18)")
            safe_set_cell(ws_wacc, f'{c_let}21', f"=MEDIAN({c_let}14:{c_let}15,{c_let}17:{c_let}18)")

    # 5. Wire Raw Data
    if 'Raw Data' in wb.sheetnames:
        ws_raw = wb['Raw Data']
        safe_set_cell(ws_raw, 'T24', "=Control!C19")
        safe_set_cell(ws_raw, 'O14', "='Data Sheet'!B1")
        safe_set_cell(ws_raw, 'O26', "='Data Sheet'!B1")
        for r in range(12, 27):
            safe_set_cell(ws_raw, f'W{r}', f'=IF(X{r}>0, U{r}/X{r}, "N/A")')

    # 6. Wire Beta-Regression
    if 'Beta-Regression' in wb.sheetnames:
        ws_beta = wb['Beta-Regression']
        safe_set_cell(ws_beta, 'F7', "=Control!C13&\" Daily Returns\"")
        safe_set_cell(ws_beta, 'O11', "=IF(COUNT(D11:D252)<120, \"Insufficient History (< 120 Days)\", SLOPE(D11:D252, H11:H252))")
        safe_set_cell(ws_beta, 'O12', "=IF(COUNT(D11:D252)<120, \"N/A\", _xlfn.COVARIANCE.S(D11:D252, H11:H252)/_xlfn.VAR.S(H11:H252))")
        safe_set_cell(ws_beta, 'L9', "=IF(ISNUMBER(O11), O11, 1)")

    # 7. Wire Dupont Analysis & Altman's Z Score
    for sh_name in ['Dupont Analysis', "Altman's Z Score"]:
        if sh_name in wb.sheetnames:
            ws_rep = wb[sh_name]
            safe_set_cell(ws_rep, 'B5', '="52 Week High "&TEXT(MAX(\'Raw Data\'!H6:H252),"#,##0.00")&" / Low "&TEXT(MIN(\'Raw Data\'!H6:H252),"#,##0.00")', font=FONT_BOLD)
            safe_set_cell(ws_rep, 'B6', '="About the Company: "&\'Data Sheet\'!B1&" (Manual, Refresh Per Run)"', font=FONT_SECTION_BOLD)

    # 8. Wire Comp_Valuation
    if 'Comp_Valuation' in wb.sheetnames:
        ws_comp = wb['Comp_Valuation']
        safe_set_cell(ws_comp, 'B30', '=\'Data Sheet\'!B1&" Comparable Valuation"')
        safe_set_cell(ws_comp, 'B11', '=IF(COUNTA(C12:C26)>=4, "Peers >= 4 (Valid Sample)", "WARNING: Insufficient Peers (< 4)")', font=FONT_MUTED)

    # 9. Clean Ratio Analysis Ranges (removes duplicated sheet name syntax in formulas)
    if 'Ratio Analysis' in wb.sheetnames:
        ws_ra = wb['Ratio Analysis']
        for col_i in range(3, 13):
            c_l = get_column_letter(col_i)
            safe_set_cell(ws_ra, f'{c_l}21', f"=IFERROR('Historical FS'!{c_l}25/SUM('Historical FS'!{c_l}51:{c_l}53),0)")
            safe_set_cell(ws_ra, f'{c_l}23', f"=IFERROR('Historical FS'!{c_l}37/SUM('Historical FS'!{c_l}51:{c_l}52),0)")
            safe_set_cell(ws_ra, f'{c_l}31', f"=IFERROR('Historical FS'!{c_l}7/SUM('Historical FS'!{c_l}51:{c_l}52),0)")

    # 10. Wire Intrinsic Valuation Organic Cross-Check (Rows 55, 65, 67 to 71)
    if 'Intrinsic Valuation' in wb.sheetnames:
        ws_iv = wb['Intrinsic Valuation']
        safe_set_cell(ws_iv, 'J55', "5 Col Median (H:L)", font=FONT_BOLD)
        safe_set_cell(ws_iv, 'J65', "5 Col Median (H:L)", font=FONT_BOLD)
        safe_set_cell(ws_iv, 'B67', "ORGANIC REINVESTMENT CROSS-CHECK (info only - not wired to DCF)")
        safe_set_cell(ws_iv, 'B68', "Organic Reinvestment (Net Capex + Change in WC)")
        for col_l in ['I', 'J', 'K', 'L']:
            safe_set_cell(ws_iv, f'{col_l}68', f"={col_l}44+{col_l}45", number_format=FMT_CURR)
        safe_set_cell(ws_iv, 'B69', "Organic Reinvestment Rate (on NOPAT)")
        for col_l in ['I', 'J', 'K', 'L']:
            safe_set_cell(ws_iv, f'{col_l}69', f"={col_l}68/{col_l}49", font=FONT_BOLD, number_format=FMT_PCT)
        safe_set_cell(ws_iv, 'B70', "Organic 4-Year Median")
        safe_set_cell(ws_iv, 'L70', "=MEDIAN(I69:L69)", font=FONT_BOLD, number_format=FMT_PCT)
        safe_set_cell(ws_iv, 'B71', "Currently used in DCF (L55 - includes FY22 blank-year and FY25 acquisition values)")
        safe_set_cell(ws_iv, 'L71', "=L55", font=FONT_BOLD, number_format=FMT_PCT)

    return wb
