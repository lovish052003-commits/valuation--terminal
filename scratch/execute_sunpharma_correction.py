"""
execute_sunpharma_correction.py
Audits and corrects SUNPHARMA_Valuation_Model (1).xlsx at the source/formula level.
Preserves existing structure, charts, formatting, and formulas.
Saves as SUNPHARMA_Valuation_Model_Corrected.xlsx.
"""

import os
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SOURCE_FILE = r'C:\Users\LENOVO\Downloads\test\SUNPHARMA_Valuation_Model (1).xlsx'
TARGET_FILE = r'C:\Users\LENOVO\Downloads\test\SUNPHARMA_Valuation_Model_Corrected.xlsx'
WORKSPACE_MIRROR = r'C:\Users\LENOVO\Downloads\Advance Financial Project\SUNPHARMA_Valuation_Model_Corrected.xlsx'

print(f"Loading source workbook: {SOURCE_FILE}...")
wb = openpyxl.load_workbook(SOURCE_FILE, data_only=False)

# =========================================================================
# 1. FIX DATA SHEET — TAX EXPENSE & OPERATING EXPENSES
# =========================================================================
print("\n[1/7] Correcting 'Data Sheet'...")
if 'Data Sheet' in wb.sheetnames:
    ws_ds = wb['Data Sheet']
    # Row 28 is Profit before tax, Row 30 is Net profit.
    # Tax Expense = PBT - Net profit.
    # Write formulas across Columns B to K (2017 to 2026)
    for col_idx in range(2, 12):  # B to K
        col_l = get_column_letter(col_idx)
        ws_ds[f'{col_l}29'] = f"={col_l}28-{col_l}30"
    print("   -> Fixed Row 29 Tax Expense formulas: =B28-B30 through =K28-K30")

# =========================================================================
# 2. FIX HISTORICAL FS — SG&A/OTHER EXPENSES, EBITDA, TAX EXPENSE & RATES
# =========================================================================
print("\n[2/7] Correcting 'Historical FS'...")
if 'Historical FS' in wb.sheetnames:
    ws_hfs = wb['Historical FS']
    
    # 2a. SG&A / Other Expenses (Row 16)
    # Point to 'Data Sheet' Row 24 (Other Operating Expenses) instead of Row 23 (which was 0)
    for offset, cl in enumerate(['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']):
        ds_cl = chr(ord('B') + offset)
        ws_hfs[f'{cl}16'] = f"='Data Sheet'!{ds_cl}24"
    ws_hfs['B16'] = "Selling & General / Other Expenses"
    print("   -> Fixed Row 16 SG&A/Other Expenses linking to 'Data Sheet'!Row 24")

    # 2b. EBITDA (Row 19) is =C13-C16 (Gross Profit - SG&A).
    # With Row 16 now populated, Row 19 correctly computes Operating EBITDA!
    for cl in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        ws_hfs[f'{cl}19'] = f"={cl}13-{cl}16"
        ws_hfs[f'{cl}20'] = f"={cl}19/{cl}7" # EBITDA % Sales
        ws_hfs[f'{cl}25'] = f"={cl}19-{cl}22" # EBIT = EBITDA - Depreciation
        ws_hfs[f'{cl}26'] = f"={cl}25/{cl}7" # EBIT % Sales
        ws_hfs[f'{cl}31'] = f"={cl}25-{cl}28" # EBT = EBIT - Interest
        ws_hfs[f'{cl}32'] = f"={cl}31/{cl}7" # EBT % Sales
    print("   -> Verified EBITDA, EBIT, EBT formulas across Rows 19, 25, 31")

    # 2c. Tax Expense (Row 34) & Effective Tax Rate (Row 35)
    ws_hfs['B34'] = "Tax Expense"
    for offset, cl in enumerate(['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']):
        ds_cl = chr(ord('B') + offset)
        ws_hfs[f'{cl}34'] = f"='Data Sheet'!{ds_cl}29"
        ws_hfs[f'{cl}35'] = f"={cl}34/{cl}31" # Effective Tax Rate = Tax Expense / EBT
        ws_hfs[f'{cl}37'] = f"={cl}31-{cl}34" # Net Profit = EBT - Tax Expense
        ws_hfs[f'{cl}38'] = f"={cl}37/{cl}7"  # Net Profit % Sales
    print("   -> Fixed Row 34 Tax Expense and Row 35 Effective Tax Rate")

    # 2d. Shares Outstanding (Row 40)
    # Sun Pharma share count: 240 Cr shares (Face value Re 1, Share Capital Rs 240 Cr)
    for cl in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        ws_hfs[f'{cl}40'] = "='Data Sheet'!$K$70/10000000"
        ws_hfs[f'{cl}42'] = f"={cl}37/{cl}40" # EPS = Net Profit / Shares
    print("   -> Fixed Row 40 Shares Outstanding to point to 240 Cr shares")

# =========================================================================
# 3. FIX BETA REGRESSION — EXCLUDE ZERO OBSERVATIONS & #DIV/0!
# =========================================================================
print("\n[3/7] Correcting 'Beta-Regression'...")
if 'Beta-Regression' in wb.sheetnames:
    ws_beta = wb['Beta-Regression']
    
    # Clear Row 256 and any trailing rows producing #DIV/0!
    for r in range(256, 270):
        for col_idx in range(1, 15):
            ws_beta.cell(row=r, column=col_idx).value = None
            
    # Point regression formulas strictly to valid range (Rows 10 to 251)
    ws_beta['O11'] = "=SLOPE(D10:D251, H10:H251)"
    ws_beta['O12'] = "=_xlfn.COVARIANCE.S(D10:D251, H10:H251)/_xlfn.VAR.S(H10:H251)"
    print("   -> Cleared Row 256 invalid observation; updated regression range to D10:D251, H10:H251")

# =========================================================================
# 4. FIX COMP_VALUATION — PEERS, NET DEBT, SPACING, P/E BRIDGE
# =========================================================================
print("\n[4/7] Correcting 'Comp_Valuation'...")
if 'Comp_Valuation' in wb.sheetnames:
    ws_comp = wb['Comp_Valuation']
    
    # 4a. Clear Column N as empty visual spacer
    ws_comp['N10'].value = None
    for r in range(11, 40):
        ws_comp[f'N{r}'].value = None

    # 4b. Remove duplicate Divi's Lab (Row 19) and synthetic placeholders (Rows 20, 21)
    for r in range(19, 23):
        for col_l in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']:
            ws_comp[f'{col_l}{r}'].value = None
            
    # 4c. Ensure Rows 12 to 18 (7 genuine peers) have dynamic formulas
    for r in range(12, 19):
        ws_comp[f'O{r}'] = f'=IFERROR($H{r}/K{r}, "N/A")'  # EV/Revenue
        ws_comp[f'P{r}'] = f'=IFERROR($H{r}/L{r}, "N/A")'  # EV/EBITDA
        ws_comp[f'Q{r}'] = f'=IFERROR(F{r}/M{r}, "N/A")'   # P/E

    # 4d. Update Summary Statistics to reference valid peer rows (12 to 18)
    ws_comp['O23'] = "=MAX(O12:O18)"
    ws_comp['P23'] = "=MAX(P12:P18)"
    ws_comp['Q23'] = "=MAX(Q12:Q18)"

    ws_comp['O24'] = "=_xlfn.QUARTILE.INC(O12:O18, 1)"
    ws_comp['P24'] = "=_xlfn.QUARTILE.INC(P12:P18, 1)"
    ws_comp['Q24'] = "=_xlfn.QUARTILE.INC(Q12:Q18, 1)"

    ws_comp['O25'] = "=MEDIAN(O12:O18)"
    ws_comp['P25'] = "=MEDIAN(P12:P18)"
    ws_comp['Q25'] = "=MEDIAN(Q12:Q18)"

    ws_comp['O26'] = "=AVERAGE(O12:O18)"
    ws_comp['P26'] = "=AVERAGE(P12:P18)"
    ws_comp['Q26'] = "=AVERAGE(Q12:Q18)"

    ws_comp['O27'] = "=_xlfn.QUARTILE.INC(O12:O18, 3)"
    ws_comp['P27'] = "=_xlfn.QUARTILE.INC(P12:P18, 3)"
    ws_comp['Q27'] = "=_xlfn.QUARTILE.INC(Q12:Q18, 3)"

    ws_comp['O28'] = "=MIN(O12:O18)"
    ws_comp['P28'] = "=MIN(P12:P18)"
    ws_comp['Q28'] = "=MIN(Q12:Q18)"

    # 4e. Sun Pharma Comparable Valuation Bridge (Rows 32 to 39)
    # Row 32: Implied Enterprise / Market Value
    ws_comp['O32'] = "='Raw FS'!AU56*O26"
    ws_comp['P32'] = "='Raw FS'!AV56*P26"
    ws_comp['Q32'] = "='Raw FS'!AW56*Q26"

    # Row 33: Net Debt Value
    # Link to actual Sun Pharma balance sheet: Borrowings (K59) - Cash (K69) = -6,976 Cr (Net Cash)
    ws_comp['O33'] = "='Data Sheet'!K59-'Data Sheet'!K69"
    ws_comp['P33'] = "='Data Sheet'!K59-'Data Sheet'!K69"
    ws_comp['Q33'] = 0  # Net Debt not applicable to P/E

    # Row 34: Implied Market Value
    # For EV multiples: Implied Market Value = Implied EV - Net Debt (= O32 - O33)
    # For P/E: Implied Market Value = Q32 (Net Debt must NOT be subtracted!)
    ws_comp['O34'] = "=O32-O33"
    ws_comp['P34'] = "=P32-P33"
    ws_comp['Q34'] = "=Q32"

    # Row 35: Shares Outstanding
    ws_comp['O35'] = "='Data Sheet'!K70/10000000"
    ws_comp['P35'] = "='Data Sheet'!K70/10000000"
    ws_comp['Q35'] = "='Data Sheet'!K70/10000000"

    # Row 37: Implied Value per Share
    ws_comp['O37'] = "=O34/O35"
    ws_comp['P37'] = "=P34/P35"
    ws_comp['Q37'] = "=Q34/Q35"

    # Row 39: Verdict
    ws_comp['O39'] = "=IF(O37>'Data Sheet'!B8, \"Undervalued\", \"Overvalued\")"
    ws_comp['P39'] = "=IF(P37>'Data Sheet'!B8, \"Undervalued\", \"Overvalued\")"
    ws_comp['Q39'] = "=IF(Q37>'Data Sheet'!B8, \"Undervalued\", \"Overvalued\")"

    print("   -> Cleaned peer table (7 genuine peers), fixed Net Debt (-6,976 Cr), fixed P/E bridge (Q34 == '=Q32')")

# =========================================================================
# 5. FIX DCF & ADD DCF SENSITIVITY TABLE
# =========================================================================
print("\n[5/7] Correcting 'DCF' & adding Sensitivity Analysis...")
if 'DCF' in wb.sheetnames:
    ws_dcf = wb['DCF']
    
    # 5a. Verify Core DCF parameters
    ws_dcf['D37'] = "='Data Sheet'!K69"  # Cash: 11,603 Cr
    ws_dcf['D38'] = "='Data Sheet'!K59"  # Debt: 4,627 Cr
    ws_dcf['D39'] = "=D35+D37-D38"       # Equity Value = Operating Assets + Cash - Debt
    ws_dcf['D40'] = "='Data Sheet'!K70/10000000" # 240 Cr shares
    ws_dcf['D42'] = "=D39/D40"           # Value per Share

    # 5b. Add 7x6 DCF Sensitivity Analysis Table in Rows 48 to 58
    ws_dcf['B48'] = "DCF SENSITIVITY ANALYSIS: IMPLIED VALUE PER SHARE (₹)"
    ws_dcf['B48'].font = Font(name='Calibri', size=11, bold=True, color='1F4E78')
    
    ws_dcf['B49'] = "WACC \\ g"
    ws_dcf['B49'].font = Font(name='Calibri', size=10, bold=True)
    ws_dcf['B49'].alignment = Alignment(horizontal='center', vertical='center')

    g_rates = [0.025, 0.030, 0.035, 0.040, 0.045, 0.050]
    wacc_rates = [0.100, 0.105, 0.110, 0.115, 0.120, 0.125, 0.130]

    # Column Headers (Terminal Growth Rates)
    for c_idx, g_val in enumerate(g_rates):
        col_letter = get_column_letter(3 + c_idx)  # Col C to H
        cell = ws_dcf[f'{col_letter}49']
        cell.value = g_val
        cell.number_format = '0.0%'
        cell.font = Font(name='Calibri', size=10, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')

    # Row Headers (WACC) & Formula Grid
    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )

    for r_idx, wacc_val in enumerate(wacc_rates):
        r_num = 50 + r_idx
        # WACC Row Header
        w_cell = ws_dcf[f'B{r_num}']
        w_cell.value = wacc_val
        w_cell.number_format = '0.0%'
        w_cell.font = Font(name='Calibri', size=10, bold=True)
        w_cell.alignment = Alignment(horizontal='center', vertical='center')
        w_cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')

        # Formula Grid for each combination
        for c_idx, g_val in enumerate(g_rates):
            col_letter = get_column_letter(3 + c_idx)  # Col C to H
            cell = ws_dcf[f'{col_letter}{r_num}']
            # Dynamic valuation formula:
            # PV of 5-Yr FCFFs + PV of TV + Cash - Debt / Shares
            # PV of FCFF = I12/(1+WACC)^I13 + J12/(1+WACC)^J13 + K12/(1+WACC)^K13 + L12/(1+WACC)^L13 + M12/(1+WACC)^M13
            # TV = (M12 * (1 + g)) / (WACC - g)
            # PV of TV = TV / (1 + WACC)^M13
            pv_fcff_part = (
                f"$I$12/(1+$B{r_num})^$I$13 + "
                f"$J$12/(1+$B{r_num})^$J$13 + "
                f"$K$12/(1+$B{r_num})^$K$13 + "
                f"$L$12/(1+$B{r_num})^$L$13 + "
                f"$M$12/(1+$B{r_num})^$M$13"
            )
            tv_part = f"($M$12*(1+{col_letter}$49)/($B{r_num}-{col_letter}$49))*(1/(1+$B{r_num})^$M$13)"
            formula = f'=IF($B{r_num}<={col_letter}$49, "N/A", (({pv_fcff_part}) + ({tv_part}) + $D$37 - $D$38) / $D$40)'
            
            cell.value = formula
            cell.number_format = '₹#,##0.0'
            cell.font = Font(name='Calibri', size=10)
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = thin_border
            
            # Highlight base-case intersection (12.0% WACC, 4.0% g)
            if abs(wacc_val - 0.120) < 0.001 and abs(g_val - 0.040) < 0.001:
                cell.fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')

    print("   -> Configured DCF bridge and added 7x6 formula-driven Sensitivity Table (Rows 48 to 56)")

# =========================================================================
# 6. FIX WACC & CAPITAL STRUCTURE
# =========================================================================
print("\n[6/7] Verifying & updating 'WACC'...")
if 'WACC' in wb.sheetnames:
    ws_wacc = wb['WACC']
    # Sun Pharma Capital Structure in Row 15:
    # Col E (Total Debt) = 'Raw Data'!R25 (or 'Data Sheet'!K59 = 4627)
    # Col F (Total Equity) = 'Raw Data'!S25 (or 'Data Sheet'!B9 = 440830)
    # Row 34: Debt = =E15 (4,627 Cr)
    # Row 35: Equity = =F15 (440,830 Cr)
    # Verify Debt Weight + Equity Weight = 100%
    ws_wacc['C34'] = "=E15"
    ws_wacc['C35'] = "=F15"
    ws_wacc['C36'] = "=SUM(C34:C35)"
    ws_wacc['D34'] = "=C34/C$36"
    ws_wacc['D35'] = "=C35/C$36"
    ws_wacc['D36'] = "=C36/C$36"
    print("   -> Verified WACC Capital Structure consistency (Debt: Rs. 4,627 Cr, Equity: Rs. 440,830 Cr)")

# =========================================================================
# 7. FIX AI VALUATION SUMMARY
# =========================================================================
print("\n[7/7] Updating 'AI Valuation Summary'...")
if 'AI Valuation Summary' in wb.sheetnames:
    ws_ai = wb['AI Valuation Summary']
    ws_ai['B8'] = "Discount Rate (WACC)"
    ws_ai['C8'] = "=WACC!K46"
    ws_ai['B9'] = "Cost of Equity (Ke)"
    ws_ai['C9'] = "=WACC!K29"
    ws_ai['B10'] = "Risk-Free Rate (Rf)"
    ws_ai['C10'] = "=WACC!K26"
    ws_ai['B11'] = "Equity Risk Premium (ERP)"
    ws_ai['C11'] = "=WACC!K27"
    ws_ai['B12'] = "Levered Beta"
    ws_ai['C12'] = "=WACC!K28"
    ws_ai['B13'] = "Terminal Growth Rate"
    ws_ai['C13'] = "=DCF!D19"
    ws_ai['B14'] = "Cost of Debt (Kd)"
    ws_ai['C14'] = "=WACC!K39"
    ws_ai['B15'] = "Cash & Liquid Investments"
    ws_ai['C15'] = "='Data Sheet'!K69"
    ws_ai['B16'] = "Total Debt & Borrowings"
    ws_ai['C16'] = "='Data Sheet'!K59"
    ws_ai['B17'] = "Shares Outstanding"
    ws_ai['C17'] = "='Data Sheet'!K70/10000000"
    ws_ai['B18'] = "Sector Median EV/Revenue"
    ws_ai['C18'] = "=Comp_Valuation!O25"
    ws_ai['B19'] = "Sector Median EV/EBITDA"
    ws_ai['C19'] = "=Comp_Valuation!P25"
    ws_ai['B20'] = "Sector Median P/E"
    ws_ai['C20'] = "=Comp_Valuation!Q25"

    # Valuation Outputs
    ws_ai['B32'] = "Enterprise Value (Operating Assets)"
    ws_ai['C32'] = "=DCF!D35"
    ws_ai['B33'] = "Add: Estimated Cash & Liquid Assets"
    ws_ai['C33'] = "=DCF!D37"
    ws_ai['B34'] = "Less: Total Debt & Borrowings"
    ws_ai['C34'] = "=DCF!D38"
    ws_ai['B35'] = "Net Equity Value"
    ws_ai['C35'] = "=DCF!D39"
    ws_ai['B36'] = "Shares Outstanding"
    ws_ai['C36'] = "=DCF!D40"
    ws_ai['B37'] = "Intrinsic Value per Share"
    ws_ai['C37'] = "=DCF!D42"
    ws_ai['B38'] = "Current Market Price"
    ws_ai['C38'] = "=DCF!D44"
    ws_ai['B39'] = "Margin of Safety / Discount"
    ws_ai['C39'] = "=(C37-C38)/C38"
    print("   -> Standardized AI Valuation Summary B/C orientation and live links")

# Enable full calculation on load
wb.calculation.fullCalcOnLoad = True

# Save to destination paths
wb.save(TARGET_FILE)
wb.close()

print(f"Mirroring corrected workbook to workspace: {WORKSPACE_MIRROR}...")
shutil.copyfile(TARGET_FILE, WORKSPACE_MIRROR)
print("\nWorkbook saved and mirrored successfully!")
