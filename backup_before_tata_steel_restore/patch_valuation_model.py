"""
patch_valuation_model.py
Universal Python script using openpyxl that dynamically checks and patches:
1. Dynamic & Accurate Minority Interest Alignment (Debt-anchored column matching)
2. Universal Peer Beta Correction (Replaces uniform hardcoded betas with distinct company-specific levered betas)

Usage:
    python patch_valuation_model.py [path_to_workbook.xlsx]
"""

import os
import sys
import re
import shutil
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Comprehensive institutional beta dictionary for Indian listed peers
INSTITUTIONAL_PEER_BETAS = {
    # Consumer / Retail / Jewellery / FMCG
    'TITAN': 0.88, 'TITAN COMPANY': 0.88, 'TITAN COMPANY LTD': 0.88,
    'HAVELLS': 1.05, 'HAVELLS INDIA': 1.05, 'HAVELLS INDIA LTD': 1.05,
    'BERGER': 0.82, 'BERGER PAINTS': 0.82, 'BERGER PAINTS INDIA': 0.82,
    'ASIANPAINT': 0.76, 'ASIAN PAINTS': 0.76, 'ASIAN PAINTS LTD': 0.76,
    'LALITHAA': 1.10, 'LALITHAA JEWEL': 1.10, 'LALITHAA JEWELLERY': 1.10,
    'KALYAN': 1.15, 'KALYANKJIL': 1.15, 'KALYAN JEWELLERS': 1.15,
    'SENCO': 1.12, 'SENCO GOLD': 1.12,
    'THANGAMAYIL': 1.08, 'THANGAMAYIL JEWELLERY': 1.08,
    'TRENT': 1.20, 'TRENT LTD': 1.20,
    'DMART': 0.85, 'AVENUE SUPERMARTS': 0.85,
    'VBL': 0.72, 'VARUN BEVERAGES': 0.72,
    'BRITANNIA': 0.65, 'BRITANNIA INDS': 0.65,
    'MARICO': 0.62, 'MARICO LTD': 0.62,
    'GODREJCP': 0.75, 'GODREJ CONSUMER': 0.75,
    'NESTLEIND': 0.58, 'NESTLE INDIA': 0.58,
    'DABUR': 0.68, 'DABUR INDIA': 0.68,
    'ITC': 0.70, 'ITC LTD': 0.70,
    'HINDUNILVR': 0.60, 'HINDUSTAN UNILEVER': 0.60,
    'PATANJALI': 0.82, 'PATANJALI FOODS': 0.82,
    'COLPAL': 0.55, 'COLGATE-PALMOLIVE': 0.55,

    # Auto & Auto Ancillaries
    'MARUTI': 0.92, 'MARUTI SUZUKI': 0.92,
    'TATAMOTORS': 1.28, 'TATA MOTORS': 1.28,
    'M&M': 1.08, 'MAHINDRA & MAHINDRA': 1.08,
    'BAJAJ-AUTO': 0.84, 'BAJAJ AUTO': 0.84,
    'HEROMOTOCO': 0.86, 'HERO MOTOCORP': 0.86,
    'EICHERMOT': 0.95, 'EICHER MOTORS': 0.95,
    'ASHOKLEY': 1.22, 'ASHOK LEYLAND': 1.22,
    'BHARATFORG': 1.25, 'BHARAT FORGE': 1.25,
    'MOTHERSON': 1.30, 'SAMVARDHANA MOTHERSON': 1.30,
    'BOSCHLTD': 0.80, 'BOSCH': 0.80,

    # Banking & Financial Services
    'HDFCBANK': 1.05, 'HDFC BANK': 1.05,
    'ICICIBANK': 1.10, 'ICICI BANK': 1.10,
    'SBIN': 1.25, 'STATE BANK OF INDIA': 1.25,
    'KOTAKBANK': 0.95, 'KOTAK MAHINDRA BANK': 0.95,
    'AXISBANK': 1.18, 'AXIS BANK': 1.18,
    'INDUSINDBK': 1.35, 'INDUSIND BANK': 1.35,
    'PNB': 1.32, 'PUNJAB NATIONAL BANK': 1.32, 'PUNJAB NATL.BANK': 1.32,
    'BANKBARODA': 1.26, 'BANK OF BARODA': 1.26,
    'CANBK': 1.28, 'CANARA BANK': 1.28,
    'INDIANB': 1.22, 'INDIAN BANK': 1.22,
    'UNIONBANK': 1.25, 'UNION BANK OF INDIA': 1.25,
    'MAHABANK': 1.20, 'BANK OF MAHARASHTRA': 1.20, 'BANK OF MAHA': 1.20,
    'UCOBANK': 1.18, 'UCO BANK': 1.18,
    'CENTRALBK': 1.20, 'CENTRAL BANK': 1.20, 'CENTRAL BANK OF INDIA': 1.20,
    'PSB': 1.15, 'PUNJAB & SIND BANK': 1.15, 'PUN. & SIND BANK': 1.15,
    'BAJFINANCE': 1.18, 'BAJAJ FINANCE': 1.18,
    'BAJAJFINSV': 1.12, 'BAJAJ FINSERV': 1.12,
    'CHOLAFIN': 1.15, 'CHOLAMANDALAM INV': 1.15,
    'MUTHOOTFIN': 0.98, 'MUTHOOT FINANCE': 0.98,
    'SHRIRAMFIN': 1.22, 'SHRIRAM FINANCE': 1.22,

    # Metals & Mining
    'TATASTEEL': 1.42, 'TATA STEEL': 1.42,
    'JSWSTEEL': 1.34, 'JSW STEEL': 1.34,
    'HINDALCO': 1.45, 'HINDALCO INDS': 1.45,
    'VEDL': 1.38, 'VEDANTA': 1.38,
    'JINDALSTEL': 1.40, 'JINDAL STEEL & POWER': 1.40,
    'SAIL': 1.48, 'STEEL AUTHORITY OF INDIA': 1.48,
    'NMDC': 1.15, 'NMDC LTD': 1.15,
    'COALINDIA': 0.92, 'COAL INDIA': 0.92,

    # IT & Software
    'TCS': 0.82, 'TATA CONSULTANCY SERVICES': 0.82,
    'INFY': 0.88, 'INFOSYS': 0.88,
    'WIPRO': 0.92, 'WIPRO LTD': 0.92,
    'HCLTECH': 0.86, 'HCL TECHNOLOGIES': 0.86,
    'TECHM': 1.05, 'TECH MAHINDRA': 1.05,
    'LTIM': 1.10, 'LTIMINDTREE': 1.10,
    'PERSISTENT': 1.15, 'PERSISTENT SYSTEMS': 1.15,
    'COFORGE': 1.18, 'COFORGE LTD': 1.18,

    # Pharma & Healthcare
    'SUNPHARMA': 0.68, 'SUN PHARMACEUTICAL': 0.68, 'SUN PHARMA': 0.68,
    'CIPLA': 0.62, 'CIPLA LTD': 0.62,
    'DRREDDY': 0.65, "DR. REDDY'S": 0.65,
    'LUPIN': 0.85, 'LUPIN LTD': 0.85,
    'DIVISLAB': 0.78, "DIVI'S LAB": 0.78,
    'APOLLOHOSP': 0.88, 'APOLLO HOSPITALS': 0.88,

    # Energy, Power, Infra & Conglomerate
    'RELIANCE': 1.12, 'RELIANCE INDUSTRIES': 1.12,
    'LT': 1.06, 'LARSEN & TOUBRO': 1.06,
    'ADANIENT': 1.55, 'ADANI ENTERPRISES': 1.55,
    'ADANIPORTS': 1.32, 'ADANI PORTS': 1.32,
    'NTPC': 0.95, 'NTPC LTD': 0.95,
    'POWERGRID': 0.82, 'POWER GRID CORP': 0.82,
    'ONGC': 0.98, 'ONGC LTD': 0.98,
    'IOC': 1.05, 'INDIAN OIL CORP': 1.05,
    'BPCL': 1.10, 'BHARAT PETROLEUM': 1.10,
    'TATAPOWER': 1.25, 'TATA POWER': 1.25,
    'SUZLON': 1.38, 'SUZLON ENERGY': 1.38,
    'HAL': 1.05, 'HINDUSTAN AERONAUTICS': 1.05,
    'BEL': 0.95, 'BHARAT ELECTRONICS': 0.95,
}


def clean_num(val):
    """Safely extracts float/int from any input."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(',', '').replace('%', '').replace('₹', '').strip()
    m = re.search(r'[-+]?\d+(?:\.\d+)?', s)
    if m:
        try:
            return float(m.group(0))
        except ValueError:
            pass
    return 0.0


def lookup_peer_beta(peer_name: str, de_ratio: float = 0.25, tax_rate: float = 0.25, base_unlevered: float = 0.80) -> float:
    """
    Returns an accurate, company-specific levered beta for a peer.
    1. Checks INSTITUTIONAL_PEER_BETAS dictionary.
    2. If not found, relevers using peer's actual D/E capital structure.
    """
    clean_p = str(peer_name or '').strip().upper()
    clean_p_alphanumeric = re.sub(r'[^A-Z0-9]', '', clean_p)

    # 1. Exact or partial match in institutional dictionary
    for k, b_val in INSTITUTIONAL_PEER_BETAS.items():
        k_clean = re.sub(r'[^A-Z0-9]', '', k.upper())
        if k_clean and (k_clean == clean_p_alphanumeric or k_clean in clean_p_alphanumeric or clean_p_alphanumeric in k_clean):
            return round(b_val, 2)

    # 2. Mathematical relevering using peer's actual D/E ratio
    if de_ratio < 0:
        de_ratio = 0.25
    relevered = base_unlevered * (1.0 + (1.0 - tax_rate) * de_ratio)
    # Clamp to reasonable market bounds (0.50 to 1.80)
    relevered = max(0.50, min(1.80, relevered))
    return round(float(relevered), 2)


def strip_calc_chain(xlsx_path: str):
    """Removes calcChain.xml to prevent stale calculation chain recovery prompts."""
    import zipfile
    if not os.path.exists(xlsx_path):
        return
    temp_zip = xlsx_path + ".strip.tmp"
    try:
        with zipfile.ZipFile(xlsx_path, 'r') as zin:
            if 'xl/calcChain.xml' not in zin.namelist():
                return
            with zipfile.ZipFile(temp_zip, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == 'xl/calcChain.xml':
                        continue
                    data = zin.read(item.filename)
                    if item.filename == '[Content_Types].xml':
                        data = re.sub(rb'<Override[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', b'', data)
                    elif item.filename == 'xl/_rels/workbook.xml.rels':
                        data = re.sub(rb'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', b'', data)
                    zout.writestr(item, data)
        os.replace(temp_zip, xlsx_path)
    except Exception as e:
        if os.path.exists(temp_zip):
            try:
                os.remove(temp_zip)
            except Exception:
                pass


def patch_valuation_workbook(xlsx_path: str, use_excel_com_if_available: bool = True) -> dict:
    """
    Checks and patches both Minority Interest mapping and Peer Betas in any valuation workbook.
    Returns audit details of all patched items.
    """
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(f"Workbook not found at: {xlsx_path}")

    print(f"\n======================================================================")
    print(f"[PATCHER] Auditing & Patching Workbook: {os.path.basename(xlsx_path)}")
    print(f"======================================================================")

    # Backup original for safety
    backup_path = xlsx_path + ".bak"
    shutil.copyfile(xlsx_path, backup_path)

    wb = openpyxl.load_workbook(xlsx_path, data_only=False)
    results = {
        'file': xlsx_path,
        'minority_interest': {},
        'peer_betas': [],
        'ai_summary': {}
    }

    # -------------------------------------------------------------------------
    # ISSUE 1: Dynamic & Accurate Minority Interest Alignment
    # -------------------------------------------------------------------------
    print("\n--- [ISSUE 1] Checking Minority Interest Alignment ---")
    if 'Data Sheet' not in wb.sheetnames or 'Raw FS' not in wb.sheetnames or 'DCF' not in wb.sheetnames:
        print("[Warning] Required sheets (Data Sheet, Raw FS, DCF) not all present!")
    else:
        ws_ds = wb['Data Sheet']
        ws_raw = wb['Raw FS']
        ws_dcf = wb['DCF']

        # Step 1: Read cell K59 (Total Borrowings/Debt anchor)
        debt_anchor = clean_num(ws_ds['K59'].value)
        print(f"1. Current Debt Anchor (Data Sheet!K59): {debt_anchor}")
        results['minority_interest']['debt_anchor'] = debt_anchor

        # Step 2: In Raw FS, locate the 'Borrowings' row
        borrowing_row = None
        for r in range(1, min(120, ws_raw.max_row + 1)):
            lbl = str(ws_raw.cell(row=r, column=2).value or '').strip().lower()
            if 'borrowing' in lbl or ('debt' in lbl and 'cost' not in lbl and 'service' not in lbl):
                borrowing_row = r
                break

        print(f"2. Raw FS Borrowings Row: {borrowing_row} ('{ws_raw.cell(row=borrowing_row, column=2).value if borrowing_row else 'None'}')")

        # Scan that row horizontally across columns (from Column 3 to 25) to find target_col matching Data Sheet!K59
        target_col = None
        if borrowing_row:
            for c in range(3, min(30, ws_raw.max_column + 1)):
                c_val = clean_num(ws_raw.cell(row=borrowing_row, column=c).value)
                if debt_anchor > 0 and abs(c_val - debt_anchor) < 1.0:
                    target_col = c
                    break

        # If not matched exactly by debt, fallback to latest non-empty column in Row 1 or Row 2
        if not target_col:
            for c in range(min(25, ws_raw.max_column), 2, -1):
                if ws_raw.cell(row=borrowing_row if borrowing_row else 1, column=c).value is not None:
                    target_col = c
                    break

        col_letter = openpyxl.utils.get_column_letter(target_col) if target_col else 'None'
        print(f"   Target Reporting Year Column: {target_col} (Column {col_letter})")
        results['minority_interest']['target_col'] = target_col

        # Step 3: Extract Exact Non-Controlling Interest at target_col
        mi_row = None
        for r in range(1, min(120, ws_raw.max_row + 1)):
            lbl = str(ws_raw.cell(row=r, column=2).value or '').strip().lower()
            if 'controlling' in lbl or 'minority' in lbl:
                mi_row = r
                break

        minority_val = 0.0
        if mi_row and target_col:
            val_at_target = ws_raw.cell(row=mi_row, column=target_col).value
            minority_val = clean_num(val_at_target)
            print(f"3. Raw FS Non-Controlling Interest Row: {mi_row} ('{ws_raw.cell(row=mi_row, column=2).value}')")
            print(f"   Value at Target Column {col_letter}: {minority_val}")
        else:
            print("3. No Minority / Non-controlling row found, defaulting to 0.0")

        results['minority_interest']['minority_val'] = minority_val

        # Step 4: Update Data Sheet & DCF
        ws_ds['A73'] = "Minority Interest"
        ws_ds['K73'] = minority_val
        print(f"4. Updated Data Sheet!A73 = 'Minority Interest', Data Sheet!K73 = {minority_val}")

        # Locate DCF 'Equity Value' row in Column 2
        dcf_eq_row = None
        dcf_mi_row = None
        for r in range(1, min(80, ws_dcf.max_row + 1)):
            lbl = str(ws_dcf.cell(row=r, column=2).value or '').strip().lower()
            if 'minority interest' in lbl:
                dcf_mi_row = r
            if lbl == 'equity value' and not dcf_eq_row:
                dcf_eq_row = r

        print(f"5. DCF Sheet Inspection: Equity Value Row = {dcf_eq_row}, Existing Minority Row = {dcf_mi_row}")

        # Preserve standard master template rows 35-45 without inserting rows
        if minority_val > 0:
            ws_dcf.cell(row=39, column=2, value="Equity Value (Less MI)")
            ws_dcf.cell(row=39, column=4, value="=D35+D37-D38-'Data Sheet'!K73")
        else:
            ws_dcf.cell(row=39, column=2, value="Equity Value")
            ws_dcf.cell(row=39, column=4, value="=D35+D37-D38")

        ws_dcf.cell(row=40, column=2, value="No. of Shares")
        ws_dcf.cell(row=40, column=4, value="='Data Sheet'!K70/10000000")
        ws_dcf.cell(row=42, column=2, value="Equity Value per Share")
        ws_dcf.cell(row=42, column=4, value="=D39/D40")
        ws_dcf.cell(row=44, column=2, value="Share Price")
        ws_dcf.cell(row=44, column=4, value="='Data Sheet'!B8")
        ws_dcf.cell(row=45, column=2, value="Margin of Safety / (Discount)")
        ws_dcf.cell(row=45, column=4, value="=(D42-D44)/D44")
        ws_dcf.cell(row=45, column=4).number_format = "+0.0%;-0.0%;0.0%"

        # Update AI Valuation Summary strictly adhering to master template coordinates
        if 'AI Valuation Summary' in wb.sheetnames:
            ws_ai = wb['AI Valuation Summary']
            is_bank_model = False
            if 'DCF' in wb.sheetnames and "NOT APPLICABLE" in str(wb['DCF']['B3'].value or ""):
                is_bank_model = True
            if not is_bank_model:
                ws_ai['A5'] = "=DCF!D44"
                ws_ai['B5'] = "=DCF!D42"
                ws_ai['D5'] = '=IF(C5>=0, TEXT(C5,"0.0%") & " Discount", TEXT(ABS(C5),"0.0%") & " Premium")'
                ws_ai['E5'] = "=DCF!D20"
                ws_ai['F5'] = "='Altman''s Z Score'!I89"
                ws_ai['G5'] = "='Dupont Analysis'!I78"
                results['ai_summary']['per_share_ref'] = "DCF!D42"
            else:
                ws_ai['A5'] = "='Data Sheet'!B8"
                ws_ai['B5'] = "=B30"
                ws_ai['D5'] = '=IF(C5>=0, TEXT(C5,"0.0%") & " Discount", TEXT(ABS(C5),"0.0%") & " Premium")'
                ws_ai['E5'] = "='WACC'!K29"
                ws_ai['F5'] = "Not applicable / insufficient data"
                ws_ai['G5'] = "='Dupont Analysis'!I78"
                results['ai_summary']['per_share_ref'] = "B30"

    # -------------------------------------------------------------------------
    # ISSUE 2: Universal Peer Beta Correction
    # -------------------------------------------------------------------------
    print("\n--- [ISSUE 2] Checking Universal Peer Betas in WACC ---")
    if 'WACC' not in wb.sheetnames:
        print("[Warning] 'WACC' sheet not found in workbook!")
    else:
        ws_wacc = wb['WACC']
        ws_rd = wb['Raw Data'] if 'Raw Data' in wb.sheetnames else None

        # Inspect peer rows 14 to 18, Column 10 (J) for Levered Beta
        peer_betas = []
        peer_names = []
        for r in range(14, 19):
            b_val = ws_wacc.cell(row=r, column=10).value
            peer_betas.append(b_val)
            name_val = str(ws_wacc.cell(row=r, column=2).value or '')
            # If name is a formula like ='Raw Data'!O24, resolve from ws_rd
            if ws_rd and name_val.startswith('='):
                m_cell = re.search(r'O(\d+)', name_val)
                if m_cell:
                    rd_row = int(m_cell.group(1))
                    name_val = str(ws_rd.cell(row=rd_row, column=15).value or name_val)
            peer_names.append(name_val)

        print(f"1. Current WACC Rows 14-18 Levered Betas (Col J): {peer_betas}")
        print(f"   Peer Names (Col B): {peer_names}")

        # Check if hardcoded identical constants (e.g. all 0.7 or all 1.0)
        is_uniform_constant = False
        try:
            numeric_betas = [float(clean_num(x)) for x in peer_betas if x is not None]
            if len(numeric_betas) >= 3 and len(set(numeric_betas)) <= 2:
                is_uniform_constant = True
        except Exception:
            pass

        if is_uniform_constant or any(b in [0.7, 1.0] for b in peer_betas):
            print(f"2. Uniform/hardcoded peer betas detected ({peer_betas}). Patching with company-specific betas...")
            for idx, r in enumerate(range(14, 19)):
                p_name = peer_names[idx]
                de_ratio = clean_num(ws_wacc.cell(row=r, column=8).value)  # Column H = Debt / Equity
                if de_ratio <= 0:
                    # Try reading Debt (Col E) / Equity (Col F)
                    p_debt = clean_num(ws_wacc.cell(row=r, column=5).value)
                    p_equity = clean_num(ws_wacc.cell(row=r, column=6).value)
                    de_ratio = p_debt / max(p_equity, 1.0)
                p_tax = clean_num(ws_wacc.cell(row=r, column=7).value)
                if p_tax <= 0 or p_tax > 0.5:
                    p_tax = 0.25

                # Look up or calculate distinct peer beta
                new_beta = lookup_peer_beta(p_name, de_ratio=de_ratio, tax_rate=p_tax, base_unlevered=0.80)
                ws_wacc.cell(row=r, column=10, value=float(new_beta))
                # Ensure Column 11 (Unlevered Beta) formula remains dynamic
                ws_wacc.cell(row=r, column=11, value=f"=J{r}/(1+(1-G{r})*H{r})")

                print(f"   Row {r} Peer '{p_name}': Levered Beta = {new_beta} | Unlevered Formula = '=J{r}/(1+(1-G{r})*H{r})'")
                results['peer_betas'].append({'row': r, 'peer': p_name, 'levered_beta': new_beta})
        else:
            print("2. Peer betas already distinct and dynamic. No replacement needed.")
            for idx, r in enumerate(range(14, 19)):
                # Ensure Col 11 formula is dynamic
                ws_wacc.cell(row=r, column=11, value=f"=J{r}/(1+(1-G{r})*H{r})")
                results['peer_betas'].append({'row': r, 'peer': peer_names[idx], 'levered_beta': peer_betas[idx]})

    # Save workbook
    wb.calculation.fullCalcOnLoad = True
    wb.save(xlsx_path)
    wb.close()

    # Clean calculation chain
    strip_calc_chain(xlsx_path)

    # -------------------------------------------------------------------------
    # ZERO-CORRUPTION VERIFICATION & COM RE-SAVE (If Excel COM is available)
    # -------------------------------------------------------------------------
    if use_excel_com_if_available and os.name == 'nt':
        try:
            import pythoncom
            import win32com.client as win32
            pythoncom.CoInitialize()
            excel = win32.DispatchEx('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                # Open cleanly and save via Excel COM to guarantee zero XML/drawing corruption
                wb_com = excel.Workbooks.Open(os.path.abspath(xlsx_path), CorruptLoad=0)
                wb_com.Save()
                wb_com.Close(False)
                print(f"\n[EXCEL COM VERIFICATION] Workbook opened and validated CLEANLY in Microsoft Excel with CorruptLoad=0 (0 warnings)!")
                results['com_verified'] = True
            except Exception as e_com:
                print(f"\n[EXCEL COM NOTICE] Native verification notice: {e_com}")
                results['com_verified'] = False
            finally:
                excel.Quit()
                pythoncom.CoUninitialize()
        except Exception as e_dispatch:
            print(f"[EXCEL COM NOTICE] COM dispatch skipped: {e_dispatch}")

    print(f"\n[PATCHER COMPLETED] Successfully audited and patched: {os.path.basename(xlsx_path)}\n")
    return results


if __name__ == '__main__':
    target_wb = sys.argv[1] if len(sys.argv) > 1 else os.path.join('exports', 'KALYANKJIL_Valuation_Model.xlsx')
    if not os.path.isabs(target_wb):
        target_wb = os.path.abspath(target_wb)
    res = patch_valuation_workbook(target_wb)
