"""
alt_valuation_export.py
Standalone alternate-valuation workbook for companies where the FCFF DCF is not
meaningful (banks / NBFCs / insurers, and highly leveraged conglomerates).

Public API
    is_dcf_applicable(screener_data, dcf_value=None) -> (bool, reason)
    export_alt_model(screener_data, out_dir) -> dict(path, filename, method, value,
                                                     price, upside_pct, verdict, ...)
Self-contained: only needs openpyxl (LibreOffice optional, used to recalc so the
cached values can be read back; otherwise the same maths is done in Python).
"""
import os, re, shutil, subprocess, tempfile, statistics
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="1F3864")
HF = Font(bold=True, color="FFFFFF")

RF, ERP, BETA = 0.07, 0.055, 1.0          # visible, editable inputs on the sheet


def _num(x, default=None):
    try:
        if x is None:
            return default
        if isinstance(x, str):
            x = x.replace(',', '').replace('%', '').strip()
            if x in ('', 'nan', '-'):
                return default
        v = float(x)
        return default if v != v else v
    except Exception:
        return default


def _table_row(sd, table, label, idx=-1):
    tables = sd.get('tables')
    if not isinstance(tables, dict):
        return None
    t = tables.get(table)
    if t is None:
        return None
    if hasattr(t, 'values'):
        for row in t.values:
            if len(row) > 0 and str(row[0]).strip().lower().startswith(label.lower()):
                vals = row[1:]
                if -len(vals) <= idx < len(vals):
                    return _num(vals[idx])
        return None
    if isinstance(t, dict):
        for r in t.get('data', []):
            if r and str(r[0]).strip().lower().startswith(label.lower()):
                vals = r[1:]
                if -len(vals) <= idx < len(vals):
                    return _num(vals[idx])
    return None


def _table_row_series(sd, table, label):
    tables = sd.get('tables')
    if not isinstance(tables, dict):
        return []
    t = tables.get(table)
    if t is None:
        return []
    if hasattr(t, 'values'):
        for row in t.values:
            if len(row) > 0 and str(row[0]).strip().lower().startswith(label.lower()):
                return [_num(v) for v in row[1:]]
        return []
    if isinstance(t, dict):
        for r in t.get('data', []):
            if r and str(r[0]).strip().lower().startswith(label.lower()):
                return [_num(v) for v in r[1:]]
    return []


def extract_inputs(sd):
    """Pull everything the alt model needs from Screener data (latest FY)."""
    price = _num(sd.get('current_price'), 0.0)
    shares = _num(sd.get('shares_in_cr'), 0.0)
    mcap = _num(sd.get('market_cap_cr'), price * shares)
    if not shares and price:
        shares = mcap / price
    eq_cap = _table_row(sd, 'balance-sheet', 'Equity Capital') or 0.0
    reserves = _table_row(sd, 'balance-sheet', 'Reserves') or 0.0
    networth = eq_cap + reserves
    if networth <= 0:
        networth = (_num(sd.get('book_value'), 0.0) or 0.0) * shares
    borrow = (_table_row(sd, 'balance-sheet', 'Borrowing') or 0.0)
    invest = (_table_row(sd, 'balance-sheet', 'Investments') or 0.0)
    pl = (sd.get('tables') or {}).get('profit-loss')
    if hasattr(pl, 'columns'):
        cols = list(pl.columns)[1:]
    elif isinstance(pl, dict):
        cols = pl.get('columns', [])[1:]
    else:
        cols = []
    # latest full fiscal year = last column that is not TTM
    k = len(cols) - 1
    while k > 0 and 'TTM' in str(cols[k]).upper():
        k -= 1
    idx = k - len(cols)
    np_ = _table_row(sd, 'profit-loss', 'Net Profit', idx) or 0.0
    sales = _table_row(sd, 'profit-loss', 'Sales', idx) or _table_row(sd, 'profit-loss', 'Revenue', idx) or 0.0
    opprofit = _table_row(sd, 'profit-loss', 'Operating Profit', idx)
    interest = _table_row(sd, 'profit-loss', 'Interest', idx) or 0.0
    dep = _table_row(sd, 'profit-loss', 'Depreciation', idx) or 0.0
    payout = (_table_row(sd, 'profit-loss', 'Dividend Payout', idx) or 0.0) / 100.0
    # peers
    pes = []
    peers_df = sd.get('peers_df')
    if hasattr(peers_df, 'columns') and hasattr(peers_df, 'values'):
        cols_p = list(peers_df.columns)
        if 'P/E' in cols_p:
            pe_idx = cols_p.index('P/E')
            for r in peers_df.values:
                pe = _num(r[pe_idx])
                if pe and 0 < pe < 200:
                    pes.append(pe)
    elif isinstance(peers_df, dict):
        cols_p = peers_df.get('columns', [])
        if 'P/E' in cols_p:
            pe_idx = cols_p.index('P/E')
            for r in peers_df.get('data', []):
                try:
                    pe = _num(r[pe_idx])
                    if pe and 0 < pe < 200:
                        pes.append(pe)
                except Exception:
                    pass
    evs = [p.get('ev_to_ebitda') for p in (sd.get('sector_peers') or [])
           if isinstance(p, dict) and _num(p.get('ev_to_ebitda')) and 0 < _num(p.get('ev_to_ebitda')) < 60]
    ratio = (interest / sales) if sales else 0.0
    company_type = str(sd.get('company_type') or '').upper()
    sec = str(sd.get('sector') or '').upper()
    is_fin = (ratio > 0.35) or ('BANK' in company_type) or ('FINANC' in sec) or ('NBFC' in company_type)
    return dict(
        name=sd.get('company_name') or sd.get('ticker') or '', ticker=sd.get('ticker') or '',
        price=price, shares=shares, mcap=mcap, networth=networth, np=np_, sales=sales,
        opprofit=opprofit if opprofit is not None else 0.0, interest=interest, dep=dep,
        borrow=borrow, invest=invest, payout=min(max(payout, 0.0), 0.9),
        peer_pe=statistics.median(pes) if pes else None,
        peer_ev=statistics.median(evs) if evs else None,
        fin_ratio=ratio, is_fin=is_fin,
    )


def is_dcf_applicable(sd, dcf_value=None):
    inp = extract_inputs(sd)
    if inp['is_fin']:
        return False, 'Financial company (interest is operating cost) - FCFF DCF not meaningful'
    if dcf_value is not None and (dcf_value is False or _num(dcf_value, 0) <= 0):
        return False, 'DCF produced no positive value'
    if inp['opprofit'] <= 0:
        return False, 'Operating profit not positive'
    return True, ''


def _py_compute(inp, ke, g, lr_roe):
    """Python mirror of the sheet maths (used if LibreOffice recalc is unavailable)."""
    bv = inp['networth'] / inp['shares']
    eps = inp['np'] / inp['shares']
    roe = inp['np'] / inp['networth'] if inp['networth'] else 0
    pay = inp['payout']
    b, pv = bv, 0.0
    for t in range(1, 11):
        r_t = roe + (lr_roe - roe) * (t - 1) / 9
        ri = (r_t - ke) * b
        pv += ri / (1 + ke) ** t
        b = b + r_t * b * (1 - pay)
    # terminal RI (ROE == long run)
    tv = ((lr_roe - ke) * b) * (1 + g) / (ke - g) / (1 + ke) ** 10
    return bv + pv + tv


def build_workbook(inp, path):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Alt Valuation'
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 52
    for c in 'CDEFGHIJKLM':
        ws.column_dimensions[c].width = 15
    ws['B1'] = f"{inp['name']} - Alternate Valuation Model"
    ws['B1'].font = Font(bold=True, size=14)
    ws['B2'] = "Used where the FCFF DCF is not meaningful. Blue = editable input. All other cells are formulas."
    fin = inp['is_fin']

    def row(r, label, val, fmt=None, inp_=False):
        ws.cell(r, 2, label)
        c = ws.cell(r, 3, val)
        if fmt:
            c.number_format = fmt
        if inp_:
            c.font = BLUE
        return c

    ws['B4'] = 'MARKET DATA & LATEST FY'; ws['B4'].font = BOLD
    row(5, 'Current price (Rs)', inp['price'], '#,##0.00', True)
    row(6, 'Shares outstanding (Cr)', inp['shares'], '#,##0.00', True)
    row(7, 'Market cap (Rs Cr)', '=C5*C6', '#,##0')
    row(8, 'Net worth (Rs Cr)', inp['networth'], '#,##0', True)
    row(9, 'Book value per share (Rs)', '=C8/C6', '#,##0.00')
    row(10, 'Net profit (Rs Cr)', inp['np'], '#,##0', True)
    row(11, 'EPS (Rs)', '=C10/C6', '#,##0.00')
    row(12, 'Return on equity', '=C10/C8', '0.0%')
    row(13, 'Dividend payout', inp['payout'], '0%', True)
    row(14, 'Current P/B', '=C5/C9', '0.00"x"')
    row(15, 'Current P/E', '=IFERROR(C5/C11,"n/a")', '0.0"x"')

    ws['B17'] = 'ASSUMPTIONS'; ws['B17'].font = BOLD
    row(18, 'Risk-free rate', RF, '0.0%', True)
    row(19, 'Equity risk premium', ERP, '0.0%', True)
    row(20, 'Beta', BETA, '0.00', True)
    row(21, 'Cost of equity', '=C18+C20*C19', '0.0%')
    row(22, 'Terminal growth', 0.06, '0.0%', True)
    row(23, 'Long-run ROE (fade target)', '=C21', '0.0%')
    ws['D23'] = 'Default = cost of equity (no excess return in perpetuity). Raise it to credit a franchise.'

    ws['B25'] = 'METHOD 1 - RESIDUAL INCOME (primary for banks / NBFCs)'; ws['B25'].font = BOLD
    hdr = ['Year', *[str(t) for t in range(1, 11)]]
    for j, h in enumerate(hdr):
        c = ws.cell(26, 3 + j, h if j else 'Year'); c.font = HF; c.fill = HDR
    ws['B26'] = 'Residual income schedule'; ws['B26'].font = HF; ws['B26'].fill = HDR
    labels = ['Opening book value / share', 'ROE', 'Residual income / share', 'Closing book value / share', 'PV of residual income']
    for i, l in enumerate(labels):
        ws.cell(27 + i, 2, l)
    for t in range(1, 11):
        col = 3 + t  # D..M
        L = ws.cell(26, col).column_letter
        P = ws.cell(26, col - 1).column_letter
        ws.cell(27, col, '=C9' if t == 1 else f'={P}30').number_format = '#,##0.0'
        ws.cell(28, col, f'=$C$12+($C$23-$C$12)*({t}-1)/9').number_format = '0.0%'
        ws.cell(29, col, f'=({L}28-$C$21)*{L}27').number_format = '#,##0.0'
        ws.cell(30, col, f'={L}27+{L}28*{L}27*(1-$C$13)').number_format = '#,##0.0'
        ws.cell(31, col, f'={L}29/(1+$C$21)^{t}').number_format = '#,##0.0'
    row(33, 'Sum of PV of residual income (Rs/share)', '=SUM(D31:M31)', '#,##0.0')
    row(34, 'PV of terminal residual income (Rs/share)', '=(($C$23-$C$21)*M30)*(1+$C$22)/($C$21-$C$22)/(1+$C$21)^10', '#,##0.0')
    row(35, 'Residual-income value per share (Rs)', '=C9+C33+C34', '#,##0.00').font = BOLD
    row(36, 'Justified P/B at long-run ROE', '=(C23-C22)/(C21-C22)', '0.00"x"')
    row(37, 'Market-implied ROE (justified P/B = current P/B)', '=C22+C14*(C21-C22)', '0.0%')

    ws['B39'] = 'METHOD 2 - RELATIVE VALUATION (peer multiples from Screener)'; ws['B39'].font = BOLD
    row(40, 'Peer median P/E', inp['peer_pe'] if inp['peer_pe'] else 'n/a', '0.0"x"', True)
    row(41, 'P/E value per share (Rs)', '=IFERROR(C40*C11,"n/a")', '#,##0.00')
    row(42, 'Target P/B (editable; default = justified P/B)', '=C36', '0.00"x"', True)
    row(43, 'P/B value per share (Rs)', '=C42*C9', '#,##0.00')
    row(44, 'Peer median EV/EBITDA', inp['peer_ev'] if inp['peer_ev'] else 'n/a', '0.0"x"', True)
    row(45, 'EBITDA (Rs Cr)', round(inp['opprofit'], 2), '#,##0', True)
    row(46, 'Borrowings (Rs Cr)', inp['borrow'], '#,##0', True)
    row(47, 'Investments (Rs Cr, treated as cash-like at 90%)', inp['invest'], '#,##0', True)
    row(48, 'EV/EBITDA value per share (Rs)', '=IFERROR((C44*C45-C46+0.9*C47)/C6,"n/a")', '#,##0.00')
    ws['E41'] = '=IF(AND(ISNUMBER(C41),C41>0),C41,"")'
    ws['E48'] = '=IF(AND(ISNUMBER(C48),C48>0),C48,"")'
    ws['D41'] = 'E col: positive values only, used in peer median'

    ws['B50'] = 'METHOD 3 - SUM OF THE PARTS (fill in segments; blank by default)'; ws['B50'].font = BOLD
    for j, h in enumerate(['Segment', 'EBITDA (Rs Cr)', 'EV/EBITDA (x)', 'EV (Rs Cr)']):
        c = ws.cell(51, 2 + j, h); c.font = HF; c.fill = HDR
    for r in range(52, 58):
        ws.cell(r, 3).font = BLUE; ws.cell(r, 4).font = BLUE
        ws.cell(r, 5, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),C{r}*D{r},0)').number_format = '#,##0'
    row(58, 'Sum of segment EV (Rs Cr)', '=SUM(E52:E57)', '#,##0')
    row(59, 'SOTP value per share (Rs)', '=IF(C58>0,(C58-C46+0.9*C47)/C6,"n/a")', '#,##0.00')

    ws['B61'] = 'RESULT'; ws['B61'].font = BOLD
    row(62, 'Method override (blank = auto)', None, None, True)
    ws['D62'] = 'Type: Residual Income / EV-EBITDA / P-E / SOTP'
    row(63, 'Company type', 'Financial (Bank / NBFC)' if fin else 'Corporate / Conglomerate')
    row(64, 'Active method',
        '=IF(C62<>"",C62,IF(C63="Financial (Bank / NBFC)","Residual Income",IF(ISNUMBER(C59),"SOTP","Peer median (EV/EBITDA, P/E)")))')
    row(65, 'Intrinsic value per share (Rs)',
        '=IF(C64="Residual Income",C35,IF(C64="EV/EBITDA",C48,IF(C64="P/E",C41,IF(C64="SOTP",C59,'
        'IFERROR(MEDIAN(E41,E48),C43)))))', '#,##0.00').font = BOLD
    row(66, 'Current price (Rs)', '=C5', '#,##0.00')
    row(67, 'Upside / (downside)', '=C65/C66-1', '0.0%')
    row(68, 'Verdict', '=IF(C67>0.15,"UNDERVALUED",IF(C67<-0.15,"OVERVALUED","FAIRLY VALUED"))')
    row(69, 'Confidence', '=IF(C63="Financial (Bank / NBFC)","Medium - model-driven; sensitive to long-run ROE","Low-Medium - relative; add SOTP segments to improve")')
    # sensitivity: RI value vs Ke and long-run ROE would need data tables; provide closed-form P/B grid
    ws['B71'] = 'SENSITIVITY - justified P/B x BVPS (Rs/share)'; ws['B71'].font = BOLD
    ws['B72'] = 'Long-run ROE  \\  Cost of equity'
    kes = [-0.01, 0.0, 0.01]
    roes = [-0.02, 0.0, 0.03, 0.06]
    for j, d in enumerate(kes):
        c = ws.cell(72, 3 + j, f'=$C$21+({d})'); c.number_format = '0.0%'; c.font = BOLD
    for i, d in enumerate(roes):
        c = ws.cell(73 + i, 2, f'=$C$23+({d})'); c.number_format = '0.0%'; c.font = BOLD
        for j in range(3):
            L = ws.cell(72, 3 + j).column_letter
            ws.cell(73 + i, 3 + j, f'=$C$9*($B{73+i}-$C$22)/({L}$72-$C$22)').number_format = '#,##0'
    ws['B78'] = 'Note: grid shows steady-state justified P/B; the headline value uses the full 10-year fade above.'
    ws.sheet_view.showGridLines = False
    wb.save(path)


def _recalc(path):
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if soffice:
        outd = tempfile.mkdtemp()
        try:
            subprocess.run([soffice, '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', outd, path],
                           capture_output=True, timeout=120)
            out = os.path.join(outd, os.path.basename(path))
            if os.path.exists(out):
                return out
        except Exception:
            pass

    # Windows Excel COM fallback
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        abs_p = os.path.abspath(path)
        wb = excel.Workbooks.Open(abs_p)
        wb.Save()
        wb.Close(SaveChanges=True)
        excel.Quit()
        pythoncom.CoUninitialize()
        return path
    except Exception:
        pass

    return None


def read_result(path):
    """Frontend info is read from the workbook itself."""
    wb = load_workbook(path, data_only=True)
    ws = wb['Alt Valuation']
    v = ws['C65'].value
    if not isinstance(v, (int, float)):
        return None
    return dict(value=float(v), price=float(ws['C66'].value), upside_pct=float(ws['C67'].value) * 100,
                verdict=ws['C68'].value, method=ws['C64'].value, company_type=ws['C63'].value,
                confidence=ws['C69'].value, cost_of_equity=ws['C21'].value,
                residual_income_value=ws['C35'].value, justified_pb=ws['C36'].value,
                implied_roe=ws['C37'].value, pb=ws['C14'].value, roe=ws['C12'].value,
                bvps=ws['C9'].value, eps=ws['C11'].value,
                pe_value=ws['C41'].value, pb_value=ws['C43'].value, ev_ebitda_value=ws['C48'].value)


def export_alt_model(sd, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    inp = extract_inputs(sd)
    ticker = re.sub(r'[^A-Za-z0-9_&-]', '', inp['ticker'] or 'COMPANY').upper()
    fname = f'{ticker}_Alt_Valuation_Model.xlsx'
    final = os.path.join(out_dir, fname)
    tmp = final + '.tmp.xlsx'
    build_workbook(inp, tmp)
    rec = _recalc(tmp)
    res = None
    if rec:
        res = read_result(rec)
        if res and rec != tmp:
            shutil.move(rec, tmp)          # keep cached values so Excel/preview show numbers
    os.replace(tmp, final)
    if not res:  # python mirror for the headline, formulas still live in file
        ke = RF + BETA * ERP
        val = _py_compute(inp, ke, 0.06, ke) if inp['is_fin'] else None
        if val is None:
            cands = []
            if inp['peer_pe'] and inp['np']:
                cands.append(inp['peer_pe'] * inp['np'] / inp['shares'])
            if inp['peer_ev']:
                cands.append((inp['peer_ev'] * inp['opprofit'] - inp['borrow'] + 0.9 * inp['invest']) / inp['shares'])
            val = statistics.median(cands) if cands else 0.0
        up = (val / inp['price'] - 1) * 100 if inp['price'] else 0.0
        bv = inp['networth'] / inp['shares'] if inp['shares'] else 0.0
        roe = inp['np'] / inp['networth'] if inp['networth'] else 0.0
        pb = inp['price'] / bv if bv else 0.0
        res = dict(
            value=round(val, 2), price=inp['price'], upside_pct=round(up, 1),
            verdict='UNDERVALUED' if up > 15 else 'OVERVALUED' if up < -15 else 'FAIRLY VALUED',
            method='Residual Income' if inp['is_fin'] else 'Peer median (EV/EBITDA, P/E)',
            company_type='Financial (Bank / NBFC)' if inp['is_fin'] else 'Corporate / Conglomerate',
            confidence='Medium - model-driven; sensitive to long-run ROE' if inp['is_fin'] else 'Low-Medium - relative; add SOTP segments to improve',
            cost_of_equity=ke, residual_income_value=val,
            justified_pb=(roe - 0.06)/(ke - 0.06) if (ke != 0.06) else 1.0,
            implied_roe=roe, pb=pb, roe=roe, bvps=bv, eps=inp['np']/inp['shares'] if inp['shares'] else 0.0,
            pe_value=(inp['peer_pe'] * inp['np'] / inp['shares']) if (inp['peer_pe'] and inp['shares']) else None,
            pb_value=None,
            ev_ebitda_value=((inp['peer_ev'] * inp['opprofit'] - inp['borrow'] + 0.9 * inp['invest']) / inp['shares']) if (inp['peer_ev'] and inp['shares']) else None
        )
    res.update(path=final, filename=fname, ticker=ticker)
    return res
