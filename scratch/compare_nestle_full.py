import openpyxl
import re

print("Loading reference workbook 'Nestle India Model.xlsx'...")
wb_ref = openpyxl.load_workbook('Nestle India Model.xlsx', data_only=True)
wb_ref_f = openpyxl.load_workbook('Nestle India Model.xlsx', data_only=False)

print("Loading generated workbook 'exports/NESTLEIND_Valuation_Model.xlsx'...")
wb_gen = openpyxl.load_workbook('exports/NESTLEIND_Valuation_Model.xlsx', data_only=True)
wb_gen_f = openpyxl.load_workbook('exports/NESTLEIND_Valuation_Model.xlsx', data_only=False)

# Check for formula errors (#VALUE!, #DIV/0!, #REF!, #NAME?, #N/A) in generated workbook
error_patterns = re.compile(r'#VALUE!|#DIV/0!|#REF!|#NAME\?|#N/A|#NUM!')
print("\n=== SCANNING FOR FORMULA ERRORS IN GENERATED WORKBOOK ===")
errors_found = {}
for sheet in wb_gen.sheetnames:
    ws = wb_gen[sheet]
    ws_f = wb_gen_f[sheet]
    errors_in_sheet = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            if val and isinstance(val, str) and error_patterns.search(val):
                formula = ws_f.cell(r, c).value
                coord = ws.cell(r, c).coordinate
                errors_in_sheet.append((coord, val, formula))
    if errors_in_sheet:
        errors_found[sheet] = errors_in_sheet
        print(f"\nSheet '{sheet}' has {len(errors_in_sheet)} errors:")
        for coord, err, f in errors_in_sheet[:10]:
            print(f"  {coord}: error={err} | formula={f}")

print("\n=== SCANNING FOR 'Jan-00' OR 1900 DATES IN GENERATED WORKBOOK ===")
jan00_found = {}
for sheet in wb_gen.sheetnames:
    ws = wb_gen[sheet]
    jan00_in_sheet = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            if val is not None:
                s_val = str(val)
                if '1900-01-00' in s_val or (hasattr(val, 'year') and val.year == 1900):
                    coord = ws.cell(r, c).coordinate
                    jan00_in_sheet.append((coord, val))
    if jan00_in_sheet:
        jan00_found[sheet] = jan00_in_sheet
        print(f"Sheet '{sheet}' has {len(jan00_in_sheet)} Jan-00 dates: {[x[0] for x in jan00_in_sheet[:10]]}")

# Sheet by sheet key valuation comparison
key_sheets = ['DCF', 'WACC', 'Intrinsic Valuation', 'Comp_Valuation', 'Forecasting', 'Dupont Analysis', 'Altman\'s Z Score', 'Data Sheet']
print("\n=== KEY SHEET COMPARISONS (REFERENCE vs GENERATED) ===")
for sheet in key_sheets:
    if sheet not in wb_ref.sheetnames or sheet not in wb_gen.sheetnames:
        print(f"Skipping {sheet}, not in both")
        continue
    ws_r = wb_ref[sheet]
    ws_g = wb_gen[sheet]
    ws_rf = wb_ref_f[sheet]
    ws_gf = wb_gen_f[sheet]
    print(f"\n--- SHEET: {sheet} ---")
    if sheet == 'DCF':
        for coord in ['D34', 'D35', 'D36', 'D37', 'D38', 'D39', 'D40', 'D41', 'D42', 'D43', 'D44', 'E44', 'H6', 'H7', 'H8', 'H9', 'H10', 'H11', 'H12', 'H13', 'H14', 'H15', 'H16', 'H17', 'H18', 'H19', 'H20', 'H21', 'H22', 'H23', 'H24', 'H25', 'H26', 'H27']:
            r_val = ws_r[coord].value
            g_val = ws_g[coord].value
            r_form = ws_rf[coord].value
            g_form = ws_gf[coord].value
            if r_val != g_val or r_form != g_form:
                print(f"  {coord} | REF: val={r_val}, form={r_form}  vs  GEN: val={g_val}, form={g_form}")
    elif sheet == 'WACC':
        for coord in ['D12', 'D13', 'D14', 'D15', 'D16', 'D17', 'D18', 'D20', 'D23', 'D24', 'D25', 'D26', 'D27', 'D28', 'D29', 'D30', 'D34', 'D35', 'D38', 'D39', 'D40']:
            r_val = ws_r[coord].value
            g_val = ws_g[coord].value
            r_form = ws_rf[coord].value
            g_form = ws_gf[coord].value
            if r_val != g_val or r_form != g_form:
                print(f"  {coord} | REF: val={r_val}, form={r_form}  vs  GEN: val={g_val}, form={g_form}")
    elif sheet == 'Intrinsic Valuation':
        for r in range(1, 30):
            r_val = ws_r.cell(r, 2).value
            g_val = ws_g.cell(r, 2).value
            label = ws_r.cell(r, 1).value
            if r_val != g_val:
                print(f"  Row {r} ({label}) | REF: {r_val} vs GEN: {g_val}")
    elif sheet == 'Forecasting':
        print("Checking Forecasting Rows 1 to 30:")
        for r in [4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]:
            r_vals = [ws_r.cell(r, c).value for c in range(2, 8)]
            g_vals = [ws_g.cell(r, c).value for c in range(2, 8)]
            if r_vals != g_vals:
                print(f"  Row {r} | REF: {r_vals} vs GEN: {g_vals}")
    elif sheet == 'Comp_Valuation':
        print("Checking Comp_Valuation Rows 1 to 40:")
        for r in [2, 3, 4, 5, 6, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38]:
            r_c = ws_r.cell(r, 3).value
            g_c = ws_g.cell(r, 3).value
            r_form = ws_rf.cell(r, 3).value
            g_form = ws_gf.cell(r, 3).value
            if r_c != g_c or r_form != g_form:
                print(f"  Row {r} Col C | REF: val={r_c}, form={r_form} vs GEN: val={g_c}, form={g_form}")
