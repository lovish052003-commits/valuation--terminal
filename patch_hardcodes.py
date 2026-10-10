# patch_hardcodes.py — run from the project folder.
# 1) excel_exporter.py: removes the 3 UNILEVER cash overrides, the UNILEVER debt/share overrides,
#    and makes live balance-sheet cash take priority over the static cash table
#    (the static table holds ITC's INVESTMENTS as "cash", which the DCF bridge then counted twice).
# 2) valuation_engine.py: removes the HUL-specific cash rule.
# All edits are exact-match; if anything is not found the script changes NOTHING.
import re, shutil, sys

EXP, ENG = "excel_exporter.py", "valuation_engine.py"
exp = open(EXP, encoding="utf-8").read()
eng = open(ENG, encoding="utf-8").read()

# ---- 1a) cash block (appears 3x: COM path, openpyxl populate, patch_valuation_workbook) ----------
rx_cash = re.compile(
    r"(?P<ind>^[ ]+)verified_cash_cr = clean_num\(\s*"
    r"screener_data\.get\('bank_cash'\) or\s*screener_data\.get\('cash_cr'\) or\s*"
    r"screener_data\.get\('cash_and_equivalents_cr'\)\s*\)\s*"
    r"if 'UNILEVER' in str\(screener_data\.get\('company_name', ''\)\)\.upper\(\):\s*"
    r"verified_cash_cr = 1580\.0\s*"
    r"elif verified_cash_cr <= 0 and cash_dict:\s*"
    r"latest_cash_key = list\(cash_dict\.keys\(\)\)\[-1\]\s*"
    r"verified_cash_cr = clean_num\(cash_dict\[latest_cash_key\]\)",
    re.M)
def cash_repl(m):
    i = m.group('ind')
    return (f"{i}# Live balance-sheet cash first; static/aggregated fields only as fallback\n"
            f"{i}verified_cash_cr = 0.0\n"
            f"{i}if cash_dict:\n"
            f"{i}    latest_cash_key = list(cash_dict.keys())[-1]\n"
            f"{i}    verified_cash_cr = clean_num(cash_dict[latest_cash_key])\n"
            f"{i}if verified_cash_cr <= 0:\n"
            f"{i}    verified_cash_cr = clean_num(\n"
            f"{i}        screener_data.get('bank_cash') or\n"
            f"{i}        screener_data.get('cash_cr') or\n"
            f"{i}        screener_data.get('cash_and_equivalents_cr')\n"
            f"{i}    )")
n_cash = len(rx_cash.findall(exp))
if n_cash != 3:
    sys.exit(f"ABORT: cash block found {n_cash}x in {EXP}, expected 3 (nothing changed)")

# ---- 1b) UNILEVER debt + share overrides (COM path only) -----------------------------------------
old_debt = ("    if verified_debt_cr == 0 and 'UNILEVER' in str(screener_data.get('company_name', '')).upper():\n"
            "        verified_debt_cr = 1478.0\n")
old_shares = ("    if 'UNILEVER' in str(screener_data.get('company_name', '')).upper() and "
              "(verified_shares_cr < 230 or verified_shares_cr > 240):\n"
              "        verified_shares_cr = 234.96\n")
for name, old in (("UNILEVER debt", old_debt), ("UNILEVER shares", old_shares)):
    if exp.count(old) != 1:
        sys.exit(f"ABORT: {name} override found {exp.count(old)}x in {EXP}, expected 1 (nothing changed)")

# ---- 2) valuation_engine HUL rule ----------------------------------------------------------------
old_eng = ("        if latest_investments > 0 and verified_cash >= 5000 and "
           "abs(verified_cash - (1580.0 + latest_investments)) < 100:\n"
           "            cash_estimate = verified_cash\n"
           "        else:\n"
           "            cash_estimate = verified_cash + latest_investments\n")
new_eng = "        cash_estimate = verified_cash + latest_investments\n"
if eng.count(old_eng) != 1 and new_eng not in eng:
    sys.exit(f"ABORT: HUL cash rule found {eng.count(old_eng)}x in {ENG}, expected 1 (nothing changed)")

# ---- all checks passed: write ---------------------------------------------------------------------
shutil.copy(EXP, EXP + ".bak5"); shutil.copy(ENG, ENG + ".bak5")
exp = rx_cash.sub(cash_repl, exp)
exp = exp.replace(old_debt, "").replace(old_shares, "")
if old_eng in eng:
    eng = eng.replace(old_eng, new_eng)
open(EXP, "w", encoding="utf-8").write(exp)
open(ENG, "w", encoding="utf-8").write(eng)
print("hardcodes removed; backups: excel_exporter.py.bak5, valuation_engine.py.bak5")
