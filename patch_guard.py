# patch_guard.py — run from the project folder. Adds a DCF-applicability guard.
# Only changes outputs for companies where the FCFF DCF is already unusable
# (value <= 0, negative EBIT/ROIC notice, financial-institution switch).
import shutil, sys
p = "institutional_control_checks.py"
s = open(p, encoding="utf-8").read()

APPL = 'OR(NOT(ISNUMBER(B5)),B5<=0,LEFT(DCF!B2,6)="NOTICE",LEFT(DCF!B2,9)="FINANCIAL")'

old_e5 = '''safe_set_cell(ws_sum, 'E5', '=IF(NOT(ISNUMBER(C5)),"n/a - check price link",IF(C5>0.15,"UNDERVALUED / BUY",IF(C5<-0.15,"OVERVALUED / SELL","FAIRLY VALUED / HOLD")))')'''
new_e5 = ('''safe_set_cell(ws_sum, 'E5', '=IF(NOT(ISNUMBER(C5)),"n/a - check price link",IF(%s,"NOT APPLICABLE - FCFF DCF unsuitable (use relative valuation)",IF(C5>0.15,"UNDERVALUED / BUY",IF(C5<-0.15,"OVERVALUED / SELL","FAIRLY VALUED / HOLD"))))')''' % APPL)

old_i5 = '''safe_set_cell(ws_sum, 'I5', '=IF(NOT(ISNUMBER(C5)),"n/a",IF(ABS(C5)>0.5,"REVIEW: DCF >50% from price","OK"))')'''
new_i5 = ('''safe_set_cell(ws_sum, 'I5', '=IF(NOT(ISNUMBER(C5)),"n/a",IF(%s,"REVIEW: DCF not applicable to this company",IF(ABS(C5)>0.5,"REVIEW: DCF >50%% from price","OK")))')''' % APPL).replace("%%", "%")

old_adv = '''         "<= 100%",
         '=IF(D21<=1,"PASS","REVIEW")'),
    ]
'''
new_adv = '''         "<= 100%",
         '=IF(D21<=1,"PASS","REVIEW")'),

        (22, "A5. DCF Applicability",
         "DCF value is positive and the FCFF framework suits the company (not negative-EBIT / financial)",
         '=IF(OR(NOT(ISNUMBER(DCF!D43)),DCF!D43<=0,LEFT(DCF!B2,6)="NOTICE",LEFT(DCF!B2,9)="FINANCIAL"),1,0)',
         "0 = applicable",
         '=IF(D22=0,"PASS","REVIEW")'),
    ]
'''
old_fmt = "    return ws\n\n\ndef safe_set_cell("
new_fmt = "    ws['D22'].number_format = \"0\"\n    return ws\n\n\ndef safe_set_cell("

for name, old in (("E5", old_e5), ("I5", old_i5), ("A4 list end", old_adv), ("return ws", old_fmt)):
    n = s.count(old)
    if n != 1:
        sys.exit(f"ABORT: {name} found {n}x, expected 1 (nothing changed)")

shutil.copy(p, p + ".bak2")
s = s.replace(old_e5, new_e5).replace(old_i5, new_i5).replace(old_adv, new_adv).replace(old_fmt, new_fmt)
open(p, "w", encoding="utf-8").write(s)
print("guard patched; backup at", p + ".bak2")
