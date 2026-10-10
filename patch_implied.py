# patch_implied.py — run from the project folder. Purely additive: two new info rows on the DCF sheet.
# Shows what the market price implies, so a low DCF value is explained rather than just reported.
import shutil, sys
p = "institutional_control_checks.py"
s = open(p, encoding="utf-8").read()

old = '''        safe_set_cell(ws_dcf, 'D83', "=D77/D79", number_format=FMT_PCT)
'''
new = '''        safe_set_cell(ws_dcf, 'D83', "=D77/D79", number_format=FMT_PCT)

        # Market-implied view (info only - feeds nothing)
        safe_set_cell(ws_dcf, 'B84', "Market-implied exit EV/EBIT multiple (Year 10)")
        safe_set_cell(ws_dcf, 'D84', "=IFERROR(($D$45*$D$41+$D$38+$D$39-$D$37-$D$78)/($M$68*(1+$D$20)^(-10)),\\"n/a\\")", number_format='0.0"x"')
        safe_set_cell(ws_dcf, 'B85', "EV / latest EBIT at current price")
        safe_set_cell(ws_dcf, 'D85', "=IFERROR(($D$45*$D$41+$D$38+$D$39-$D$37)/$H$8,\\"n/a\\")", number_format='0.0"x"')
'''
if s.count(old) != 1:
    sys.exit(f"ABORT: anchor found {s.count(old)}x, expected 1 (nothing changed)")
shutil.copy(p, p + ".bak4")
open(p, "w", encoding="utf-8").write(s.replace(old, new))
print("implied-multiple rows added; backup at", p + ".bak4")
