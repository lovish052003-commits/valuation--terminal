# patch_h51.py (v2) — run from the project folder
import re, shutil, sys
p = "excel_exporter.py"
src = open(p, encoding="utf-8").read()

rx_cols = re.compile(r"cols_seq\s*=\s*\['G',\s*'H',\s*'I',\s*'J',\s*'K',\s*'L'\]")
rx_55   = re.compile(r'"=MEDIAN\(H52:L52\)"')
rx_65   = re.compile(r'"=MEDIAN\(H62:L62\)"')

for name, rx in (("cols_seq", rx_cols), ("L55 median", rx_55), ("L65 median", rx_65)):
    n = len(rx.findall(src))
    if n != 2:
        sys.exit(f"ABORT: expected exactly 2 matches for {name}, found {n}")

shutil.copy(p, p + ".bak")
src = rx_cols.sub("cols_seq = ['H', 'I', 'J', 'K', 'L']", src)
src = rx_55.sub('"=MEDIAN(I52:L52)"', src)
src = rx_65.sub('"=MEDIAN(I62:L62)"', src)
open(p, "w", encoding="utf-8").write(src)
print("patched both export paths; backup at", p + ".bak")
