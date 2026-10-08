# patch_labels.py — run from the project folder, after patch_h51.py
import shutil, sys
p = "institutional_control_checks.py"
s = open(p, encoding="utf-8").read()
reps = [
 ('"5 Col Median (H:L)"', '"4 Col Median (I:L)"', 2),
 ("=MAX('Intrinsic Valuation'!H52:L52)", "=MAX('Intrinsic Valuation'!I52:L52)", 1),
 ("Currently used in DCF (L55 - includes FY22 blank-year and FY25 acquisition values)",
  "Currently used in DCF (L55 - median of I:L; acquisition-year spikes can distort it)", 1),
]
for old, new, n in reps:
    if s.count(old) != n:
        sys.exit(f"ABORT: {old[:40]!r} found {s.count(old)}x, expected {n}")
shutil.copy(p, p + ".bak")
for old, new, n in reps:
    s = s.replace(old, new)
open(p, "w", encoding="utf-8").write(s)
print("ok")
