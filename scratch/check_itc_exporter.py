import os

with open('excel_exporter.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

print(f"Total lines in excel_exporter.py: {len(lines)}")
itc_lines = []
for i, line in enumerate(lines, 1):
    l_strip = line.strip()
    if l_strip.startswith('#'):
        continue
    if "'ITC'" in line or '"ITC"' in line or "ITC Ltd" in line or "ITC Limited" in line:
        itc_lines.append((i, l_strip))

print(f"Total non-comment code lines mentioning ITC: {len(itc_lines)}")
for idx, l in itc_lines:
    print(f"  Line {idx:4d}: {l[:100]}")
