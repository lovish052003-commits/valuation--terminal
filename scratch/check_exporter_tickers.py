import re
import os

with open('excel_exporter.py', 'r', encoding='utf-8') as f:
    text = f.read()

# Find occurrences like screener_data.get('ticker') == '...' or ticker == '...'
matches = re.findall(r'(?:ticker|screener_data\.get\([\'\"]ticker[\'\"]\))\s*==\s*[\'\"]([A-Z0-9_]+)[\'\"]', text)
print('Direct ticker equality checks:', set(matches))

lines = text.split('\n')
for idx, line in enumerate(lines, 1):
    for match in matches:
        if f"== '{match}'" in line or f'== "{match}"' in line:
            print(f"Line {idx}: {line.strip()}")
