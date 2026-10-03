import re

with open('excel_exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

# Check Dupont updates
dupont_matches = [m.start() for m in re.finditer(r'Dupont Analysis', code)]
print(f"Dupont Analysis references in exporter: {len(dupont_matches)}")

# Check Altman updates
altman_matches = [m.start() for m in re.finditer(r"Altman", code)]
print(f"Altman references in exporter: {len(altman_matches)}")

# Check Summary updates
summary_matches = [m.start() for m in re.finditer(r"AI Valuation Summary", code)]
print(f"AI Valuation Summary references in exporter: {len(summary_matches)}")

# Check Data Sheet updates
datasheet_matches = [m.start() for m in re.finditer(r"Data Sheet", code)]
print(f"Data Sheet references in exporter: {len(datasheet_matches)}")
