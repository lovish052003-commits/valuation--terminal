with open('excel_exporter.py', 'r', encoding='utf-8') as f:
    for idx, line in enumerate(f, 1):
        if "'DCF'" in line or '"DCF"' in line or 'populate_dcf' in line.lower():
            print(f"{idx}: {line.strip()[:100]}")
