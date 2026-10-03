import os, sys, glob

downloads_dir = r"C:\Users\LENOVO\Downloads"
files = os.listdir(downloads_dir)

ticker = "NESTLEIND"
company_name = "Nestle India Ltd"
search_tokens = set([ticker.lower(), "nestle"])

matched = []
for f in files:
    if f.endswith('.xlsx'):
        fl = f.lower()
        if 'model' in fl or 'template' in fl or 'itc' in fl or 'test' in fl:
            continue
        for tok in search_tokens:
            if tok in fl:
                matched.append(os.path.join(downloads_dir, f))
                break

print("Matched files in Downloads:", matched)
if matched:
    latest = max(matched, key=os.path.getmtime)
    print("Latest matched file:", latest)
