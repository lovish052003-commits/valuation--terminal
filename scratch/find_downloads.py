import glob, os, datetime

downloads_dir = r"C:\Users\LENOVO\Downloads"
for pat in ["*WIPRO*.xlsx", "*Valuation_Model*.xlsx"]:
    matches = glob.glob(os.path.join(downloads_dir, pat))
    for m in matches:
        mtime = datetime.datetime.fromtimestamp(os.path.getmtime(m))
        size = os.path.getsize(m)
        print(f"{m} | Size: {size} | Modified: {mtime}")
