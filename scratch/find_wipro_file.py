import os, glob, datetime

for root, dirs, files in os.walk(r"C:\Users\LENOVO\Downloads"):
    for f in files:
        if "wipro" in f.lower() and f.endswith(".xlsx"):
            full = os.path.join(root, f)
            mtime = datetime.datetime.fromtimestamp(os.path.getmtime(full))
            print(f"{full} | Size: {os.path.getsize(full)} | Modified: {mtime}")

for f in glob.glob(r"C:\Users\LENOVO\Desktop\*wipro*.xlsx"):
    mtime = datetime.datetime.fromtimestamp(os.path.getmtime(f))
    print(f"{f} | Size: {os.path.getsize(f)} | Modified: {mtime}")
