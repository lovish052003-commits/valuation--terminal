import os
import time

for fn in os.listdir('exports'):
    if 'ADANI' in fn:
        p = os.path.join('exports', fn)
        mtime = os.path.getmtime(p)
        print(f"{fn}: mtime={time.ctime(mtime)}, size={os.path.getsize(p)}")
