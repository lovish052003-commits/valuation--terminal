import sqlite3, os, shutil, glob

appdata = os.environ.get('LOCALAPPDATA', '')
history_patterns = [
    os.path.join(appdata, r'Google\Chrome\User Data\*\History'),
    os.path.join(appdata, r'Microsoft\Edge\User Data\*\History'),
    os.path.join(appdata, r'BraveSoftware\Brave-Browser\User Data\*\History')
]

found = []
for pat in history_patterns:
    found.extend(glob.glob(pat))

print("Found history files:", found)
for hpath in found:
    tmp = 'scratch_hist.db'
    try:
        shutil.copy2(hpath, tmp)
        conn = sqlite3.connect(tmp)
        c = conn.cursor()
        query = "SELECT target_path, tab_url, site_url FROM downloads WHERE target_path LIKE '%.xlsx%' ORDER BY start_time DESC LIMIT 15"
        c.execute(query)
        rows = c.fetchall()
        print(f"\n--- From {hpath} ({len(rows)} downloads): ---")
        for r in rows:
            print("Target:", r[0])
            print("Tab URL:", r[1])
            print("Site URL:", r[2])
            print("-" * 40)
        conn.close()
        if os.path.exists(tmp):
            os.remove(tmp)
    except Exception as e:
        print(f"Error on {hpath}: {e}")
        if os.path.exists(tmp):
            os.remove(tmp)
