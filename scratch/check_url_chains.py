import sqlite3, os, shutil

hpath = os.path.expanduser(r'~\AppData\Local\Google\Chrome\User Data\Default\History')
tmp = 'scratch_hist_chains.db'
shutil.copy2(hpath, tmp)
conn = sqlite3.connect(tmp)
c = conn.cursor()

c.execute("""
    SELECT d.id, d.target_path, c.url
    FROM downloads d
    JOIN downloads_url_chains c ON d.id = c.id
    WHERE d.target_path LIKE '%Nestle%' OR d.target_path LIKE '%BRITANNIA%' OR d.target_path LIKE '%ITC%'
    ORDER BY d.start_time DESC LIMIT 20
""")
for r in c.fetchall():
    print(f"ID: {r[0]} | Path: {r[1]} | URL: {r[2]}")

conn.close()
os.remove(tmp)
